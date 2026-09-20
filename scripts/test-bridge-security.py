#!/usr/bin/env python3
"""Регрессионные тесты безопасности и устойчивости моста.

Проверяет ровно те дефекты, которые были найдены и исправлены:
  * обход каталога при отдаче статики (`/assets/../../bridge/bridge.py` читал исходники,
    а `…/state.db` — всю историю сессий Hermes);
  * инъекция команд в POST /open (session_id попадал в строку `cmd.exe /c start …`);
  * CSRF: POST с чужого origin (`Origin: null`, страница-песочница) запускал агента;
  * падение запроса до отправки ответа на мусорном `shape`/`context`;
  * /stop без указания цели гасил ВСЕ прогоны, включая чужие.

Запуск (мост должен быть поднят на своём обычном порту):
    python scripts/test-bridge-security.py
"""
from __future__ import annotations

import json
import os
import ssl
import subprocess
import sys
import time
import http.client
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HOST = os.environ.get("HERMES_BRIDGE_HOST", "localhost")
PORT = int(os.environ.get("HERMES_BRIDGE_PORT", "3443"))
BASE = f"https://{HOST}:{PORT}"
PROOF = Path(os.environ.get("TEMP", "/tmp")) / "hermes_inj_proof.txt"

_CTX = ssl.create_default_context()
_CTX.check_hostname = False
_CTX.verify_mode = ssl.CERT_NONE

results: list[tuple[bool, str]] = []


def check(ok: bool, title: str, detail: str = "") -> None:
    results.append((ok, title))
    print(f"  {'PASS' if ok else 'FAIL'}  {title}" + (f"  — {detail}" if detail else ""))


def request(path: str, method: str = "GET", body: dict | None = None,
            headers: dict | None = None, path_as_is: bool = False) -> tuple[int, bytes]:
    url = BASE + path
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    if data:
        req.add_header("Content-Type", "application/json")
    opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=_CTX))
    try:
        with opener.open(req, timeout=45) as resp:
            return resp.status, resp.read(4000)
    except urllib.error.HTTPError as e:
        return e.code, e.read(4000)
    except Exception as e:                                     # noqa: BLE001
        return 0, str(e).encode()


def stream_head(payload: dict, seconds: float = 25.0) -> tuple[str, dict]:
    """Открывает SSE-поток, читает только первые кадры и закрывает соединение.

    Ждать завершения прогона нельзя: тест станет медленным и флейковым. Возвращает
    заголовок статуса и разобранное событие start (в нём run_id для адресного стопа).
    """
    conn = http.client.HTTPSConnection(HOST, PORT, context=_CTX, timeout=seconds)
    try:
        conn.request("POST", "/chat", body=json.dumps(payload).encode("utf-8"),
                     headers={"Content-Type": "application/json", "X-Hermes-Bridge": "1"})
        resp = conn.getresponse()
        if resp.status != 200:
            return f"HTTP {resp.status}", {}
        text = ""                                              # читаем кадр до пустой строки,
        while "\n\n" not in text:                              # а не фиксированные N байт:
            raw = resp.readline()                              # иначе чтение ждёт дельты модели
            if not raw:
                break
            text += raw.decode("utf-8", "replace")
    except Exception as e:                                     # noqa: BLE001
        return f"ошибка соединения: {e}", {}
    finally:
        try:
            conn.close()
        except Exception:                                      # noqa: BLE001
            pass
    start: dict = {}
    if "event: start" in text:
        try:
            start = json.loads(text.split("data:", 1)[1].split("\n")[0])
        except Exception:                                      # noqa: BLE001
            start = {}
    return "SSE " + text[:40].replace("\n", " "), start


def main() -> int:
    print(f"Мост: {BASE}\n")

    print("1. Обход каталога при отдаче статики")
    for evil in ("/assets/../../bridge/bridge.py", "/assets/../../../.git/config",
                 "/assets/../../../../../AppData/Local/hermes/state.db"):
        code, body = request(evil)
        leaked = b"import" in body[:400] or body[:15] == b"SQLite format 3"
        check(code in (403, 404) and not leaked, f"{evil} → {code}", f"{len(body)} байт утечки" if leaked else "")

    print("\n2. Инъекция команд в POST /open")
    PROOF.unlink(missing_ok=True)
    code, body = request("/open", "POST",
                         {"session_id": f'zz & echo PWNED>"{PROOF}" & echo x'},
                         {"X-Hermes-Bridge": "1"})
    time.sleep(1.5)
    check(not PROOF.exists(), "session_id с метасимволами отклонён", body.decode("utf-8", "replace")[:80])
    check(code == 200 and b"false" in body, f"ответ не выполняется командой (HTTP {code})")
    PROOF.unlink(missing_ok=True)

    print("\n3. CSRF: POST с чужого origin")
    code, _ = request("/chat", "POST", {"prompt": "x"}, {"Origin": "https://evil.example"})
    check(code == 403, f"чужой Origin → {code}")
    code, _ = request("/chat", "POST", {"prompt": "x"}, {"Origin": "null"})
    check(code == 403, f"Origin: null без заголовка → {code}")
    code, _ = request("/chat", "POST", {"prompt": "x"}, {"X-Hermes-Bridge": "1", "Origin": "null"})
    check(code == 403, f"Origin: null с заголовком → {code}")
    code, _ = request("/stop", "POST", {}, {"Origin": "https://evil.example"})
    check(code == 403, f"чужой Origin на /stop → {code}")

    print("\n4. Мусорные типы в теле запроса не роняют соединение")
    head, start = stream_head({"prompt": "тест устойчивости", "shape": [1, 2, 3], "context": "мусор",
                               "max_turns": "не число"})
    check(head.startswith("SSE") and start.get("run_id"),
          f"shape=[1,2,3] и context=строка → {head}", str(start.get("run_id", "")))

    print("\n4b. Стоп бьёт только по указанному прогону (run_id)")
    if start.get("run_id"):
        code, body = request("/stop", "POST", {"run_id": start["run_id"]}, {"X-Hermes-Bridge": "1"})
        try:
            killed = json.loads(body).get("killed")
        except Exception:                                      # noqa: BLE001
            killed = None
        check(code == 200 and killed == 1, f"/stop по run_id → killed={killed}")
    code, body = request("/stop", "POST", {"session_id": "20260101_000000_ffffff"},
                         {"X-Hermes-Bridge": "1"})
    try:
        killed = json.loads(body).get("killed")
    except Exception:                                          # noqa: BLE001
        killed = None
    check(code == 200 and killed == 0, f"стоп по чужой сессии не убивает ничего → killed={killed}")

    print("\n5. /stop без цели ничего не убивает")
    code, body = request("/stop", "POST", {}, {"X-Hermes-Bridge": "1"})
    try:
        payload = json.loads(body)
    except Exception:                                          # noqa: BLE001
        payload = {}
    check(code == 200 and payload.get("killed") == 0, f"пустой /stop → {body.decode('utf-8', 'replace')[:60]}")

    print("\n6. Загрузка вложений")
    import base64
    png = base64.b64encode(b"\x89PNG\r\n\x1a\n" + b"0" * 32).decode()
    code, body = request("/upload", "POST", {"name": "../побег.png", "kind": "image", "data": png},
                         {"X-Hermes-Bridge": "1"})
    try:
        up = json.loads(body)
    except Exception:                                          # noqa: BLE001
        up = {}
    name = str(up.get("name") or "")
    check(code == 200 and up.get("ok") and ".." not in name and "/" not in name and "\\" not in name,
          f"имя санитизировано: {name!r}")
    check(str(up.get("path") or "").replace("\\", "/").endswith("/pane/uploads/" + name),
          "файл лёг только в pane/uploads")
    code, body = request("/upload", "POST", {"name": "x.png", "kind": "image", "data": "!!!не base64!!!"},
                         {"X-Hermes-Bridge": "1"})
    check(code == 200 and b'"ok": false' in body.replace(b'"ok":false', b'"ok": false'),
          "мусорный base64 отклонён без падения")
    code, body = request("/chat", "POST", {"prompt": "x", "attachments": [
        {"path": "C:/Windows/System32/drivers/etc/hosts", "name": "hosts", "kind": "file"}]},
        {"X-Hermes-Bridge": "1"})
    check(code == 200, "путь вне uploads не ломает ход (вложение отброшено)")

    bad = [t for ok, t in results if not ok]
    print(f"\nитог: {len(results) - len(bad)}/{len(results)} проверок пройдено"
          + (f"; провалены: {bad}" if bad else ""))
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
