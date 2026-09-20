#!/usr/bin/env python3
"""Проверка вложений: загрузка файлов в мост и передача их агенту.

Запуск: python scripts/test-attachments.py ["текст задачи"]
Проверяет: /upload принимает UTF-8 JSON и кладёт файл в workspace/pane/uploads,
пути вне uploads отклоняются, а ход с вложениями доходит до агента (в запросе есть блок вложений).
"""
from __future__ import annotations

import base64
import json
import os
import ssl
import sys
import time
import http.client
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HOST, PORT = "localhost", int(os.environ.get("HERMES_BRIDGE_PORT", "3443"))
UPLOADS = ROOT / "workspace" / "pane" / "uploads"
CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE
ok = True


def check(name: str, cond: bool, detail: str = "") -> None:
    global ok
    ok = ok and bool(cond)
    print(("  PASS  " if cond else "  FAIL  ") + name + (f" — {detail}" if detail else ""))


def post(path: str, payload: dict, timeout: int = 60) -> tuple[int, dict]:
    conn = http.client.HTTPSConnection(HOST, PORT, context=CTX, timeout=timeout)
    conn.request("POST", path, body=json.dumps(payload).encode("utf-8"),
                 headers={"Content-Type": "application/json", "X-Hermes-Bridge": "1"})
    r = conn.getresponse()
    raw = r.read()
    try:
        data = json.loads(raw.decode("utf-8"))
    except Exception:                                          # noqa: BLE001
        data = {"raw": raw[:200].decode("utf-8", "replace")}
    conn.close()
    return r.status, data


def main() -> int:
    prompt = sys.argv[1] if len(sys.argv) > 1 else "Что в приложенных файлах? Ответь одной строкой."
    print(f"мост: https://{HOST}:{PORT}\n")

    print("1. Загрузка файлов")
    png = (ROOT / "addin" / "assets" / "icon-64.png").read_bytes()
    csv = "позиция,номинал,ток\nNXB-63 1P C10,10 А,10\nNXB-63 3P C25,25 А,25\n".encode("utf-8")
    st1, d1 = post("/upload", {"name": "фото-щита.png", "kind": "image",
                               "data": base64.b64encode(png).decode()})
    st2, d2 = post("/upload", {"name": "спецификация автоматов.csv", "kind": "file",
                               "data": base64.b64encode(csv).decode()})
    check("картинка принята", st1 == 200 and d1.get("ok"), str(d1)[:90])
    check("документ принят", st2 == 200 and d2.get("ok"), str(d2)[:90])
    for d in (d1, d2):
        p = Path(str(d.get("path") or ""))
        check(f"файл на диске: {p.name}", p.is_file() and p.parent == UPLOADS,
              f"{p.stat().st_size} байт" if p.is_file() else "нет файла")

    print("\n2. Защита: путь вне uploads не принимается")
    st3, d3 = post("/chat", {"prompt": "прочитай файл", "attachments": [
        {"path": str(ROOT / "bridge" / "bridge.py"), "name": "bridge.py", "kind": "file"}]}, timeout=120)
    check("чужой путь отброшен (ход идёт без вложений)", st3 == 200, f"HTTP {st3}")

    print("\n3. Ход с вложениями доходит до агента")
    body = {"prompt": prompt, "model": os.environ.get("HERMES_BENCH_MODEL", "deepseek-flash"),
            "provider": "deepseek", "attachments": [d1, d2]}
    conn = http.client.HTTPSConnection(HOST, PORT, context=CTX, timeout=300)
    conn.request("POST", "/chat", body=json.dumps(body).encode("utf-8"),
                 headers={"Content-Type": "application/json", "X-Hermes-Bridge": "1"})
    r = conn.getresponse()
    tools, answer, err = 0, "", ""
    for raw in r:
        line = raw.decode("utf-8", "replace").strip()
        if line.startswith("event:"):
            kind = line.split(":", 1)[1].strip()
        elif line.startswith("data:"):
            try:
                data = json.loads(line[5:])
            except Exception:                                  # noqa: BLE001
                continue
            if kind == "tool":
                tools += 1
                print(f"     инструмент #{tools}: {data.get('name')}")
            elif kind == "result":
                answer = data.get("text", "")
            elif kind == "error":
                err = str(data)[:200]
            elif kind == "done":
                break
    conn.close()
    check("агент получил ход", not err, err)
    check("агент открывал вложения инструментами", tools >= 1, f"кругов: {tools}")
    print("\n  ответ:", (answer or "(пусто)")[:400].replace("\n", " "))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
