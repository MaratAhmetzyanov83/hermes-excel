#!/usr/bin/env python3
"""
Hermes Excel Bridge — локальный HTTPS-мост между надстройкой Excel и агентом Hermes.

Что делает:
  * отдаёт статику надстройки (taskpane.html/js/css, иконки) по https://localhost:3443/
  * POST /chat   — принимает промпт + контекст листа, запускает `hermes chat --format stream-json`
                   и стримит события агента (текст, инструменты, результат) обратно в Excel по SSE
  * GET  /models — список провайдеров и моделей (из кэша моделей Hermes) для выпадающего списка
  * GET  /sessions — последние сессии Hermes (из state.db), чтобы продолжить любую из Excel
  * GET  /session?id=<sid> — последние реплики сессии: панель восстанавливает ленту после
                    перезапуска Excel, а не начинает с пустого чата
  * POST /open   — открыть ту же сессию в обычном Hermes (CLI), т.е. «продолжить в Hermes»
  * POST /stop   — прервать текущий прогон

Сессии, созданные из Excel, лежат в общем сторе Hermes ($HERMES_HOME/state.db),
поэтому их видно и в desktop-приложении, и через `hermes --resume <id>`.

Зависимости: только стандартная библиотека Python 3.9+.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sqlite3
import ssl
import subprocess
import sys
import threading
import time
from urllib.parse import unquote
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

# ---------------------------------------------------------------- пути/настройки

ROOT = Path(__file__).resolve().parent.parent          # .../hermes-excel
ADDIN_DIR = ROOT / "addin"
WORKSPACE = ROOT / "workspace"
CERT_DIR = Path.home() / ".office-addin-dev-certs"

HOST = os.environ.get("HERMES_BRIDGE_HOST", "127.0.0.1")
PORT = int(os.environ.get("HERMES_BRIDGE_PORT", "3443"))
ORIGIN = f"https://localhost:{PORT}"

HERMES = os.environ.get("HERMES_BIN") or shutil.which("hermes") or "hermes"

# Бот для Excel = отдельный профиль Hermes: свой конфиг, память, скиллы и СВОИ сессии.
# Пустой PROFILE или "default" — работаем в основном профиле.
PROFILE = (os.environ.get("HERMES_BRIDGE_PROFILE") or "excel").strip()
_HOME = Path(os.environ.get("HERMES_HOME") or (Path.home() / ".hermes"))
HERMES_HOME = _HOME if PROFILE in ("", "default") else (_HOME / "profiles" / PROFILE)
STATE_DB = HERMES_HOME / "state.db"
MODELS_CACHE = HERMES_HOME / "provider_models_cache.json"
# у свежего бота кэша моделей может не быть — берём его из основного профиля (только чтение)
FALLBACK_MODELS_CACHE = _HOME / "provider_models_cache.json"

MAX_CONTEXT_ROWS = 400          # сколько строк выделения уходит модели
MAX_CONTEXT_COLS = 60
# Идентификатор сессии Hermes: 20260919_144821_ea810f. Валидируем перед подстановкой куда-либо.
SESSION_ID_RE = re.compile(r"[A-Za-z0-9_-]{6,64}")
# Панель отвечает текстом и блоками ```csv/```formulas: браузер, картинки, cron и делегирование
# ей не нужны. Узкий набор вдвое уменьшает промпт (схемы инструментов) и убирает круги
# «попробовал инструмент — политика запретила». Расширяется через HERMES_BRIDGE_TOOLSETS.
PANE_TOOLSETS = os.environ.get("HERMES_BRIDGE_TOOLSETS", "file")
PANE_SKILLS = [s for s in os.environ.get("HERMES_BRIDGE_SKILLS", "excel-pane-contract").split(",") if s]
# Бюджет прогона: на 80 % модель получает предупреждение и успевает свернуть работу.
RUN_BUDGET = float(os.environ.get("HERMES_BRIDGE_RUN_BUDGET", "240"))
RUN_TIMEOUT = float(os.environ.get("HERMES_BRIDGE_TIMEOUT", "1800"))

WORKSPACE.mkdir(parents=True, exist_ok=True)
# Файлы запросов и рабочая папка прогонов — ВНЕ workspace: агент запускается с --in и видит файлы
# вокруг себя, поэтому брошенные `.query-*.txt` прошлых ходов он принимал за текущую задачу
# (наблюдалось: агент 60 с рассуждал о чужом промпте вместо своего).
QUERY_DIR = Path(os.environ.get("TEMP") or os.environ.get("TMP") or str(WORKSPACE)) / "hermes-bridge"
try:
    QUERY_DIR.mkdir(parents=True, exist_ok=True)
    for stale in QUERY_DIR.glob("*.txt"):
        stale.unlink(missing_ok=True)
except OSError:
    QUERY_DIR = WORKSPACE
PANE_DIR = WORKSPACE / "pane"
PANE_DIR.mkdir(parents=True, exist_ok=True)
for stale in WORKSPACE.glob(".query-*.txt"):                  # уборка с прошлых версий моста
    stale.unlink(missing_ok=True)

_runs: dict[str, dict] = {}          # run_id -> {"proc": Popen, "session": str}
_pending_stops: dict[str, float] = {}  # цель стопа, пришедшая раньше регистрации прогона
_runs_lock = threading.Lock()
_cfg_cache: dict[str, object] = {"at": 0.0, "default": {}}


def log(*a: object) -> None:
    """Пишет и в консоль, и в файл workspace/bridge.log.

    Файл важен, когда мост поднимает сторож или ярлык автозапуска: окна нет, и без файла
    диагностировать нечего (логи проверки в AGENTS.md читают именно workspace/bridge.log).
    """
    line = " ".join(str(x) for x in a)
    stamped = f"{time.strftime('[%H:%M:%S]')} {line}"
    print(stamped, flush=True)
    try:
        with (WORKSPACE / "bridge.log").open("a", encoding="utf-8") as f:
            f.write(stamped + "\n")
    except OSError:
        pass


# ---------------------------------------------------------------- вспомогательное

def hermes_default() -> dict[str, str]:
    """Модель/провайдер по умолчанию из config.yaml (кэш 60 с)."""
    if time.time() - float(_cfg_cache["at"]) < 60 and _cfg_cache["default"]:
        return _cfg_cache["default"]  # type: ignore[return-value]
    out = {"model": "", "provider": ""}
    try:
        cmd = [HERMES]
        if PROFILE and PROFILE != "default":
            cmd += ["-p", PROFILE]
        cmd += ["config", "get", "model"]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=90, cwd=str(WORKSPACE))
        for line in (proc.stdout or "").splitlines():
            m = re.match(r"\s*(\w+):\s*(.+?)\s*$", line)
            if not m:
                continue
            key, val = m.group(1), m.group(2)
            if key == "default":
                out["model"] = val
            elif key == "provider":
                out["provider"] = val
    except Exception as exc:                                  # noqa: BLE001
        log("config get model failed:", exc)
    _cfg_cache.update(at=time.time(), default=out)
    return out


def model_catalog() -> dict:
    """Провайдеры и модели из кэша моделей бота (+ запасной кэш основного профиля)."""
    providers: list[dict] = []
    seen: set[str] = set()
    for cache in (MODELS_CACHE, FALLBACK_MODELS_CACHE):
        try:
            raw = json.loads(cache.read_text(encoding="utf-8"))
        except Exception:                                    # noqa: BLE001
            continue
        for name, entry in raw.items():
            models = entry.get("models") if isinstance(entry, dict) else None
            if not models:
                continue
            clean = name.split("#")[0]
            if clean.startswith("custom:"):
                clean = "custom"
            if clean in seen:
                continue
            seen.add(clean)
            providers.append({"provider": clean, "models": list(models)})
    providers.sort(key=lambda p: p["provider"])
    return {"default": hermes_default(), "profile": PROFILE, "providers": providers}


def session_messages(session_id: str, limit: int = 12) -> list[dict]:
    """Последние реплики сессии — чтобы панель восстановила ленту чата после перезапуска Excel.

    Только чтение; отдаём лишь роль и текст (служебные tool-сообщения пропускаем).
    """
    if not STATE_DB.exists() or not SESSION_ID_RE.fullmatch(session_id or ""):
        return []
    try:
        con = sqlite3.connect(f"file:{STATE_DB}?mode=ro", uri=True, timeout=5)
        con.row_factory = sqlite3.Row
        rows = con.execute(
            """SELECT role, content, display_kind
                 FROM messages
                WHERE session_id = ? AND role IN ('user','assistant')
                  AND COALESCE(active,1)=1
                  AND content IS NOT NULL AND TRIM(content) <> ''
                ORDER BY id DESC
                LIMIT ?""",
            (session_id, int(limit)),
        ).fetchall()
    except Exception as exc:                                  # noqa: BLE001
        log("session_messages query failed:", exc)
        return []
    out: list[dict] = []
    for r in reversed(rows):
        text = str(r["content"] or "")
        # В панель не нужно тащить служебные обёртки и служебные роли
        if r["display_kind"] in ("system", "tool_call"):
            continue
        out.append({"role": r["role"], "text": text[:4000]})
    return out


def list_sessions(limit: int = 40) -> list[dict]:
    """Последние сессии из общего стора Hermes (только чтение)."""
    if not STATE_DB.exists():
        return []
    try:
        con = sqlite3.connect(f"file:{STATE_DB}?mode=ro", uri=True, timeout=5)
        con.row_factory = sqlite3.Row
        rows = con.execute(
            """SELECT id, title, model, source, message_count, started_at,
                      last_activity_at, estimated_cost_usd
                 FROM sessions
                WHERE COALESCE(archived,0)=0 AND COALESCE(hidden,0)=0
                ORDER BY COALESCE(last_activity_at, started_at) DESC
                LIMIT ?""",
            (int(limit),),
        ).fetchall()
        con.close()
        return [dict(r) for r in rows]
    except Exception as exc:                                  # noqa: BLE001
        log("sessions query failed:", exc)
        return []


def error_hint(text: str) -> str:
    """Что сделать пользователю по тексту сбоя. Пусто — подсказки нет."""
    low = (text or "").lower()
    if any(w in low for w in ("sign-in", "sign in", "oauth", "rejected your", "auth failed", "401", "403")):
        return f"hermes -p {PROFILE} auth add <провайдер> --type oauth  (или выберите другую модель в панели)"
    if "rate limit" in low or "429" in low or "quota" in low:
        return "исчерпана квота или лимит — подождите минуту либо переключите модель в панели"
    if any(w in low for w in ("too long", "context length", "maximum context")):
        return "выделите меньший диапазон и повторите"
    if any(w in low for w in ("timeout", "timed out", "budget")):
        return "прогон прерван по времени — разбейте задачу на части"
    if any(w in low for w in ("not found", "не найден")):
        return f"проверьте провайдера и модель: hermes -p {PROFILE} config get model"
    return ""


def build_query(prompt: str, context: dict | None, file_name: str, sheet: str,
                rng: str, shape: tuple[int, int] | None) -> str:
    """Собирает итоговый промпт: сначала контекст листа, затем задачу пользователя."""
    parts: list[str] = []
    empty = bool(context and context.get("empty"))
    if context and context.get("csv"):
        rows, cols = shape if shape else (0, 0)
        head = (f"[Контекст Excel] Файл: {file_name or 'неизвестно'} | Лист: {sheet or '—'} | "
                f"Диапазон: {rng or '—'} | {rows} строк x {cols} колонок | CSV с шапкой:")
        parts.append(head)
        parts.append(context["csv"])
        note = context.get("note")
        if note:
            parts.append(f"({note})")
        parts.append(
            "Если результат нужно положить в книгу — верни его отдельным блоком ```csv (с шапкой, "
            "столбцы через запятую) либо блоком ```formulas с формулами Excel по одной в строке. "
            "Перед блоками кратко поясни вывод. Формулы пиши на английском (SUM, SUMIFS, XLOOKUP)."
        )
        parts.append("Файл открыт в Excel прямо сейчас, данные выше — это и есть его содержимое: "
                     "искать книгу на диске и проверять файлы не нужно.")
    elif empty:
        parts.append(
            f"[Контекст Excel] Файл: {file_name or 'неизвестно'} | Лист: {sheet or '—'} | "
            f"Диапазон: {rng or '—'} | выделение ПУСТОЕ, данных нет."
        )
        parts.append(
            "Пользователь просит подготовить содержимое, чтобы записать его в этот диапазон (или рядом с ним). "
            "Не задавай уточняющих вопросов, если можно сделать разумное предположение — сделай работу. "
            "Ответь одним коротким предложением, что именно подготовил, и обязательно верни готовый результат "
            "ОТДЕЛЬНЫМ блоком ```csv (первая строка — шапка, разделитель — запятая) либо блоком ```formulas "
            "с формулами по одной в строке. Никакого текста после блока."
        )
    parts.append(prompt or "Проанализируй выделенный диапазон.")
    return "\n\n".join(parts)



# ---------------------------------------------------------------- HTTP

class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "HermesExcelBridge/1.0"

    # ---- утилиты ответа
    def _cors(self) -> None:
        self.send_header("Access-Control-Allow-Origin", ORIGIN)
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Hermes-Bridge")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")

    def _json(self, obj: object, code: int = 200) -> None:
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def _origin_ok(self) -> bool:
        """POST принимаем только от панели (или CLI).

        Два независимых условия: (1) обязательный кастомный заголовок — его нельзя поставить «простым»
        запросом из браузера без preflight, поэтому чужая страница мост не дёрнет; (2) если Origin
        пришёл — он должен быть нашим. `null` (песочница <iframe sandbox>, data:, file:) запрещён:
        раньше он пускался, и любая веб-страница могла запускать агента на этой машине.
        """
        if self.headers.get("X-Hermes-Bridge") != "1":
            return False
        origin = self.headers.get("Origin")
        if origin is None:
            return True
        return origin in (ORIGIN, ORIGIN.replace("localhost", "127.0.0.1"))

    def _drain(self) -> None:
        """Вычитать тело отклонённого запроса: иначе в keep-alive соединении оно станет
        началом следующего запроса и сервер ответит 400 на корректный GET."""
        try:
            length = int(self.headers.get("Content-Length") or 0)
            if length > 0:
                self.rfile.read(length)
        except Exception:                                     # noqa: BLE001
            pass

    def _body(self) -> dict:
        """Тело запроса. Сбой разбора НЕ молчит: раньше `except: return {}` превращал кривое тело
        в пустой запрос (панель присылала данные, а агент отвечал «данных нет» — и концов не найти)."""
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except (TypeError, ValueError):
            log("body: непонятный Content-Length:", self.headers.get("Content-Length"))
            return {}
        if not length:
            return {}
        try:
            raw = self.rfile.read(length)
        except Exception as exc:                              # noqa: BLE001
            log("body: чтение не удалось:", exc)
            return {}
        try:
            return json.loads(raw.decode("utf-8") or "{}")
        except Exception as exc:                              # noqa: BLE001
            log(f"body: разбор JSON не удался ({exc}); первые 120 байт: {raw[:120]!r}")
            self._body_failed = True
            return {}

    @staticmethod
    def _safe_static(name: str) -> Path | None:
        """Отдаём только файлы из addin/: ни `..`, ни абсолютных путей, ни выхода по symlink.

        Раньше `ADDIN_DIR / name` резолвился по файловой системе, и `GET /assets/../../bridge/bridge.py`
        отдавал исходники, а `GET /assets/../../../AppData/Local/hermes/state.db` — всю историю сессий.
        """
        if not name or name.startswith("/") or "\\\\" in name or ".." in name.split("/"):
            return None
        try:
            candidate = (ADDIN_DIR / name).resolve()
            candidate.relative_to(ADDIN_DIR.resolve())
        except (OSError, ValueError):
            return None
        return candidate

    # ---- GET
    def do_GET(self) -> None:                                 # noqa: N802
        path = self.path.split("?")[0]
        if path == "/health":
            self._json({"ok": True, "hermes": HERMES, "profile": PROFILE,
                        "profile_home": str(HERMES_HOME), "workspace": str(WORKSPACE),
                        "state_db": str(STATE_DB), "origin": ORIGIN})
            return

        if path == "/models":
            self._json(model_catalog())
            return
        if path == "/sessions":
            q = self.path.split("limit=")[-1] if "limit=" in self.path else "40"
            try:
                limit = max(1, min(200, int(q)))
            except ValueError:
                limit = 40
            self._json({"sessions": list_sessions(limit)})
            return
        if path == "/session":
            sid = ""
            if "id=" in self.path:
                sid = unquote(self.path.split("id=", 1)[1].split("&")[0])
            if not SESSION_ID_RE.fullmatch(sid):
                self._json({"error": "нужен параметр id=<session_id>"}, 400)
                return
            self._json({"session_id": sid, "messages": session_messages(sid)})
            return
        if path in ("/", "/index.html", "/taskpane.html"):
            self._static(ADDIN_DIR / "taskpane.html", "text/html; charset=utf-8")
            return
        if path.startswith("/assets/") or path in ("/taskpane.js", "/taskpane.css", "/commands.html",
                                                   "/manifest.xml", "/test-office.html", "/mock-office.js"):
            name = path.lstrip("/")
            if name == "manifest.xml":
                name = "manifest.xml"
            target = self._safe_static(name)
            if target is None:
                log("static: отказано в доступе", path)
                self._json({"error": "forbidden", "path": path}, 403)
                return
            ctype = {".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
                     ".png": "image/png", ".html": "text/html; charset=utf-8",
                     ".xml": "application/xml; charset=utf-8"}.get(
                Path(name).suffix, "application/octet-stream")
            self._static(target, ctype)
            return
        self._json({"error": "not found", "path": path}, 404)

    def _static(self, path: Path, ctype: str) -> None:
        try:
            data = path.read_bytes()
        except OSError:
            self._json({"error": f"missing {path.name}"}, 404)
            return
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self._cors()
        self.end_headers()
        self.wfile.write(data)

    # ---- POST
    def do_OPTIONS(self) -> None:                             # noqa: N802
        self.send_response(204)
        self._cors()
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_POST(self) -> None:                                # noqa: N802
        if not self._origin_ok():
            self._drain()
            self._json({"error": "forbidden: нужен заголовок X-Hermes-Bridge: 1 и свой Origin"}, 403)
            return
        path = self.path.split("?")[0]
        if path == "/chat":
            self._chat()
            return
        if path == "/stop":
            body = self._body()
            self._stop(str(body.get("session_id") or ""), str(body.get("run_id") or ""))
            return
        if path == "/open":
            body = self._body()
            self._json(self._open_session(str(body.get("session_id") or "")))
            return
        self._drain()
        self._json({"error": "not found", "path": path}, 404)

    # ---- /chat: SSE-стрим ответа агента
    def _chat(self) -> None:
        body = self._body()
        if getattr(self, "_body_failed", False):
            # Тело не разобралось (например, клиент прислал кириллицу не в UTF-8). Раньше мост молча
            # стартовал прогон с пустым промптом, и агент отвечал «данных нет» — вместо ошибки.
            self._json({"error": "тело запроса не разобралось как UTF-8 JSON", "hint":
                        "кодируйте тело в UTF-8 (curl: --data-binary @file.json)"}, 400)
            return
        prompt = str(body.get("prompt") or "")
        # Мусорные типы не должны ронять запрос до отправки ответа: раньше shape:[1,2,3]
        # давал ValueError «too many values to unpack» и клиент получал оборванное соединение.
        context = body.get("context")
        if not isinstance(context, dict):
            context = {}
        shape_raw = body.get("shape")
        shape = shape_raw if isinstance(shape_raw, list) and len(shape_raw) == 2 else None
        session_id = str(body.get("session_id") or "").strip()
        model = str(body.get("model") or "").strip()
        provider = str(body.get("provider") or "").strip()
        reasoning = str(body.get("reasoning") or "").strip()
        try:
            max_turns = str(max(1, min(200, int(body.get("max_turns") or 40))))
        except (TypeError, ValueError):
            max_turns = "40"
        query = build_query(prompt, context, str(body.get("file_name") or ""),
                            str(body.get("sheet") or ""), str(body.get("range") or ""), shape)

        run_id = uuid.uuid4().hex[:8]
        qfile = QUERY_DIR / f"query-{run_id}.txt"
        qfile.write_text(query, encoding="utf-8")
        log(f"chat {run_id}: csv={'да' if context.get('csv') else 'нет'} "
            f"empty={'да' if context.get('empty') else 'нет'} session={session_id or 'новый'} "
            f"промпт={prompt[:60]!r} файл={qfile}")

        cmd = [HERMES]
        if PROFILE and PROFILE != "default":
            cmd += ["-p", PROFILE]
        cmd += ["chat", "-Q", "--format", "stream-json", "--query-file", str(qfile),
                "--max-turns", max_turns, "--in", str(PANE_DIR)]
        if PANE_TOOLSETS:
            cmd += ["-t", PANE_TOOLSETS]
        for skill in PANE_SKILLS:
            cmd += ["-s", skill]
        if RUN_BUDGET > 0:
            cmd += ["--run-budget", str(int(RUN_BUDGET))]

        if session_id and session_id != "new":
            cmd += ["--resume", session_id]
        if model and model != "auto":
            cmd += ["-m", model]
        # Провайдеры вида "custom:http://..." — это динамически добавленные OpenAI-совместимые
        # эндпоинты, их нельзя передать в --provider по имени: пусть Hermes маршрутизирует сам.
        if provider and provider != "auto" and not provider.startswith("custom"):
            cmd += ["--provider", provider]

        if reasoning and reasoning != "auto":
            cmd += ["--reasoning", reasoning]

        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache, no-transform")
        self.send_header("Connection", "keep-alive")
        self.send_header("X-Accel-Buffering", "no")
        self.send_header("Transfer-Encoding", "chunked")
        self._cors()
        self.end_headers()

        def sse(event: str, data: object) -> None:
            payload = f"event: {event}\ndata: {json.dumps(data, ensure_ascii=False)}\n\n".encode("utf-8")
            self.wfile.write(b"%x\r\n" % len(payload) + payload + b"\r\n")
            self.wfile.flush()

        proc = None
        sid = session_id
        watchdog = None
        try:
            try:
                proc = subprocess.Popen(
                    cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, cwd=str(PANE_DIR),
                    text=True, encoding="utf-8", errors="replace", bufsize=1,
                    env={**os.environ, "PYTHONIOENCODING": "utf-8", "NO_COLOR": "1"},
                    creationflags=getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
                )
            except FileNotFoundError:
                sse("error", {"message": f"не найден исполняемый файл hermes ({HERMES})"})
                sse("done", {"session_id": "", "exit_code": 127})
                return

            # Регистрируем прогон сразу после старта процесса: стоп, нажатый до прихода
            # события start, должен найти цель, а не промахнуться.
            with _runs_lock:
                _runs[run_id] = {"proc": proc, "session": session_id}
                pending = run_id in _pending_stops or (session_id and session_id in _pending_stops)
                if pending:
                    _pending_stops.pop(run_id, None)
                    _pending_stops.pop(session_id, None)
            if pending:
                self._kill(proc)
                sse("error", {"message": "прогон остановлен до старта"})
                return

            # Таймаут не может зависеть от вывода процесса: молчащий прогон (задумавшаяся модель,
            # зависший инструмент) держал панель до получаса. Сторож убивает его сам.
            watchdog = threading.Timer(RUN_TIMEOUT, self._on_timeout, args=(proc, run_id))
            watchdog.daemon = True
            watchdog.start()

            sse("start", {"run_id": run_id, "session_id": sid, "model": model or "auto",
                          "provider": provider or "auto", "timeout_s": int(RUN_TIMEOUT)})

            assert proc.stdout is not None
            for line in proc.stdout:
                line = line.strip()
                if not line:
                    continue
                try:
                    evt = json.loads(line)
                except json.JSONDecodeError:
                    sse("log", {"text": line})
                    continue
                kind = evt.get("type")
                if kind == "system" and evt.get("session_id"):
                    sid = str(evt["session_id"])
                    sse("session", {"session_id": sid, "model": evt.get("model", "")})
                elif kind == "text":
                    sse("delta", {"text": evt.get("text", "")})
                elif kind == "tool" or kind == "tool_use" or kind == "tool_call":
                    sse("tool", {"name": evt.get("name") or evt.get("tool") or "tool",
                                 "detail": evt.get("detail") or evt.get("input") or ""})
                elif kind == "result":
                    sid = str(evt.get("session_id") or sid)
                    code = int(evt.get("exit_code") or 0)
                    if code != 0:
                        # Служебный сбой — не ответ ассистента: панель покажет его как ошибку
                        # с готовым действием, а не как текст модели (раньше так выглядела,
                        # например, смерть OAuth-токена провайдера).
                        text = str(evt.get("text") or "")
                        sse("error", {"message": text or f"прогон завершился с кодом {code}",
                                      "exit_code": code, "session_id": sid,
                                      "hint": error_hint(text)})
                    else:
                        sse("result", {"session_id": sid, "text": evt.get("text", ""),
                                       "tokens": evt.get("tokens", {}),
                                       "duration_ms": evt.get("duration_ms", 0),
                                       "exit_code": 0})
                else:
                    sse("log", {"text": json.dumps(evt, ensure_ascii=False)[:400]})

            code = proc.wait(timeout=60)
            hint = ""
            if sid:
                flag = f"-p {PROFILE} " if PROFILE and PROFILE != "default" else ""
                hint = f"hermes {flag}chat --resume {sid}"
            sse("done", {"session_id": sid, "exit_code": code, "resume_hint": hint})
        except (BrokenPipeError, ConnectionResetError):
            if proc and proc.poll() is None:
                self._kill(proc)
        except Exception as exc:                              # noqa: BLE001
            try:
                sse("error", {"message": str(exc)})
                sse("done", {"session_id": sid, "exit_code": 1})
            except Exception:                                 # noqa: BLE001
                pass
        finally:
            if watchdog is not None:
                watchdog.cancel()
            with _runs_lock:
                _runs.pop(run_id, None)
            qfile.unlink(missing_ok=True)
            try:
                self._end_chunked()
            except Exception:                                 # noqa: BLE001
                pass

    def _end_chunked(self) -> None:
        """Терминатор chunked-потока ровно один раз: второй «0\\r\\n\\r\\n» оставался в keep-alive
        соединении и разбирался как мусор в начале следующего запроса."""
        if getattr(self, "_chunk_done", False):
            return
        self._chunk_done = True
        self.wfile.write(b"0\r\n\r\n")
        self.wfile.flush()

    # ---- прочее
    @staticmethod
    def _kill(proc: subprocess.Popen) -> None:
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                           capture_output=True, timeout=30)
        except Exception:                                     # noqa: BLE001
            try:
                proc.kill()
            except Exception:                                 # noqa: BLE001
                pass

    def _on_timeout(self, proc: subprocess.Popen, run_id: str) -> None:
        """Сторож прогона: убивает процесс (и его детей) по истечении RUN_TIMEOUT."""
        try:
            if proc.poll() is None:
                log(f"таймаут прогона {run_id} ({int(RUN_TIMEOUT)} с) — останавливаю процесс")
                self._kill(proc)
        except Exception:                                     # noqa: BLE001
            pass

    def _stop(self, session_id: str, run_id: str) -> None:
        """Останавливает ТОЛЬКО указанный прогон.

        Раньше при переданном session_id без run_id условие убивало все записи в _runs, то есть
        чужой прогон в другом окне Excel. Пустое тело не убивает ничего.
        """
        if not run_id and not session_id:
            self._json({"ok": True, "killed": 0, "note": "нужен run_id или session_id"})
            return
        killed = 0
        with _runs_lock:
            for key, rec in list(_runs.items()):
                if run_id and key != run_id:
                    continue
                if not run_id and session_id and rec.get("session") != session_id:
                    continue
                proc = rec.get("proc")
                if proc is not None and proc.poll() is None:
                    self._kill(proc)
                    killed += 1
            if not killed:                                    # прогон ещё не зарегистрировался
                _pending_stops[run_id or session_id] = time.time()
                for stale, ts in list(_pending_stops.items()):
                    if time.time() - ts > 300:
                        _pending_stops.pop(stale, None)
        self._json({"ok": True, "killed": killed})

    def _open_session(self, session_id: str) -> dict:
        """Открыть ту же сессию бота в обычном Hermes (новое окно консоли).

        Без shell: раньше строка вида `start "…" cmd /k "…" --resume {session_id}` собиралась
        с shell=True, и session_id попадал в неё как есть — «zz & calc.exe» исполнялось.
        """
        if not SESSION_ID_RE.fullmatch(session_id or ""):
            return {"ok": False, "error": "некорректный session_id"}
        cmd = [HERMES]
        if PROFILE and PROFILE != "default":
            cmd += ["-p", PROFILE]
        cmd += ["chat", "--resume", session_id]
        try:
            subprocess.Popen(cmd, cwd=str(WORKSPACE),
                             creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0))
            return {"ok": True, "session_id": session_id, "profile": PROFILE}
        except Exception as exc:                              # noqa: BLE001
            return {"ok": False, "error": str(exc)}

    def log_message(self, fmt: str, *args: object) -> None:    # тише в консоли
        log("http", self.address_string(), fmt % args)


def main() -> int:
    crt, key = CERT_DIR / "localhost.crt", CERT_DIR / "localhost.key"
    if not crt.exists() or not key.exists():
        print("Нет dev-сертификатов. Выполни:\n  npx --yes office-addin-dev-certs install --days 365",
              file=sys.stderr)
        return 2
    ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(certfile=str(crt), keyfile=str(key))

    # На Windows SO_REUSEADDR позволяет двум процессам слушать один порт — второй мост
    # молча перехватывал бы часть запросов. Поэтому отключаем переиспользование адреса
    # и падаем с понятной ошибкой, если порт уже занят.
    class SingleInstanceServer(ThreadingHTTPServer):
        allow_reuse_address = False

    try:
        httpd = SingleInstanceServer((HOST, PORT), Handler)
    except OSError as exc:
        print(f"Порт {HOST}:{PORT} уже занят другим процессом ({exc}).\n"
              f"Останови старый мост: netstat -ano | findstr :{PORT}  →  taskkill /F /PID <pid>",
              file=sys.stderr)
        return 3
    httpd.socket = ctx.wrap_socket(httpd.socket, server_side=True)

    log(f"Hermes Excel Bridge → {ORIGIN}  (hermes={HERMES})")
    log(f"статика: {ADDIN_DIR}   рабочая папка агента: {WORKSPACE}")
    log("Ctrl+C — остановить")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        log("остановлен")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
