#!/usr/bin/env python3
"""Перехватывает файл запроса, который мост отдаёт агенту в текущем ходе.

Нужен для отладки расхождений «панель прислала контекст, а агент его не видит»:
печатает ровно то, что уходит в --query-file.
"""
from __future__ import annotations

import json
import os
import ssl
import threading
import time
import urllib.request
from pathlib import Path

QDIR = Path(os.environ.get("TEMP", "/tmp")) / "hermes-bridge"
payload = {
    "prompt": "Посчитай сумму по колонке b и ответь одной строкой.",
    "model": "deepseek-flash",
    "provider": "deepseek",
    "file_name": "книга.xlsx",
    "sheet": "Данные",
    "range": "A1:B4",
    "context": {"csv": "a,b\nкабель,10\nлоток,5\nшкаф,2"},
    "shape": [4, 2],
}

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
answer: list[str] = []


def send() -> None:
    req = urllib.request.Request("https://localhost:3443/chat", data=json.dumps(payload).encode(),
                                method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("X-Hermes-Bridge", "1")
    try:
        opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))
        with opener.open(req, timeout=180) as resp:
            for raw in resp:
                line = raw.decode("utf-8", "replace").strip()
                if line.startswith("data:") and '"text"' in line:
                    answer.append(line[:200])
                if line.startswith("event: done"):
                    break
    except Exception as exc:                                   # noqa: BLE001
        answer.append(f"(соединение: {exc})")


t = threading.Thread(target=send, daemon=True)
t.start()

captured = None
deadline = time.time() + 60
while time.time() < deadline and captured is None:
    files = sorted(QDIR.glob("*.txt")) if QDIR.exists() else []
    for f in files:
        try:
            captured = f.read_text(encoding="utf-8")
        except OSError:
            continue
        if captured:
            break
    time.sleep(0.2)

print("=== ФАЙЛ ЗАПРОСА, который увидел агент ===")
print(captured if captured else "(не поймал — ход завершился слишком быстро)")
t.join(timeout=200)
print("\n=== ответ ===")
for line in answer[:3]:
    print(line)
