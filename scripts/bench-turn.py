#!/usr/bin/env python3
"""Замер хода панели «как из Excel»: UTF-8 тело, реальный CSV-контекст, счёт времени.

Запуск: python scripts/bench-turn.py ["текст задачи"]
Печатает время до старта, до первого текста, до результата, число кругов с инструментами и ответ.
"""
from __future__ import annotations

import json
import os
import ssl
import sys
import time
import http.client

HOST, PORT = "localhost", int(os.environ.get("HERMES_BRIDGE_PORT", "3443"))
CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE

payload = {
    "prompt": sys.argv[1] if len(sys.argv) > 1 else "Посчитай сумму по колонке b и ответь одной строкой.",
    "model": "deepseek-flash",
    "provider": "deepseek",
    "file_name": "книга.xlsx",
    "sheet": "Данные",
    "range": "A1:B4",
    "context": {"csv": "a,b\nкабель,10\nлоток,5\nшкаф,2"},
    "shape": [4, 2],
}

conn = http.client.HTTPSConnection(HOST, PORT, context=CTX, timeout=300)
body = json.dumps(payload).encode("utf-8")                    # строго UTF-8
conn.request("POST", "/chat", body=body,
             headers={"Content-Type": "application/json", "X-Hermes-Bridge": "1",
                      "Origin": f"https://{HOST}:{PORT}"})
t0 = time.time()
resp = conn.getresponse()
print(f"HTTP {resp.status} за {time.time() - t0:.2f} с")
if resp.status != 200:
    print(resp.read().decode("utf-8", "replace")[:300])
    sys.exit(1)

tools = 0
first_text = None
answer = ""
for raw in resp:
    line = raw.decode("utf-8", "replace").strip()
    if line.startswith("event:"):
        kind = line.split(":", 1)[1].strip()
    elif line.startswith("data:"):
        dt = time.time() - t0
        try:
            data = json.loads(line[5:])
        except Exception:                                     # noqa: BLE001
            continue
        if kind == "tool":
            tools += 1
            print(f"  {dt:6.2f} с  инструмент #{tools}: {data.get('name')}")
        elif kind == "delta" and first_text is None:
            first_text = dt
            print(f"  {dt:6.2f} с  первый текст")
        elif kind == "result":
            print(f"  {dt:6.2f} с  результат: токены {data.get('tokens')}, прогон {data.get('duration_ms')} мс")
            answer = data.get("text", "")
        elif kind == "error":
            print(f"  {dt:6.2f} с  ОШИБКА: {str(data)[:160]}")
        elif kind == "done":
            print(f"  {dt:6.2f} с  конец (exit={data.get('exit_code')})")
            break

print(f"\nвсего {time.time() - t0:.1f} с · кругов с инструментами: {tools}")
print("ответ:", answer[:300] or "(пусто)")
conn.close()
