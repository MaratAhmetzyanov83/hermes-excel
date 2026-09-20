#!/usr/bin/env python3
"""Замер прогона бота: показывает, на каком событии и сколько времени тратится.

Запуск: python scripts/bench-bot.py ["промпт"]
Печатает таблицу «секунда · тип события · краткое содержание», итог и разбивку по фазам.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
WORKSPACE = ROOT / "workspace"
WORKSPACE.mkdir(exist_ok=True)
PROFILE = "excel"
prompt = sys.argv[1] if len(sys.argv) > 1 else "Сколько строк в данных? Ответь одним предложением."
qfile = WORKSPACE / ".bench-query.txt"
qfile.write_text(prompt, encoding="utf-8")

cmd = ["hermes", "-p", PROFILE, "chat", "-Q", "--format", "stream-json"]
cmd += [w for w in os.environ.get("BENCH_FLAGS", "").split() if w]     # сравнение конфигураций
cmd += ["--query-file", str(qfile), "--max-turns", os.environ.get("BENCH_TURNS", "40"), "--in", str(WORKSPACE)]
print("$", " ".join(cmd[:6]), "…\n")

t0 = time.time()
proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, cwd=str(WORKSPACE),
                        text=True, encoding="utf-8", errors="replace", bufsize=1)
first: dict[str, float] = {}
tools = 0
rows: list[tuple[float, str, str]] = []
try:
    for line in proc.stdout:                                  # type: ignore[union-attr]
        dt = time.time() - t0
        line = line.strip()
        if not line:
            continue
        try:
            evt = json.loads(line)
        except json.JSONDecodeError:
            rows.append((dt, "log", line[:90]))
            continue
        kind = str(evt.get("type"))
        first.setdefault(kind, dt)
        if kind in ("tool", "tool_use", "tool_call"):
            tools += 1
            detail = json.dumps(evt.get("input") or evt.get("detail") or {}, ensure_ascii=False)[:80]
            rows.append((dt, f"tool #{tools}", f"{evt.get('name') or evt.get('tool')} {detail}"))
        elif kind == "text":
            if "text" not in first:
                rows.append((dt, "text(первый)", str(evt.get("text", ""))[:60]))
            first["text"] = first.get("text", dt)
        elif kind == "result":
            rows.append((dt, "result", f"exit={evt.get('exit_code')} tokens={evt.get('tokens')}"))
        else:
            rows.append((dt, kind, json.dumps(evt, ensure_ascii=False)[:90]))
        if len(rows) > 60:
            break
finally:
    proc.wait(timeout=30)
total = time.time() - t0

print(f"{'сек':>7}  {'событие':<14} содержание")
for dt, kind, brief in rows:
    print(f"{dt:7.2f}  {kind:<14} {brief}")
print(f"\nвсего: {total:.1f} с · кругов с инструментами: {tools}")
for kind, dt in sorted(first.items(), key=lambda kv: kv[1]):
    print(f"  {kind:<12} впервые на {dt:.2f} с")
qfile.unlink(missing_ok=True)
