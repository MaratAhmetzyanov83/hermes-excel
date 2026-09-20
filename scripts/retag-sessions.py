#!/usr/bin/env python3
"""Переносит метку источника у сессий, созданных надстройкой Excel (разовая починка истории).

Зачем: мост запускает `hermes chat -Q --query-file` — одноразовый прогон. Hermes помечает такие
сессии `source='oneshot'`, а oneshot входит в `INTERNAL_LISTING_SOURCES`: их прячут ВСЕ
человеческие списки (десктоп, TUI, `hermes sessions list`). Чаты, начатые в Excel, из-за этого
не были видны в Hermes. Мост теперь передаёт `--source excel`, а этот скрипт переводит на ту же
метку уже созданные сессии и (по желанию) прячет мусор от тестов.

    python scripts/retag-sessions.py --dry-run          # показать, что изменится
    python scripts/retag-sessions.py                    # перевести oneshot → excel
    python scripts/retag-sessions.py --hide-tests       # плюс скрыть явные тестовые прогоны
    python scripts/retag-sessions.py --profile excel --to excel

Перед изменением делается резервная копия state.db рядом с оригиналом.
"""
from __future__ import annotations

import argparse
import os
import shutil
import sqlite3
import sys
import time
from pathlib import Path

# Явные следы тестовых прогонов (мои проверки и бенчмарки) — их прячем из списка чатов.
TEST_MARKERS = (
    "тест устойчивости", "ответить одним словом", "ответ одним словом",
    "сколько будет 2+2", "перечислить числа", "вычислить 2+2", "решить пример 2+2",
    "count data rows", "sum column b", "посчитать строки в данных",
    "посчитать сумму по колонке", "сумма по колонке b", "ascii-marker",
    "[контекст excel] файл: hermes-auto", "проверка-тела",
)


def profile_db(profile: str) -> Path:
    base = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".local/share")) / "hermes"
    return (base / "profiles" / profile / "state.db") if profile != "default" else (base / "state.db")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--profile", default=os.environ.get("HERMES_BRIDGE_PROFILE", "excel"))
    ap.add_argument("--to", default="excel", help="новая метка источника (по умолчанию excel)")
    ap.add_argument("--hide-tests", action="store_true", help="скрыть явные тестовые прогоны")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    db = profile_db(a.profile)
    if not db.exists():
        print(f"нет стора {db}")
        return 1
    if not a.dry_run:
        backup = db.with_suffix(f".db.bak-{time.strftime('%Y%m%d-%H%M%S')}")
        shutil.copy2(db, backup)
        print(f"резервная копия: {backup}")

    con = sqlite3.connect(str(db))
    con.row_factory = sqlite3.Row
    rows = con.execute("SELECT id, source, title, message_count FROM sessions").fetchall()
    to_retag, to_hide = [], []
    for r in rows:
        title = (r["title"] or "").strip().lower()
        if a.hide_tests and any(m in title for m in TEST_MARKERS):
            to_hide.append(r["id"])          # тестовые прогоны прячем независимо от текущей метки
            continue
        if (r["source"] or "") in ("oneshot", "cli") and (r["source"] or "") != a.to:
            to_retag.append(r["id"])

    print(f"сессий всего: {len(rows)} | переведу в '{a.to}': {len(to_retag)} | скрою тестовых: {len(to_hide)}")
    if a.dry_run:
        for r in rows:
            if r["id"] in to_retag:
                print(f"  → {a.to}: {(r['title'] or '(без названия)')[:60]}")
            elif r["id"] in to_hide:
                print(f"  → скрыть: {(r['title'] or '(без названия)')[:60]}")
        print("\n(режим проверки: ничего не изменено)")
        return 0

    with con:
        if to_retag:
            con.executemany("UPDATE sessions SET source=? WHERE id=?", [(a.to, i) for i in to_retag])
        if to_hide:
            con.executemany("UPDATE sessions SET hidden=1 WHERE id=?", [(i,) for i in to_hide])
    # index настроен на эти поля — обновим статистику, чтобы списки сразу увидели изменения
    con.execute("ANALYZE")
    con.commit()
    left = con.execute("SELECT source, count(*) FROM sessions GROUP BY source").fetchall()
    print("итог по источникам:", ", ".join(f"{r[0]}={r[1]}" for r in left))
    print("проверьте списки: hermes -p %s sessions list --limit 5" % a.profile)
    con.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
