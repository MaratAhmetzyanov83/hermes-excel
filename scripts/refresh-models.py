#!/usr/bin/env python3
"""Живое обновление кэша моделей профиля (то, что делает `hermes model --refresh`, но без TUI).

Зачем: выпадающий список моделей в панели Excel строится из `provider_models_cache.json`
профиля бота; у кэша TTL 1 час, и после смены тарифа/добавления ключа список в панели
остаётся старым. Мост вызывает этот скрипт по `GET /models?refresh=1`, кнопка ⟳ в панели —
её фронт.

Запуск вручную:
    <venv>/Scripts/python.exe scripts/refresh-models.py [--home <каталог профиля>]

Печатает JSON: {"home": ..., "providers": {name: {"before": n, "after": n}}, "errors": {...}}
Кэш пишет сам hermes_cli.models: пустой ответ провайдера НЕ затирает прошлый список (так
задумано в Hermes — обрыв сети не должен обнулять каталог).
"""
from __future__ import annotations

import json
import os
import pathlib
import sys
import time
from typing import Any


def main(argv: list[str]) -> int:
    home = pathlib.Path(os.environ.get("HERMES_HOME") or (pathlib.Path.home() / ".hermes"))
    if "--home" in argv:
        home = pathlib.Path(argv[argv.index("--home") + 1])
    engine = pathlib.Path(os.environ.get("HERMES_ENGINE") or (pathlib.Path.home() / "AppData" / "Local"
                                                              / "hermes" / "hermes-agent"))
    os.environ["HERMES_HOME"] = str(home)
    sys.path.insert(0, str(engine))

    from hermes_cli import models  # noqa: PLC0415 — после подмены HERMES_HOME, иначе кэш не тот

    cache = home / "provider_models_cache.json"
    try:
        before = json.loads(cache.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        before = {}

    report: dict[str, Any] = {"home": str(home), "providers": {}, "errors": {}}
    for key in list(before):
        provider = key.split("#")[0]
        started = time.time()
        try:
            live = list(models.provider_model_ids(provider, force_refresh=True))
            report["providers"][provider] = {
                "before": len(before.get(key, {}).get("models") or []),
                "after": len(live),
                "seconds": round(time.time() - started, 1),
            }
        except Exception as exc:  # noqa: BLE001 — один упавший провайдер не отменяет остальные
            report["errors"][provider] = f"{type(exc).__name__}: {exc}"
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
