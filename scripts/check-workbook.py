#!/usr/bin/env python3
"""Проверяет книгу, которая сама открывает панель: есть ли в ней части xl/webextensions.

Зачем отдельный скрипт: PowerShell 5.1 не видит тип `System.IO.Compression.ZipArchiveMode`
без лишних Add-Type, а Python читает .xlsx даже когда файл открыт в Excel.

Запуск: python scripts/check-workbook.py [путь]        (по умолчанию workspace/hermes-auto.xlsx)
Код возврата 0 — панель открывается из книги, 1 — нет.
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    wb = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "workspace" / "hermes-auto.xlsx"
    manifest = ROOT / "addin" / "manifest.xml"
    if not wb.exists():
        print(f"нет книги {wb} — соберите: python scripts/auto-open-workbook.py create")
        return 1
    try:
        with zipfile.ZipFile(wb) as z:                      # zipfile открывает с общим доступом
            names = z.namelist()
            parts = [n for n in names if "webextension" in n.lower()]
            ref = ""
            for n in parts:
                if n.endswith(".xml"):
                    ref += z.read(n).decode("utf-8", "replace")
    except Exception as exc:                                # noqa: BLE001
        print(f"не смог прочитать {wb}: {exc}")
        return 1

    addin_id = ""
    if manifest.exists():
        m = re.search(r"<Id>([0-9a-fA-F-]{36})</Id>", manifest.read_text(encoding="utf-8"))
        addin_id = m.group(1) if m else ""

    ok = any(p.endswith("webextension1.xml") for p in parts)
    same = (not addin_id) or (addin_id.lower() in ref.lower())
    print(f"части webextensions: {parts or 'нет'}")
    if addin_id:
        print(f"id манифеста: {addin_id} — {'совпадает с книгой' if same else 'НЕ совпадает с книгой'}")
    if not ok:
        print("панель не откроется сама: нет xl/webextensions/webextension1.xml")
        return 1
    if not same:
        print("книга ссылается на другой Id — пересоберите: python scripts/auto-open-workbook.py create")
        return 1
    print("книга корректно открывает панель")
    return 0


if __name__ == "__main__":
    sys.exit(main())
