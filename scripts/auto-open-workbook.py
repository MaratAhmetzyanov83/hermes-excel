#!/usr/bin/env python3
"""Собирает книгу, которая САМА открывает панель надстройки Hermes при открытии.

Зачем: Excel не показывает сайлоад-надстройку на ленте и не открывает панель, пока
она ни разу не активирована в текущем процессе. Но книга может нести часть
`xl/webextensions` — «авто-открытие панели», ссылающуюся на установленный манифест
(`store="developer" storeType="Registry"`). Excel открывает панель при открытии книги
и после этого кнопка надстройки появляется на ленте. Никаких ручных кликов.

Использование:
  python scripts/auto-open-workbook.py create [path.xlsx] [--addin-id ID]
  python scripts/auto-open-workbook.py tag    <existing.xlsx> [--addin-id ID]

Путь можно не указывать: по умолчанию workspace/hermes-auto.xlsx (папка создаётся сама) —
именно этот вызов делают install.ps1 и CI.
"""
from __future__ import annotations

import argparse
import os
import re
import uuid
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MANIFEST = ROOT / "addin" / "manifest.xml"

CONTENT_TYPES_OVERRIDES = (
    '<Override PartName="/xl/webextensions/taskpanes.xml" '
    'ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>'
    '<Override PartName="/xl/webextensions/webextension1.xml" '
    'ContentType="application/vnd.ms-office.webextension+xml"/>'
)
ROOT_REL = (
    '<Relationship Id="rIdHermesTaskpanes" '
    'Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" '
    'Target="xl/webextensions/taskpanes.xml"/>'
)
TASKPANES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">'
    '<wetp:taskpane dockstate="right" visibility="1" width="430" row="1">'
    '<wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'r:id="rId1"/></wetp:taskpane></wetp:taskpanes>'
)
TASKPANES_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" '
    'Target="webextension1.xml"/></Relationships>'
)

ROWS = [("month", "revenue", "orders"), ("2026-01", 120000, 340), ("2026-02", 138000, 377),
        ("2026-03", 131500, 360), ("2026-04", 145200, 402), ("2026-05", 110000, 281)]


def sheet_xml() -> str:
    body = []
    for ri, row in enumerate(ROWS, start=1):
        cells = []
        for ci, val in enumerate(row):
            col = "ABC"[ci]
            if isinstance(val, int):
                cells.append(f'<c r="{col}{ri}"><v>{val}</v></c>')
            else:
                cells.append(f'<c r="{col}{ri}" t="inlineStr"><is><t>{val}</t></is></c>')
        body.append(f'<row r="{ri}">' + "".join(cells) + "</row>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<dimension ref="A1:C{len(ROWS)}"/><sheetData>' + "".join(body) + "</sheetData></worksheet>")


def minimal_workbook() -> dict[str, bytes]:
    parts = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            "</Types>",
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Данные" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": sheet_xml(),
    }
    return {k: v.encode("utf-8") for k, v in parts.items()}


def webextension_xml(addin_id: str) -> str:
    # Свой instance-id на книгу: Excel различает открытые экземпляры надстройки по нему.
    instance_id = "{" + str(uuid.uuid4()).upper() + "}"
    reference = f'<we:reference id="{addin_id}" version="1.0.0.0" store="developer" storeType="Registry"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" '
            f'id="{instance_id}">{reference}'
            "<we:alternateReferences/><we:properties/><we:bindings/>"
            '<we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
            "</we:webextension>")


def with_taskpane_parts(parts: dict[str, bytes], addin_id: str) -> dict[str, bytes]:
    out = {n: d for n, d in parts.items() if not n.startswith("xl/webextensions/")}
    ct = out["[Content_Types].xml"].decode("utf-8")
    if "webextensiontaskpanes" not in ct:
        out["[Content_Types].xml"] = ct.replace("</Types>", CONTENT_TYPES_OVERRIDES + "</Types>").encode("utf-8")
    rels = out["_rels/.rels"].decode("utf-8")
    if "webextensiontaskpanes" not in rels:
        out["_rels/.rels"] = rels.replace("</Relationships>", ROOT_REL + "</Relationships>").encode("utf-8")
    out["xl/webextensions/taskpanes.xml"] = TASKPANES_XML.encode("utf-8")
    out["xl/webextensions/_rels/taskpanes.xml.rels"] = TASKPANES_RELS.encode("utf-8")
    out["xl/webextensions/webextension1.xml"] = webextension_xml(addin_id).encode("utf-8")
    return out


def write_workbook(path: Path, parts: dict[str, bytes]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    try:
        os.replace(tmp, path)
    except PermissionError:
        tmp.unlink(missing_ok=True)
        raise SystemExit(f"{path} занят (открыт в Excel) — закройте книгу и повторите")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mode", choices=("create", "tag"))
    # Путь необязателен: `create` без аргументов собирает рабочую книгу авто-открытия.
    # Так его вызывают install.ps1, sideload.ps1 и CI — не заставляем помнить путь.
    ap.add_argument("path", type=Path, nargs="?", default=ROOT / "workspace" / "hermes-auto.xlsx")
    ap.add_argument("--addin-id", default=None)
    a = ap.parse_args()
    a.path.parent.mkdir(parents=True, exist_ok=True)

    addin_id = a.addin_id
    if not addin_id:
        m = re.search(r"<Id>([^<]+)</Id>", MANIFEST.read_text(encoding="utf-8"))
        if not m:
            raise SystemExit(f"не нашёл <Id> в {MANIFEST}")
        addin_id = m.group(1)

    if a.mode == "create":
        parts = minimal_workbook()
    else:
        with zipfile.ZipFile(a.path) as zin:
            parts = {i.filename: zin.read(i.filename) for i in zin.infolist()}

    write_workbook(a.path, with_taskpane_parts(parts, addin_id))
    print(f"{a.mode}: {a.path} — при открытии сам поднимет надстройку {addin_id}")


if __name__ == "__main__":
    main()
