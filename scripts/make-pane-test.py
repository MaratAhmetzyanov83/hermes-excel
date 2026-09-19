#!/usr/bin/env python3
"""Генерирует addin/test-office.html — реальная панель (taskpane.html) на строгом моке Office.js.

Смысл: панель ломается только внутри Excel, где свойства Range недоступны до load().
Тест грузит НАСТОЯЩИЙ taskpane.html (та же разметка, тот же CSS, тот же taskpane.js) с моком,
который ведёт себя так же строго. Запуск:  python scripts/make-pane-test.py
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ADDIN = ROOT / "addin"
SRC = ADDIN / "taskpane.html"
OUT = ADDIN / "test-office.html"

COLLECT = """
  <script>
    // собираем результат после того, как панель отработала
    setTimeout(function () {
      const w = window.__writes || [];
      const out = {
        ctxLabel: (document.getElementById("ctxLabel") || {}).textContent || "",
        selection: (document.getElementById("ctxCsv") || {}).value ? "прочитано" : "пусто",
        csvHead: ((document.getElementById("ctxCsv") || {}).value || "").slice(0, 34),
        suggestions: [...document.querySelectorAll(".sug b")].map((b) => b.textContent),
        modelOptions: document.getElementById("modelSel") ? document.getElementById("modelSel").options.length : 0,
        writeActions: [...document.querySelectorAll(".acts button")].map((b) => b.textContent),
        toast: (document.getElementById("toast") && !document.getElementById("toast").classList.contains("hidden"))
          ? document.getElementById("toast").textContent : "",
        session: (document.getElementById("sessionInfo") || {}).textContent || "",
        writes: w.map((x) => x.kind + "@" + x.address),
      };
      const pre = document.createElement("pre");
      pre.id = "testout";
      pre.textContent = JSON.stringify(out);
      document.body.appendChild(pre);
    }, 3000);
  </script>
"""


def main() -> None:
    html = SRC.read_text(encoding="utf-8")
    # мок должен загрузиться ДО taskpane.js
    html = html.replace('<script src="/taskpane.js"></script>',
                        '<script src="/mock-office.js"></script>\n  <script src="/taskpane.js"></script>')
    # Office.js из CDN тут не нужен: мок подменяет всё, что использует панель
    html = html.replace('<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>', '')
    html = html.replace("</body>", COLLECT + "</body>")
    html = html.replace("<title>Hermes</title>", "<title>Тест панели (мок Office.js)</title>")
    OUT.write_text(html, encoding="utf-8")
    print(f"готово: {OUT} ({len(html)} байт)")


if __name__ == "__main__":
    main()
