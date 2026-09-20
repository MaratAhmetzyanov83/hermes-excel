#!/usr/bin/env python3
"""Генерирует addin/test-office.html — реальная панель (taskpane.html) на строгом моке Office.js.

Смысл: панель ломается только внутри Excel, где свойства Range недоступны до load().
Тест грузит НАСТОЯЩИЙ taskpane.html (та же разметка, тот же CSS, тот же taskpane.js) с моком,
который ведёт себя так же строго, и прогоняет сценарии записи в лист:

  фаза 1 — запись в НЕПУСТОЙ диапазон с confirm → панель обязана спросить, а не писать молча;
  фаза 2 — принудительная запись → снимок листа + предложение «Отменить» в тосте;
  фаза 3 — нажатие «Отменить» → возврат прежнего содержимого (ещё одна запись в мок).

Запуск:  python scripts/make-pane-test.py
Проверка: headless Chrome --dump-dom https://localhost:3443/test-office.html
          (режимы: ?mock=blank, ?mock=big, ?mock=nosync)
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ADDIN = ROOT / "addin"
SRC = ADDIN / "taskpane.html"
OUT = ADDIN / "test-office.html"

COLLECT = """
  <script>
    // Панель на моке: гоняем сценарии записи и собираем результат в <pre id="testout">.
    const OUT = {};
    const toastText = () => {
      const t = document.getElementById("toast");
      return t && !t.classList.contains("hidden") ? t.textContent : "";
    };
    setTimeout(() => {                       // фаза 1: непустая цель → должно спросить
      writeCsv("итог,значение\\nвсего,17", { confirm: true });
    }, 1000);
    setTimeout(() => {                       // фаза 2: фиксируем вопрос и пишем принудительно
      OUT.confirmButtons = [...document.querySelectorAll(".acts button")].map((b) => b.textContent);
      OUT.writesBeforeForce = (window.__writes || []).length;
      writeCsv("итог,значение\\nвсего,17", { forced: true, quiet: true });
    }, 1800);
    setTimeout(() => {                       // фаза 3: нажимаем «Отменить»
      OUT.toastBeforeUndo = toastText();
      OUT.undoButton = !!document.querySelector(".toast-act");
      OUT.writesAfterWrite = (window.__writes || []).length;
      const undo = document.querySelector(".toast-act");
      if (undo) undo.click();
    }, 2600);
    setTimeout(() => {                       // сбор итога
      const w = window.__writes || [];
      OUT.ctxLabel = (document.getElementById("ctxLabel") || {}).textContent || "";
      OUT.selection = (document.getElementById("ctxCsv") || {}).value ? "прочитано" : "пусто";
      OUT.csvHead = ((document.getElementById("ctxCsv") || {}).value || "").slice(0, 34);
      OUT.ctxLines = ((document.getElementById("ctxCsv") || {}).value || "").split("\\n").length;
      OUT.suggestions = [...document.querySelectorAll(".sug b")].map((b) => b.textContent);
      OUT.modelOptions = document.getElementById("modelSel") ? document.getElementById("modelSel").options.length : 0;
      OUT.previewBlocks = document.querySelectorAll(".preview pre.code").length;
      OUT.writes = w.map((x) => x.kind + "@" + x.address);
      OUT.restoredPayload = (w[w.length - 1] || {}).payload || null;
      OUT.toastAfterUndo = toastText();
      OUT.session = (document.getElementById("sessionInfo") || {}).textContent || "";
      OUT.mock = window.__mock || {};
      const pre = document.createElement("pre");
      pre.id = "testout";
      pre.textContent = JSON.stringify(OUT);
      document.body.appendChild(pre);
    }, 3400);
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
