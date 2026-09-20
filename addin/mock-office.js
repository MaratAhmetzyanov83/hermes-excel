/* Строгий мок Office.js для оффлайн-проверки панели.
   Ключевое: свойства Range доступны ТОЛЬКО после load() — как в настоящем API.
   Без этого баги вида «Свойство "address" недоступно» не воспроизводятся вне Excel.

   Режимы через query-параметры (для разных сценариев в тестах):
     ?mock=blank     — выделение пустое (проверка ветки «данных нет»)
     ?mock=big       — выделение 1200×4 (проверка обрезки до 400 строк)
     ?mock=nosync    — sync() отклоняется с ошибкой (проверка обработки сбоев Excel)
*/
(function () {
  "use strict";

  const params = new URLSearchParams(location.search);
  const MODE = params.get("mock") || "";

  function notLoaded(prop) {
    const e = new Error('Свойство "' + prop + '" недоступно. Прежде чем прочесть его значение, ' +
      'вызовите метод загрузки для содержащего его объекта и вызовите "context.sync()" для контекста ' +
      'связанного запроса.');
    e.name = "OfficeExtension.Error";
    throw e;
  }

  function colToIndex(col) {
    let n = 0;
    for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function indexToCol(i) {
    let s = "", n = i + 1;
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  const RANGE_PROPS = ["address", "rowCount", "columnCount", "text", "values", "formulas"];

  function makeRange(data, loaded) {
    const loadedSet = loaded || new Set();
    const target = {};
    RANGE_PROPS.forEach((prop) => {
      Object.defineProperty(target, prop, {
        get() { return loadedSet.has(prop) ? data[prop] : notLoaded(prop); },
        set(v) {
          if (prop === "values" || prop === "formulas") {
            window.__writes.push({ address: data.address, kind: prop, payload: v });
          } else {
            throw new Error("свойство " + prop + " только для чтения");
          }
        },
        configurable: true,
      });
    });
    target.load = function (props) { String(props).split(",").forEach((p) => loadedSet.add(p.trim())); };
    target.getCell = function (r, c) {
      const head = data.address.split("!")[0];
      const first = data.address.split("!")[1].split(":")[0].match(/^([A-Z]+)(\d+)$/);
      const col = indexToCol(colToIndex(first[1]) + (c || 0));
      const row = Number(first[2]) + (r || 0);
      return makeRange({
        address: head + "!" + col + row, rowCount: 1, columnCount: 1,
        text: [[String(data.text?.[r]?.[c] ?? "")]], values: [[data.values?.[r]?.[c] ?? ""]],
        formulas: [[String(data.values?.[r]?.[c] ?? "")]],
      });
    };
    target.getResizedRange = function (dr, dc) {
      const head = data.address.split("!")[0];
      const first = data.address.split("!")[1].split(":")[0].match(/^([A-Z]+)(\d+)$/);
      const rows = dr + 1, cols = dc + 1;
      const text = Array.from({ length: rows }, (_, r) =>
        Array.from({ length: cols }, (_, c) => String(data.text?.[r]?.[c] ?? "")));
      const values = Array.from({ length: rows }, (_, r) =>
        Array.from({ length: cols }, (_, c) => data.values?.[r]?.[c] ?? ""));
      const formulas = values.map((row) => row.map((v) => (typeof v === "number" ? v : String(v))));
      const last = indexToCol(colToIndex(first[1]) + dc) + (Number(first[2]) + dr);
      return makeRange({
        address: head + "!" + first[1] + first[2] + ":" + last,
        rowCount: rows, columnCount: cols, text, values, formulas,
      });
    };
    return target;
  }

  const SHEET = "ЛИСТ1";
  // выделение I10:L18, заполнено (чтобы проверить и обычный путь, и авто-запись)
  let VALUES = [
    ["месяц", "выручка", "заказы", "средний чек"],
    ["2026-01", 120000, 340, 353], ["2026-02", 138000, 377, 366], ["2026-03", 131500, 360, 365],
    ["2026-04", 145200, 402, 361], ["2026-05", 110000, 281, 391], ["2026-06", 152300, 420, 363],
    ["2026-07", 149900, 410, 366], ["2026-08", 158700, 431, 368],
  ];
  if (MODE === "blank") {
    VALUES = [["", "", "", ""], ["", "", "", ""], ["", "", "", ""],
              ["", "", "", ""], ["", "", "", ""], ["", "", "", ""],
              ["", "", "", ""], ["", "", "", ""], ["", "", "", ""]];
  }
  if (MODE === "big") {
    VALUES = [["a", "b", "c", "d"]].concat(
      Array.from({ length: 1199 }, (_, i) => ["r" + i, i, i * 2, i * 3]));
  }
  const TEXT = VALUES.map((r) => r.map((v) => String(v)));
  let ADDR = "I10:L18";
  if (MODE === "big") ADDR = "I10:L" + (10 + VALUES.length - 1);

  window.__writes = [];
  window.__mock = { sheet: SHEET, selection: SHEET + "!" + ADDR, mode: MODE, rows: VALUES.length };

  window.Excel = {
    run: function (fn) {
      const sel = makeRange({
        address: window.__mock.selection, rowCount: VALUES.length, columnCount: 4,
        text: TEXT, values: VALUES,
      });
      const sheetObj = {
        name: SHEET, load() {},
        getRange: () => makeRange({
          address: window.__mock.selection, rowCount: 1, columnCount: 1,
          text: [[TEXT[0][0]]], values: [[VALUES[0][0]]], formulas: [[VALUES[0][0]]],
        }),
      };
      const workbook = {
        name: "test.xlsx", load() {},
        getSelectedRange: () => sel,
        getActiveCell: () => makeRange({
          address: SHEET + "!I10", rowCount: 1, columnCount: 1,
          text: [[TEXT[0][0]]], values: [[VALUES[0][0]]], formulas: [[VALUES[0][0]]],
        }),
        worksheets: { getActiveWorksheet: () => sheetObj, getItem: () => sheetObj },
      };
      const ctx = {
        workbook,
        sync: async () => {
          if (MODE === "nosync") throw new Error("Excel занят (мок)");
        },
      };
      return Promise.resolve(fn(ctx));
    },
  };

  window.Office = {
    onReady: (cb) => setTimeout(() => cb({ host: "Excel", platform: "PC" }), 0),
    context: { document: { addHandlerAsync: () => {} } },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
  };
})();
