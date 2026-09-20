/* Hermes для Excel — панель-чат.
   Мост (bridge.py) запускает `hermes -p <бот> chat --format stream-json` и стримит события по SSE.
   Панель умеет: читать выделение как контекст, писать ответ обратно в лист (автоматически),
   переподключаться к мосту и продолжать свои сессии. */

const BASE = location.origin;
const MAX_ROWS = 400, MAX_COLS = 60;

const $ = (id) => document.getElementById(id);
const state = {
  office: false, booted: false, workbook: "", sessionId: "", runId: null,
  busy: false, controller: null, messages: [], bridgeOk: null,
  ctx: { csv: "", meta: {}, empty: false, anchor: "" },
  autoWrite: true, useCtx: true, tries: 0,
  atts: [],                 // приложенные файлы: {name, kind, path, preview}
  ctxDirty: false,          // пользователь правил CSV контекста руками — не затирать
  frozen: null,             // цель записи, замороженная в момент отправки
  lastWrite: null,          // снимок листа для «Отменить»
  toolCount: 0, _raf: null, _status: null, _acc: "",
};

/* ------------------------------------------------------------------ утилиты */

function toCsv(rows) {
  return rows.map((r) => (r || []).map((c) => {
    const s = c == null ? "" : String(c);
    return /[",\n;]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(",")).join("\n");
}

function parseCsv(text) {
  const rows = []; let row = [], cell = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += ch;
    } else if (ch === '"') q = true;
    else if (ch === "," || ch === "\t") { row.push(cell); cell = ""; }
    else if (ch === "\n") { row.push(cell); rows.push(row); row = []; cell = ""; }
    else if (ch !== "\r") cell += ch;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((c) => c !== ""));
}

function allBlocks(md) {
  const out = [];
  const re = /```([a-z]*)\s*\n([\s\S]*?)```/gi;
  let m;
  while ((m = re.exec(md))) {
    const lang = (m[1] || "").toLowerCase();
    const kind = ["csv", "tsv", "table"].includes(lang) ? "csv"
      : ["formulas", "formula", "excel", "xl"].includes(lang) ? "formulas" : "other";
    out.push({ lang: lang || "блок", kind, body: m[2].replace(/\s+$/, "") });
  }
  return out;
}

/* Первый блок каждого вида — для кнопок записи. ВАЖНО: раньше брался только первый блок в ответе,
   и вторая таблица молча терялась. Теперь рядом есть и полный список (all) для превью. */
function blocks(md) {
  const all = allBlocks(md);
  const pick = (kind) => (all.find((b) => b.kind === kind) || {}).body || null;
  return { csv: pick("csv"), formulas: pick("formulas"), all };
}

/* текст ответа без служебных блоков — для показа в чате */
function stripBlocks(md) {
  return md.replace(/```[a-z]*\s*\n[\s\S]*?```/gi, "").replace(/\n{3,}/g, "\n\n").trim();
}

let toastTimer = null;
/* Тост с необязательным действием (кнопка): нужен для «Записано … · Отменить» —
   без отката любая запись в лист тревожит пользователя. */
function toast(text, bad, action) {
  const el = $("toast");
  el.textContent = text;
  el.className = "toast" + (bad ? " bad" : "");
  if (action) {
    const b = document.createElement("button");
    b.className = "toast-act";
    b.textContent = action.label;
    b.onclick = () => { el.classList.add("hidden"); action.fn(); };
    el.appendChild(b);
  }
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add("hidden"), action ? 10000 : (bad ? 5000 : 2600));
}

function storageKey() { return "hermes.session." + (state.workbook || "default"); }

/* ------------------------------------------------------------------ чат-рендер */

function suggestions() {
  const header = (state.ctx.csv.split("\n")[0] || "").split(",").map((s) => s.trim()).filter(Boolean);
  const list = [
    { t: "Сводка по выделенному", s: "Что в этих данных: итоги, тренд, аномалии" },
    { t: "Найти ошибки и дубли", s: "проверь пропуски, дубликаты и странные значения" },
    { t: "Привести к единому виду", s: "стандартизируй заголовки и формат чисел" },
  ];
  if (header.length >= 2) {
    list.unshift({ t: `Итоги по «${header[header.length - 1]}»`, s: `сгруппируй по «${header[0]}»` });
  }
  return list.slice(0, 4);
}

function emptyState() {
  const chat = $("chat");
  chat.innerHTML = '<div class="empty-state">Чем помочь с таблицей?</div>';
  const box = document.createElement("div");
  box.className = "suggestions";
  suggestions().forEach((s) => {
    const b = document.createElement("button");
    b.className = "sug";
    const t = document.createElement("b"); t.textContent = s.t;
    const d = document.createElement("span"); d.textContent = s.s;
    b.appendChild(t); b.appendChild(d);
    b.onclick = () => { $("prompt").value = s.t + " — " + s.s; autoGrow(); send(); };
    box.appendChild(b);
  });
  chat.appendChild(box);
}

function addMessage(role, text) {
  const wrap = document.createElement("div");
  wrap.className = "msg " + role;
  if (role === "assistant") {
    const who = document.createElement("div");
    who.className = "who";
    who.innerHTML = '<span class="dot"></span><span>' + (state.botLabel || "Hermes") + "</span><span class='meta'></span>";
    wrap.appendChild(who);
  }
  const body = document.createElement("div");
  body.className = "body";
  body.textContent = text || "";
  wrap.appendChild(body);
  $("chat").appendChild(wrap);
  $("chat").scrollTop = $("chat").scrollHeight;
  return wrap;
}

/* Автопрокрутка только когда пользователь у нижнего края: иначе чтение длинного ответа
   постоянно сбивается вниз. */
function scrollDown(force) {
  const c = $("chat");
  const near = c.scrollHeight - c.scrollTop - c.clientHeight < 90;
  if (force || near) c.scrollTop = c.scrollHeight;
}

/* Статус хода с первой секунды: до первого текста модель молчит, и панель выглядела зависшей. */
function startStatus(bubble, t0) {
  const meta = bubble.querySelector(".who .meta");
  state.toolCount = 0;
  const tick = () => {
    const s = ((Date.now() - t0) / 1000).toFixed(0);
    if (meta) meta.textContent = ` · ${s} с · ${state.toolCount} инстр.`;
  };
  tick();
  state._status = setInterval(tick, 1000);
}

function stopStatus() {
  if (state._status) { clearInterval(state._status); state._status = null; }
}

/* Отрисовка дельт не чаще одного кадра: на каждый токен шло два regex по всему тексту и форс-layout. */
function renderDelta(bubble, text) {
  state._acc = text;
  if (state._raf) return;
  state._raf = requestAnimationFrame(() => {
    state._raf = null;
    bubble.querySelector(".body").textContent = stripBlocks(state._acc);
    scrollDown();
  });
}

function previewBlocks(bubble, all) {
  if (!all || !all.length) return;
  const box = document.createElement("div");
  box.className = "preview";
  all.forEach((blk) => {
    const cap = document.createElement("div");
    cap.className = "cap";
    const lines = blk.body.split("\n").length;
    cap.textContent = `\`\`\`${blk.lang} · ${lines} строк — это уедет в лист`;
    const pre = document.createElement("pre");
    pre.className = "code";
    pre.textContent = blk.body.length > 1400 ? blk.body.slice(0, 1400) + "\n…" : blk.body;
    box.appendChild(cap);
    box.appendChild(pre);
  });
  bubble.appendChild(box);
}

function showTyping() {
  const t = document.createElement("div");
  t.className = "typing"; t.id = "typing";
  t.innerHTML = '<i></i><i></i><i></i><span style="margin-left:6px">Hermes думает…</span>';
  $("chat").appendChild(t); scrollDown();
}
function hideTyping() { const t = $("typing"); if (t) t.remove(); }

/* ------------------------------------------------------------------ Office.js */

async function readSelection() {
  const demo = "month,revenue,orders\n2026-01,120000,340\n2026-02,138000,377\n2026-03,131500,360\n2026-04,145200,402\n2026-05,110000,281";
  if (!state.office) {
    state.ctx = { csv: demo, meta: { sheet: "Демо", range: "A1:C6", shape: [6, 3], note: "" }, empty: false, anchor: "" };
    $("ctxCsv").value = demo;
    $("ctxLabel").textContent = "демо-данные (предпросмотр без Excel)";
    if (!$("chat").querySelector(".msg")) emptyState();
    return;
  }
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = ctx.workbook.getSelectedRange();
    sheet.load("name"); ctx.workbook.load("name");
    sel.load("address,rowCount,columnCount");     // сначала только размеры — выделение бывает на миллион строк
    await ctx.sync();
    state.workbook = ctx.workbook.name;

    let rows = sel.rowCount, cols = sel.columnCount, target = sel, note = "";
    if (rows > MAX_ROWS || cols > MAX_COLS) {
      rows = Math.min(rows, MAX_ROWS); cols = Math.min(cols, MAX_COLS);
      note = `взяты первые ${rows} строк и ${cols} колонок выделения`;
      target = sel.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
    }
    target.load("address,text,values");
    await ctx.sync();

    const grid = target.values || [];
    const blank = !grid.some((r) => (r || []).some((c) => c !== "" && c !== null));
    // Якорь = первая ячейка ИСХОДНОГО выделения. Берём из уже загруженной строки адреса
    // ("Лист1!I10:L18" → "Лист1!I10"): у свежесозданного getCell(0,0) свойство address ещё не загружено.
    const anchor = String(sel.address || "").split(":")[0];

    state.ctx = {
      csv: blank ? "" : toCsv(target.text && target.text.length ? target.text : grid),
      meta: { sheet: sheet.name, range: target.address, shape: [grid.length, (grid[0] || []).length], note },
      empty: blank, anchor,
    };
    if (!state.ctxDirty) $("ctxCsv").value = state.ctx.csv;   // ручные правки не затираем
    $("ctxLabel").textContent = `${state.ctx.meta.sheet}!${target.address.replace(/^.*!/, "")}`
      + (blank ? " · пусто" : "") + (note ? " · " + note : "");
    const saved = localStorage.getItem(storageKey());
    if (saved && !state.sessionId) { setSession(saved); loadHistory(); }
    if (!$("chat").querySelector(".msg")) emptyState();
  });
}

async function withSheet(fn) {
  if (!state.office) { toast("Предпросмотр без Excel: запись выключена", true); return; }
  try { await Excel.run(fn); }
  catch (e) { toast("Ошибка Excel: " + (e.message || e), true); }
}

/* Цель записи замораживается в момент ОТПРАВКИ запроса: ход длится секунды-минуты, а пользователь
   за это время ходит по листу, и без снимка результат приземлялся поверх чужих данных. */
function freezeTarget() {
  state.frozen = {
    anchor: state.ctx.anchor,
    sheet: state.ctx.meta.sheet || "",
    empty: !!state.ctx.empty,
  };
}

function targetAddress() {
  const f = state.frozen || {};
  return f.anchor || state.ctx.anchor || "";
}

function anchorCell(ctx) {
  const addr = targetAddress();
  if (addr && addr.includes("!")) {
    const sheet = addr.split("!")[0].replace(/^'|'$/g, "");
    const cell = addr.split("!")[1];
    return ctx.workbook.worksheets.getItem(sheet).getRange(cell);
  }
  return ctx.workbook.getActiveCell();
}

function writeGrid(grid, asFormulas, opts) {
  const rows = grid.length;
  const cols = Math.max(...grid.map((r) => r.length));
  const o = opts || {};
  withSheet(async (ctx) => {
    const rng = anchorCell(ctx).getResizedRange(rows - 1, cols - 1);
    rng.load("address,values,formulas");                  // узнаём, что там лежит СЕЙЧАС
    await ctx.sync();
    const prevValues = rng.values || [];
    const prevFormulas = rng.formulas || [];
    const busy = prevValues.some((r) => (r || []).some((c) => c !== "" && c !== null));
    if (busy && o.confirm && !o.forced && !o.quiet) {
      askOverwrite(rng.address, grid, asFormulas, rows, cols);   // не пишем молча поверх данных
      return;
    }
    if (asFormulas) rng.formulas = grid; else rng.values = grid;
    await ctx.sync();
    state.lastWrite = { address: rng.address, values: prevValues, formulas: prevFormulas, rows, cols };
    toast(`Записано ${rows}×${cols} в ${rng.address}`, false,
          { label: "Отменить", fn: undoWrite });
  });
}

/* Что вернуть при отмене: формулу там, где была формула, значение — где было значение. */
function undoWrite() {
  const lw = state.lastWrite;
  if (!lw) { toast("Отменять нечего", true); return; }
  state.lastWrite = null;
  withSheet(async (ctx) => {
    const rng = ctx.workbook.worksheets.getActiveWorksheet().getRange(lw.address);
    const back = (lw.values || []).map((row, i) => (row || []).map(
      (v, j) => ((lw.formulas || [])[i] || [])[j] || v));
    if (back.length) { rng.formulas = back; await ctx.sync(); }
    toast("Запись отменена — в листе снова как было");
  });
}

function askOverwrite(address, grid, asFormulas, rows, cols) {
  const wrap = addMessage("assistant", `В ${address} уже есть данные. Записывать поверх?`);
  const acts = document.createElement("div");
  acts.className = "acts";
  const mk = (label, fn) => {
    const b = document.createElement("button");
    b.textContent = label;
    b.onclick = () => { acts.remove(); fn(); };
    acts.appendChild(b);
  };
  mk(`перезаписать ${rows}×${cols}`, () => writeGrid(grid, asFormulas, { forced: true, quiet: true }));
  mk("записать в текущее выделение", () => {
    state.frozen = null;                                   // пишем туда, где курсор сейчас
    withSheet(async (ctx) => {
      const rng = ctx.workbook.getSelectedRange().getCell(0, 0).getResizedRange(rows - 1, cols - 1);
      if (asFormulas) rng.formulas = grid; else rng.values = grid;
      await ctx.sync();
      toast(`Записано ${rows}×${cols} в ${rng.address}`);
    });
  });
  mk("не писать", () => toast("Оставил лист как есть"));
  wrap.appendChild(acts);
  scrollDown(true);
}

function writeCsv(csv, opts) {
  const rows = parseCsv(csv || "");
  if (!rows.length) return false;
  const w = Math.max(...rows.map((r) => r.length));
  writeGrid(rows.map((r) => { const c = r.slice(); while (c.length < w) c.push(""); return c; }), false, opts);
  return true;
}

function writeFormulas(block, opts) {
  const lines = (block || "").split("\n").map((s) => s.trim()).filter(Boolean)
    .map((s) => [s.replace(/^`|`$/g, "")]);
  if (!lines.length) return false;
  writeGrid(lines, true, opts);
  return true;
}

/* ------------------------------------------------------------------ мост */

function renderBridge() {
  const el = $("sessionInfo");
  if (state.bridgeOk === true) {
    el.className = "muted";
    el.textContent = (state.botLabel ? state.botLabel + " · " : "") + (state.sessionId || "новый чат");
  } else if (state.bridgeOk === false) {
    el.className = "muted";
    el.innerHTML = 'мост не запущен — <a href="#" id="btnRetry" class="link">подключиться</a>';
    const b = $("btnRetry"); if (b) b.onclick = (e) => { e.preventDefault(); loadCatalog(0); };
  } else {
    el.className = "muted"; el.textContent = "подключение…";
  }
}

async function loadCatalog(attempt = 0) {
  try {
    const r = await fetch(BASE + "/models", { cache: "no-store" });
    const data = await r.json();
    const provs = data.providers || [];
    const selP = $("providerSel"), selM = $("modelSel");
    selP.innerHTML = '<option value="auto">как в Hermes</option>';
    provs.forEach((p) => selP.add(new Option(p.provider, p.provider)));
    const def = data.default || {};
    state.botLabel = data.profile && data.profile !== "default" ? `бот ${data.profile}` : "Hermes";
    const bn = $("botName"); if (bn) bn.textContent = data.profile && data.profile !== "default" ? "· " + data.profile : "";
    refreshModels(provs, def.model);
    if (!$("chat").querySelector(".msg")) emptyState();
    const savedP = localStorage.getItem("hermes.provider");
    if (savedP && [...selP.options].some((o) => o.value === savedP)) {
      selP.value = savedP; refreshModels(provs, localStorage.getItem("hermes.model"));
    }
    state.bridgeOk = true; state.tries = 0; renderBridge();
  } catch (e) {
    state.bridgeOk = false; renderBridge();
    if (attempt < 40) setTimeout(() => loadCatalog(attempt + 1), 3000);   // сам поднимется, когда мост вернётся
  }
}

function refreshModels(provs, preselect) {
  const selP = $("providerSel"), selM = $("modelSel");
  selM.innerHTML = "";
  selM.add(new Option("auto", "auto"));
  const chosen = selP.value || "auto";
  (provs || []).filter((p) => chosen === "auto" || p.provider === chosen)
    .forEach((p) => p.models.forEach((m) => selM.add(new Option(`${m} · ${p.provider}`, `${m}|${p.provider}`))));
  const saved = preselect || localStorage.getItem("hermes.model") || "auto";
  [...selM.options].forEach((o) => { if (o.value === saved) selM.value = saved; });
}

async function loadSessions() {
  try {
    const r = await fetch(BASE + "/sessions?limit=40", { cache: "no-store" });
    const data = await r.json();
    const sel = $("sessionPick");
    sel.innerHTML = '<option value="">продолжить другой чат…</option>';
    (data.sessions || []).forEach((s) => {
      const when = new Date((s.last_activity_at || s.started_at) * 1000)
        .toLocaleString("ru-RU", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
      sel.add(new Option(`${when} · ${(s.title || s.id).slice(0, 46)}`, s.id));
    });
  } catch (e) { /* мост молчит */ }
}

/* Лента чата после перезапуска Excel: сессия на книгу сохранена, но панель начинала с пустого
   экрана, и «что я тут делал вчера» приходилось вспоминать через CLI. */
async function loadHistory() {
  if (!state.sessionId) return;
  try {
    const r = await fetch(BASE + "/session?id=" + encodeURIComponent(state.sessionId), { cache: "no-store" });
    const d = await r.json();
    const msgs = d.messages || [];
    if (!msgs.length) return;
    const chat = $("chat");
    chat.innerHTML = "";
    msgs.forEach((m) => {
      const el = addMessage(m.role === "user" ? "user" : "assistant", stripBlocks(m.text) || m.text);
      el.classList.add("old");
    });
    scrollDown(true);
  } catch (e) { /* мост молчит — не мешаем работать */ }
}

function setSession(id) {
  state.sessionId = id || "";
  if (id) localStorage.setItem(storageKey(), id); else localStorage.removeItem(storageKey());
  renderBridge();
}

/* ------------------------------------------------------------------ вложения */

const MAX_ATT = 8;

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onerror = () => reject(new Error("не смог прочитать файл"));
    fr.onload = () => resolve(String(fr.result).split(",")[1] || "");
    fr.readAsDataURL(file);
  });
}

function renderAtts() {
  const box = $("atts"), row = $("attsRow");
  box.classList.toggle("hidden", state.atts.length === 0);
  row.innerHTML = "";
  state.atts.forEach((a) => {
    const chip = document.createElement("div");
    chip.className = "att-chip" + (a.busy ? " busy" : "") + (a.error ? " failed" : "");
    if (a.preview) {
      const img = document.createElement("img");
      img.src = a.preview;
      chip.appendChild(img);
    }
    const nm = document.createElement("span");
    nm.className = "nm";
    nm.textContent = a.name + (a.error ? " — " + a.error : a.busy ? " — загружаю…" : "");
    nm.title = a.path || a.name;
    chip.appendChild(nm);
    const x = document.createElement("span");
    x.className = "x";
    x.textContent = "✕";
    x.title = "убрать";
    x.onclick = () => {
      if (a.preview) URL.revokeObjectURL(a.preview);
      state.atts = state.atts.filter((b) => b !== a);
      renderAtts();
    };
    chip.appendChild(x);
    row.appendChild(chip);
  });
}

/* Файл уходит в мост (POST /upload) и ложится в workspace/pane/uploads: агенту передаём путь. */
async function addFiles(files) {
  const list = [...(files || [])].filter(Boolean);
  if (!list.length) return;
  for (const f of list) {
    if (state.atts.length >= MAX_ATT) { toast(`Больше ${MAX_ATT} файлов за раз не беру`, true); break; }
    const isImg = (f.type || "").startsWith("image/") || /\.(png|jpe?g|gif|webp|bmp)$/i.test(f.name);
    const att = { name: f.name || "file", kind: isImg ? "image" : "file", busy: true,
                  preview: isImg ? URL.createObjectURL(f) : "" };
    state.atts.push(att);
    renderAtts();
    try {
      const data = await fileToBase64(f);
      const r = await fetch(BASE + "/upload", {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-Hermes-Bridge": "1" },
        body: JSON.stringify({ name: f.name, kind: att.kind, data }),
      });
      const d = await r.json();
      if (!d.ok) throw new Error(d.error || ("HTTP " + r.status));
      att.path = d.path; att.busy = false;
    } catch (e) {
      att.busy = false; att.error = (e.message || "ошибка").slice(0, 40);
      toast("Файл не приложился: " + (e.message || e), true);
    }
    renderAtts();
  }
}

function clearAtts() {
  state.atts.forEach((a) => { if (a.preview) URL.revokeObjectURL(a.preview); });
  state.atts = [];
  renderAtts();
}

/* ------------------------------------------------------------------ отправка */

async function send() {
  if (state.busy) return;
  const ready = state.atts.filter((a) => a.path);
  if (state.atts.some((a) => a.busy)) { toast("Секунду: файл ещё загружается", true); return; }
  const prompt = $("prompt").value.trim() || (ready.length ? "Посмотри вложение и скажи, что в нём." : "");
  if (!prompt) return;
  if (state.bridgeOk === false) { toast("Мост не запущен: start-bridge.cmd", true); loadCatalog(0); return; }

  const manualCsv = $("ctxCsv").value.trim();
  const useCsv = state.useCtx ? (manualCsv || state.ctx.csv) : "";
  const payload = {
    prompt,
    session_id: state.sessionId || "",
    model: ($("modelSel").value || "auto").split("|")[0],
    provider: ($("modelSel").value || "auto").includes("|")
      ? $("modelSel").value.split("|")[1] : $("providerSel").value,
    reasoning: $("reasonSel").value,
    file_name: state.workbook, sheet: state.ctx.meta.sheet || "", range: state.ctx.meta.range || "",
    shape: state.ctx.meta.shape || null,
    context: { csv: useCsv, note: state.ctx.meta.note || "", empty: !useCsv },
    attachments: ready.map((a) => ({ path: a.path, name: a.name, kind: a.kind })),
  };

  $("prompt").value = "";
  autoGrow();
  if (!$("chat").querySelector(".msg")) $("chat").innerHTML = "";
  addMessage("user", prompt + (ready.length ? "\n📎 " + ready.map((a) => a.name).join(", ") : ""));
  freezeTarget();                                  // пишем туда, где было выделено при отправке
  state.busy = true; state.runId = null; state.toolCount = 0;
  $("btnSend").classList.add("hidden"); $("btnStop").classList.remove("hidden");
  showTyping();

  const t0 = Date.now();
  state.controller = new AbortController();
  let acc = "";
  let bubble = null;
  try {
    const res = await fetch(BASE + "/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Hermes-Bridge": "1" },
      body: JSON.stringify(payload),
      signal: state.controller.signal,
    });
    if (!res.ok || !res.body) throw new Error("HTTP " + res.status);
    clearAtts();                                   // файлы уже на диске и переданы агенту
    const reader = res.body.getReader();
    const dec = new TextDecoder();
    let buf = "";
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += dec.decode(value, { stream: true });
      let idx;
      while ((idx = buf.indexOf("\n\n")) >= 0) {
        const frame = buf.slice(0, idx); buf = buf.slice(idx + 2);
        const ev = /^event:\s*(.+)$/m.exec(frame);
        const dt = /^data:\s*(.+)$/m.exec(frame);
        if (!ev || !dt) continue;
        let data = {}; try { data = JSON.parse(dt[1]); } catch (e) { continue; }
        const kind = ev[1].trim();
        if (kind === "start") {
          state.runId = data.run_id;
          hideTyping();
          bubble = addMessage("assistant", "");
          bubble.classList.add("working");
          startStatus(bubble, t0);                       // видно, что работа идёт, ещё до первого текста
        } else if (kind === "session") {
          setSession(data.session_id);
        } else if (kind === "delta") {
          acc += data.text || "";
          if (!bubble) { hideTyping(); bubble = addMessage("assistant", ""); startStatus(bubble, t0); }
          bubble.classList.remove("working");
          renderDelta(bubble, acc);
        } else if (kind === "tool" || kind === "log") {
          // Инструменты пишем всегда, а не только после первого текста: именно они показывают,
          // что агент жив, когда модель «думает» десятки секунд.
          if (!bubble) { hideTyping(); bubble = addMessage("assistant", ""); startStatus(bubble, t0); }
          state.toolCount = (state.toolCount || 0) + (kind === "tool" ? 1 : 0);
          let tl = bubble.querySelector(".tools");
          if (!tl) { tl = document.createElement("div"); tl.className = "tools"; bubble.appendChild(tl); }
          tl.textContent += (tl.textContent ? "\n" : "") + "▸ " + (data.name || "log") +
            (data.detail ? " — " + String(data.detail).slice(0, 90) : "");
        } else if (kind === "result") {
          if (data.text) {
            acc = data.text;
            if (!bubble) { hideTyping(); bubble = addMessage("assistant", ""); }
            bubble.querySelector(".body").textContent = stripBlocks(acc);
          }
          setSession(data.session_id || state.sessionId);
        } else if (kind === "error") {
          hideTyping(); stopStatus();
          const b = addMessage("assistant", data.message || "ошибка моста");
          b.classList.add("failed");
          if (data.hint) {
            const h = document.createElement("div");
            h.className = "hint";
            h.title = "нажмите, чтобы скопировать";
            h.textContent = data.hint;
            h.onclick = () => { navigator.clipboard.writeText(data.hint); toast("команда скопирована"); };
            b.appendChild(h);
          }
          bubble = null;
        } else if (kind === "done") {
          hideTyping(); stopStatus();
          if (!bubble) bubble = addMessage("assistant", acc ? stripBlocks(acc) : "(пустой ответ)");
          const who = bubble.querySelector(".who .meta");
          if (who) who.textContent = ` · ${((Date.now() - t0) / 1000).toFixed(1)} с`
            + (state.toolCount ? ` · ${state.toolCount} инстр.` : "");
          bubble.classList.remove("working");
          bubble.classList.add("done");
          const b = blocks(acc);
          previewBlocks(bubble, b.all);                   // что именно уедет в лист — до записи
          const acts = document.createElement("div");
          acts.className = "acts";
          const mk = (label, fn) => {
            const btn = document.createElement("button");
            btn.textContent = label;
            btn.onclick = fn;
            acts.appendChild(btn);
          };
          if (b.csv || b.formulas) {
            if (b.csv) mk("записать таблицу в лист", () => writeCsv(b.csv, { quiet: true }));
            if (b.formulas) mk("записать формулы", () => writeFormulas(b.formulas, { quiet: true }));
            mk("вставить текст в ячейку", () => {
              withSheet(async (ctx) => {
                anchorCell(ctx).values = [[stripBlocks(acc) || acc]];
                await ctx.sync(); toast("текст записан");
              });
            });
          }
          if (acc) mk("скопировать ответ", () => {
            navigator.clipboard.writeText(stripBlocks(acc) || acc); toast("ответ скопирован");
          });
          if (b.csv) mk("скопировать CSV", () => {
            navigator.clipboard.writeText(b.csv); toast("CSV скопирован");
          });
          bubble.appendChild(acts);
          if (state.autoWrite) {
            // confirm: не затираем непустой диапазон молча — панель спросит и предложит вариант
            if (b.csv) writeCsv(b.csv, { confirm: true });
            else if (b.formulas) writeFormulas(b.formulas, { confirm: true });
          }
          loadSessions();
        }
      }
    }
  } catch (e) {
    hideTyping();
    if (e.name !== "AbortError") {
      addMessage("assistant", "Не получилось: " + (e.message || e));
      toast("Сбой связи с мостом", true);
    }
  } finally {
    state.busy = false; state.controller = null;
    stopStatus();
    $("btnSend").classList.remove("hidden"); $("btnStop").classList.add("hidden");
  }
}

async function stop() {
  try {
    await fetch(BASE + "/stop", {
      method: "POST", headers: { "Content-Type": "application/json", "X-Hermes-Bridge": "1" },
      body: JSON.stringify({ run_id: state.runId || "", session_id: state.runId ? "" : state.sessionId }),
    });
  } catch (e) { /* мост не ответил */ }
  if (state.controller) state.controller.abort();
  hideTyping();
  toast("Остановлено");
}

async function openInHermes() {
  const sid = state.sessionId || $("sessionPick").value;
  if (!sid) { toast("Сначала напишите Hermes", true); return; }
  try {
    const r = await fetch(BASE + "/open", {
      method: "POST", headers: { "Content-Type": "application/json", "X-Hermes-Bridge": "1" },
      body: JSON.stringify({ session_id: sid }),
    });
    const d = await r.json();
    toast(d.ok ? "Открываю чат в Hermes…" : "Не вышло: " + (d.error || "?"), !d.ok);
  } catch (e) { toast("Мост недоступен", true); }
}

/* ------------------------------------------------------------------ запуск */

function autoGrow() {
  const t = $("prompt");
  t.style.height = "auto";
  t.style.height = Math.min(t.scrollHeight, 132) + "px";
}

function wire() {
  $("btnSend").onclick = send;
  $("btnStop").onclick = stop;
  $("btnOpen").onclick = openInHermes;
  $("btnNew").onclick = () => {
    setSession("");
    state.messages = [];
    emptyState();
    toast("Новый чат");
  };
  $("btnSettings").onclick = () => $("settings").classList.toggle("hidden");
  $("btnRefresh").onclick = () => { state.ctxDirty = false; readSelection().then(() => toast("Выделение обновлено")).catch((e) => toast(e.message, true)); };
  $("btnEditCtx").onclick = () => $("ctxEdit").classList.toggle("hidden");
  $("btnCtxDone").onclick = () => { $("ctxEdit").classList.add("hidden"); toast("Данные для контекста обновлены"); };
  $("ctxCsv").addEventListener("input", () => { state.ctxDirty = true; });   // правки не теряются
  // чип контекста — переключатель: решение «брать выделение или нет» принимается на каждой реплике,
  // а тумблер в настройках для этого слишком далеко
  $("ctxLabel").title = "нажмите, чтобы включать/выключать контекст выделения";
  $("ctxLabel").style.cursor = "pointer";
  $("ctxLabel").onclick = () => {
    state.useCtx = !state.useCtx;
    $("useCtx").checked = state.useCtx;
    localStorage.setItem("hermes.usectx", state.useCtx ? "1" : "0");
    $("ctxLabel").classList.toggle("off", !state.useCtx);
    toast(state.useCtx ? "Контекст выделения включён" : "Контекст выделения выключен");
  };
  $("modelSel").onchange = () => localStorage.setItem("hermes.model", $("modelSel").value);
  $("providerSel").onchange = () => {
    localStorage.setItem("hermes.provider", $("providerSel").value);
    fetch(BASE + "/models").then((r) => r.json()).then((d) => refreshModels(d.providers, "")).catch(() => {});
  };
  $("sessionPick").onchange = () => {
    const v = $("sessionPick").value;
    if (v) { setSession(v); loadHistory(); toast("Продолжаю чат " + v); }
  };
  $("prompt").addEventListener("input", autoGrow);
  // вложения: скрепка, перетаскивание в панель, вставка из буфера
  $("btnAttach").onclick = () => $("filePick").click();
  $("filePick").onchange = () => { addFiles($("filePick").files); $("filePick").value = ""; };
  let dragDepth = 0;
  document.addEventListener("dragenter", (e) => {
    if (![...(e.dataTransfer?.types || [])].includes("Files")) return;
    e.preventDefault(); dragDepth++; document.body.classList.add("drag");
  });
  document.addEventListener("dragover", (e) => {
    if ([...(e.dataTransfer?.types || [])].includes("Files")) e.preventDefault();
  });
  document.addEventListener("dragleave", () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) document.body.classList.remove("drag");
  });
  document.addEventListener("drop", (e) => {
    dragDepth = 0; document.body.classList.remove("drag");
    const files = e.dataTransfer?.files;
    if (files && files.length) { e.preventDefault(); addFiles(files); }
  });
  document.addEventListener("paste", (e) => {
    const files = [...(e.clipboardData?.files || [])];
    if (files.length) { e.preventDefault(); addFiles(files); }
  });
  $("prompt").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); }
  });
  $("autoWrite").checked = state.autoWrite;
  $("autoWrite").onchange = () => { state.autoWrite = $("autoWrite").checked; localStorage.setItem("hermes.autowrite", state.autoWrite ? "1" : "0"); };
  $("useCtx").checked = state.useCtx;
  $("useCtx").onchange = () => { state.useCtx = $("useCtx").checked; localStorage.setItem("hermes.usectx", state.useCtx ? "1" : "0"); };
  if (localStorage.getItem("hermes.autowrite") === "0") { state.autoWrite = false; $("autoWrite").checked = false; }
  if (localStorage.getItem("hermes.usectx") === "0") { state.useCtx = false; $("useCtx").checked = false; }
}

async function boot(isOffice) {
  if (state.booted) return;
  state.booted = true;
  state.office = isOffice;
  wire();
  emptyState();
  await loadCatalog(0);
  loadSessions();
  if (state.office) {
    const saved = localStorage.getItem(storageKey());
    if (saved) { setSession(saved); loadHistory(); }
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
        clearTimeout(state._t);
        state._t = setTimeout(() => readSelection().catch(() => {}), 400);
      });
    } catch (e) { /* старые сборки Office */ }
  }
  readSelection().catch((e) => toast("Чтение выделения: " + (e.message || e), true));
}

/* Основной путь — Office.js; страховка, если он не загрузился (нет сети/CDN). */
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(() => {
    const hasDoc = typeof Excel !== "undefined" && !!(Office.context && Office.context.document);
    boot(hasDoc);
  });
  setTimeout(() => boot(false), 3000);
} else {
  boot(false);
}
