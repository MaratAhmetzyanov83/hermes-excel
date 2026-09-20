# AGENTS.md — установка и работа с надстройкой «Hermes для Excel»

Этот файл — инструкция для ИИ-агента (Hermes, Claude Code, Codex и т.п.), который ставит надстройку
на машину пользователя. Выполняй шаги по порядку и **проверяй каждый** — процедура ниже проверена
на живой Windows-машине, включая все грабли из раздела «Грабли».

## Быстрый путь на новой машине

```bash
git clone https://github.com/MaratAhmetzyanov83/hermes-excel.git "$USERPROFILE/hermes-excel"
cd "$USERPROFILE/hermes-excel"
install.cmd            # сертификаты → бот → сайлоад → сторож → мост → doctor, идемпотентно
doctor.cmd             # таблица проверок с готовыми командами-исправлениями
```

`install.cmd -DryRun` показывает план, ничего не меняя; `-Profile <имя>` задаёт имя бота;
`-NoBridge` пропускает сторожа. Снять всё: `scripts\uninstall.ps1` (удаление идемпотентно).
Если сертификата нет и нет Node, установщик выдаст его через `openssl` из git-bash и добавит
в доверенные корни — Node для установки не обязателен.

## Что это такое

Надстройка Excel (Office.js), внутри которой работает агент Hermes: панель-чат в задачной области Excel,
локальный HTTPS-мост между Excel и Hermes, отдельный бот-профиль со своими сессиями.

```
Excel (Office.js pane) ──HTTPS+SSE──► bridge.py (127.0.0.1:3443)
        ▲                                   │ запускает на каждую реплику
        │  ◄── дельты ответа, лог инструментов│ hermes -p <бот> chat --format stream-json
        └── запись результата в лист         │ --resume <сессия> -m <модель>
                                             ▼
                            $HERMES_HOME/profiles/<бот>/state.db  ← сессии бота
```

## Требования

| Что | Зачем | Проверка |
|---|---|---|
| Windows 10/11 + Excel (desktop, Office 16.0) | Host надстройки, реестровый сайлоад | `ls "/c/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"` |
| Hermes Agent, настроенный и с рабочей моделью | Мост вызывает `hermes chat` | `hermes --version`, `hermes -z "скажи ок"` |
| Python 3.9+ (тот, что видит `hermes`) | Мост и скрипты; зависимостей нет | `python -V` |
| Node.js (только для сертификатов) | `npx office-addin-dev-certs` | `node -v` |
| git-bash/POSIX-shell как shell агента | команды ниже в POSIX-синтаксисе | — |

Скрипты серверной части — чистый stdlib Python; Node нужен один раз, чтобы выдать доверенный
сертификат для `https://localhost`.

## Установка (Windows)

```bash
# 0. Куда ставим
git clone https://github.com/<owner>/hermes-excel.git "$USERPROFILE/hermes-excel"
cd "$USERPROFILE/hermes-excel"

# 1. Доверенный сертификат для https://localhost (Excel требует HTTPS для панелей)
npx --yes office-addin-dev-certs install --days 365
#    → появится ~/.office-addin-dev-certs/{localhost.crt,localhost.key,ca.crt}

# 2. Бот для Excel (в Hermes бот = профиль: свои сессии, память, скиллы, модель)
hermes profile create excel --clone --description "Excel-бот: живёт в надстройке Excel"
hermes -p excel config set model.default deepseek-flash     # быстрая модель: панель ждёт ответа

# 3. Регистрация надстройки в Excel + сборка книги с авто-открытием панели
powershell -ExecutionPolicy Bypass -File scripts/sideload.ps1

# 4. Запуск моста
start-bridge.cmd          # или: python bridge/bridge.py
```

После этого в Excel: кнопка **Hermes** на вкладке «Главная» (появляется после первой активации
надстройки в этом процессе Excel). Панель открывается сама, если открыть книгу
`workspace/hermes-auto.xlsx` — она несёт часть `xl/webextensions` с авто-открытием.

## Проверка установки (по шагам, каждая команда с ожидаемым результатом)

**Сначала просто запусти `doctor.cmd`** — он проверяет всё это разом (мост, бот, сторож, сертификаты,
реестр, книгу, BOM, python, Excel) и печатает `PASS/WARN/FAIL` с готовой командой-исправлением
(машиночитаемо: `doctor.ps1 -Json`; поднять мост заодно: `doctor.cmd fix`). Шаги ниже — то же самое
вручную, когда нужно понять, где именно порвалось.

```bash
# 1. Мост жив и смотрит в нужный профиль
curl -s --ssl-no-revoke https://localhost:3443/health
#    ожидаем: {"ok": true, "profile": "excel", "profile_home": "...profiles\\excel", ...}
#    ВАЖНО: curl на Windows идёт через Schannel → нужен --ssl-no-revoke, иначе
#    (35) CRYPT_E_NO_REVOCATION_CHECK. Для PowerShell/WebView2 сертификат доверенный.

# 2. Список моделей и модель бота по умолчанию
curl -s --ssl-no-revoke https://localhost:3443/models | python -m json.tool | head -20
#    ожидаем: "profile": "excel", "default": {"model": "<быстрая модель>", ...}, провайдеры с моделями

# 3. Надстройка зарегистрирована в Excel
powershell -NoProfile -Command "(Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer').'<addin-id>'"
#    ожидаем: путь к addin\manifest.xml (id берётся из <Id> в addin/manifest.xml)

# 4. Манифест валиден по схеме Microsoft
cd "$LOCALAPPDATA/Temp" && npx --yes office-addin-manifest validate "<путь>/addin/manifest.xml"
#    ожидаем: "The manifest is valid."

# 5. Панель проходит оффлайн-тест на строгом моке Office.js (без Excel!)
python scripts/make-pane-test.py
"/c/Program Files/Google/Chrome/Application/chrome.exe" --headless=new --disable-gpu \
  --user-data-dir="$LOCALAPPDATA/Temp/chrome-pane" --virtual-time-budget=15000 \
  --dump-dom "https://localhost:3443/test-office.html" | grep -o '<pre id="testout">.*</pre>'
#    ожидаем: ctxLabel = "<Лист>!<диапазон>", toast = "", modelOptions > 0, suggestions непустые

# 6. Диалог end-to-end (создаст сессию в сторе бота)
#    POST-эндпоинты требуют заголовок X-Hermes-Bridge: 1 (защита от CSRF), тело — строго UTF-8.
#    Кириллица в `curl -d` из git-bash уезжает в cp1251 → мост отвечает 400. Пишите тело в файл.
python - <<'PY'
import json, pathlib
pathlib.Path("workspace/q.json").write_text(json.dumps(
    {"prompt": "Сколько строк данных?", "model": "auto", "context": {"csv": "a,b\n1,2\n3,4"}}),
    encoding="utf-8")
PY
curl -sN --ssl-no-revoke -X POST -H "Content-Type: application/json" -H "X-Hermes-Bridge: 1" \
  --data-binary @workspace/q.json https://localhost:3443/chat | grep -E "^event: (result|done)" -A1
hermes -p excel sessions list --limit 3     # та же сессия видна в сторе бота

# 6b. Регрессии моста без Excel: обход каталога, инъекция в /open, CSRF, мусорное тело,
#     адресный стоп — 13 проверок
python scripts/test-bridge-security.py

# 6c. Панель: оффлайн-тест на строгом моке Office.js, включая сценарии записи в лист
python scripts/make-pane-test.py
#    затем в headless Chrome: /test-office.html (режимы ?mock=blank, ?mock=big, ?mock=nosync)

# 7. Панель реально открылась внутри Excel (по логу моста)
grep -a "taskpane.html" workspace/bridge.log
#    ожидаем: GET /taskpane.html?_host_Info=Excel$Win32$16.01$ru-RU$$$$0 → 200,
#    затем /taskpane.css, /taskpane.js, /models, /sessions
```

Если шаг 7 пуст — панель в Excel ещё не открывалась (или Excel запущен до установки):
открой `workspace/hermes-auto.xlsx` (он поднимет панель сам) либо перезапусти Excel.

## Структура репозитория

| Путь | Что это |
|---|---|
| `bridge/bridge.py` | HTTPS+SSE мост: `/chat`, `/models`, `/sessions`, `/stop`, `/open`, `/health`; статика панели с того же origin |
| `addin/manifest.xml` | манифест надстройки (Id, иконки, лента, SourceLocation) |
| `addin/taskpane.html/.css/.js` | сама панель-чат: контекст выделения, стрим ответа, запись в лист |
| `addin/mock-office.js` | строгий мок Office.js для оффлайн-тестов |
| `scripts/make-pane-test.py` | генерирует `addin/test-office.html` из настоящего `taskpane.html` |
| `scripts/auto-open-workbook.py` | книга, которая сама открывает панель (`create` / `tag`) |
| `scripts/sideload.ps1` | регистрация в реестре + сборка авто-открывающей книги |
| `scripts/install.ps1` / `install.cmd` | установка одной командой (идемпотентно), `-DryRun`, `-Profile`, `-NoBridge` |
| `scripts/uninstall.ps1` | снятие всего, что поставил install (`-Profile` — удалить и бота, `-Purge` — и workspace, `-DryRun`) |
| `scripts/bridge-watchdog.ps1` | сторож: поднимает мост, если тот не отвечает (раз в минуту + при входе в Windows) |
| `scripts/install-watchdog.ps1` | ставит/снимает сторожа, `-Status` |
| `scripts/doctor.ps1` / `doctor.cmd` | проверка надстройки одной командой (11 пунктов), `-Json`, `-Fix` |
| `scripts/test-bridge-security.py` | регрессии моста: обход каталога, инъекция, CSRF, адресный стоп (13 проверок) |
| `scripts/check-workbook.py` | книга авто-открытия: части webextensions и совпадение Id с манифестом |
| `scripts/bench-turn.py`, `scripts/bench-bot.py` | замер хода и разбор прогона по событиям |
| `scripts/autostart-install.ps1` | (устарело) автозапуск ярлыком — заменён сторожем |
| `scripts/make-icons.py` | иконки 16/32/64/80 (нужен Pillow) |
| `start-bridge.cmd` | `start` (по умолчанию), `stop`, `status`, `excel` (мост + Excel с авто-открывающей книгой) |
| `.github/workflows/ci.yml` | CI: манифест, BOM у `.ps1`, синтаксис Python, скан секретов, сверка Id книги, защиты моста по AST |
| `workspace/` | рабочая папка агента и логи (`bridge.log`, `watchdog.log`) (в git не попадает) |

## Грабли (все встречены вживую — не наступай снова)

1. **Сайлоад на Windows — это `HKCU\Software\Microsoft\Office\16.0\WEF\Developer`**: значение
   `<Id надстройки>` = полный путь к манифесту. Доверенный каталог (`TrustedCatalogs`) на проверенной
   машине Office игнорировал (и с локальным путём, и с UNC `\\localhost\C$\...`, и после перезапуска).
   Не трать часы на каталог — сначала проверь `WEF\Developer`.
2. **Надстройка не появляется на ленте, пока её не активируют один раз в процессе Excel.** Отсюда
   книга с авто-открытием (`scripts/auto-open-workbook.py`), у неё `store="developer" storeType="Registry"`.
3. **`DefaultSettings/RequestedWidth` — невалидный элемент** для TaskPaneApp в схеме 1.1. Проверяй
   манифест `office-addin-manifest validate` перед сайлоадом.
4. **Office.js: любое чтение свойства требует `load()` + `sync()` на ЭТОМ объекте.** `sel.getCell(0,0).address`
   бросает «Свойство "address" недоступно». Якорь берётся из уже загруженной строки:
   `sel.address.split(":")[0]`. Сначала грузи дешёвые `address,rowCount,columnCount`, потом `text,values`
   на обрезанном диапазоне — выделение целого столбца иначе подвешивает панель.
5. **Windows `SO_REUSEADDR` позволяет двум процессам слушать один порт** — устаревший мост молча
   отвечает части запросов. В `bridge.py` переиспользование адреса выключено, второй запуск падает
   с внятной ошибкой. Убивать зависшие процессы:
   `Get-NetTCPConnection -LocalPort 3443 -State Listen | Stop-Process`. `taskkill /F` в git-bash
   калечится (`//F`) и молча ничего не делает.
6. **PowerShell 5.1 и кириллица**: `.ps1` сохраняй в UTF-8 **с BOM**, иначе парсер ломается на русских
   строках. В репозитории так и сделано (`utf-8-sig`).
7. **`curl` на Windows = Schannel**: добавляй `--ssl-no-revoke`; проверять доверие сертификата лучше
   через PowerShell `Invoke-WebRequest`.
8. **Excel с открытым модальным диалогом отвергает COM** (`RPC_E_CALL_REJECTED`): сначала закрой диалог,
   потом `GetActiveObject('Excel.Application')` → `$wb.Save()` → `Quit()`. Принудительное убийство
   оставляет панель «Восстановление документов» при следующем запуске.
9. **Один ход = один процесс `hermes chat`** (старт 2–4 с). Промпт-кэш живёт через `--resume`
   (`cache_read` ≈ 16k), так что продолжение диалога дешёвое, но «живого» зеркала в desktop нет —
   для этого нужен `hermes serve` (JSON-RPC/WebSocket, порт 9119) вместо спавна процесса.
10. **Мост должен переподключаться, а не падать**: панель опрашивает `/models` с повторами; при
    перезапуске моста она сама поднимется. Не делай одиночный fetch на старте — панель останется
    «мост недоступен» навсегда.
11. **Пустое выделение — частый реальный случай.** Мост получает `context.empty` и подсказывает модели:
    «данных нет, не задавай уточняющих вопросов, верни результат блоком ```csv». Без этого модель
    отвечает встречным вопросом, и пользователь думает, что «оно не работает».
12. **`python` из PATH на Windows — часто заглушка Microsoft Store** (`…\WindowsApps\python.exe`):
    печатает «Python» и выходит, мост не поднимается, в логе ошибок ровно строка `Python`. Ищите
    интерпретатор перебором кандидатов (venv Hermes → проект → PATH кроме `WindowsApps` → `py -3`)
    и **проверяйте кандидата запуском**, а не наличием файла.
13. **Автозапуск ярлыком срабатывает один раз при входе в Windows.** Мост упал позже — панель навсегда
    против мёртвого сервера. Нужен сторож: задача планировщика раз в минуту (`schtasks /SC MINUTE /MO 1`),
    ярлык автозагрузки вызывает его же с `-Force` (мгновенный старт при входе).
14. **Не сторожите сторож именованным мьютексом.** Один зависший запуск (у нас — `Start-Process` с
    `-RedirectStandardOutput` на процессе, который не отдал поток) держал мьютекс, и все следующие
    запуски молча выходили: лог пуст, мост лежит. Замок-файл со «старением» (3 минуты) самовосстанавливается.
    И не перенаправляйте потоки дочернего процесса — пусть мост пишет свой лог (`workspace/bridge.log`).
15. **Кириллица в `curl -d` из git-bash уезжает в cp1251** (`\xf1\xf3\xec\xec\xfb` = «суммы»),
    `json.loads` в мосте падает, и при `except: return {}` прогон стартовал с ПУСТЫМ промптом —
    агент честно отвечал «данных нет», а виноват был тест. Мост теперь отвечает `400` и пишет в лог
    первые байты; тела для проверок пишите в файл и передавайте `--data-binary @file.json`.
16. **`write_file` сам добавляет UTF-8 BOM для `.ps1`.** Если добавить BOM ещё раз, PowerShell 5.1
    падает на первой строке (`?# : не является именем командлета`). Нормализация: снять все ведущие BOM
    и добавить ровно один (utf-8-sig + CRLF).
17. **PowerShell 5.1 не видит тип `System.IO.Compression.ZipArchiveMode`** (не резолвится даже после
    `Add-Type`), а `ZipFile::OpenRead` не открывает `.xlsx`, который держит Excel. Разбор архива отдайте
    Python (`scripts/check-workbook.py`). И выставляйте `[Console]::OutputEncoding = UTF8`, иначе
    кириллица в консоли и в `-Json` уходит в cp866 и разбирается криво.
18. **Интерактивные подтверждения вешают агента.** `hermes profile delete <имя>` без `-y` спрашивает
    «Type '<имя>' to confirm» и ждёт ввода — в скриптах добавляйте `-y`, а после удаления
    `hermes profile purge-identity <имя>` (иначе остаётся «хвост» сессий/маршрутизации).

## Если `git push` не проходит (TLS/VPN)

На машине с VPN или DPI-фильтрацией `git push` может рваться на TLS — `schannel: failed to receive
handshake` либо `unexpected eof while reading` — хотя обычные запросы к `api.github.com` через `gh`
работают. Залить изменения без git-транспорта:

```bash
git add -A && git commit -m "…"
python scripts/push_via_api.py --repo <owner>/<name>      # коммит собирается через Git Data API
```

Скрипт берёт байты **из индекса git** (`git cat-file blob :path`), то есть ровно то, что отправил бы
обычный push, а не то, что лежит на диске; умеет завести первый коммит в пустом репозитории
(Contents API, иначе Git Data API отвечает `409 Git Repository is empty`) и повторяет запросы при
обрывах. Проверка, что всё уехало ровно так, как в индексе:

```bash
git ls-files -s | while read mode sha stage path; do
  remote=$(gh api "repos/<owner>/<name>/git/trees/main?recursive=1" \
    --jq ".tree[] | select(.path==\"$path\") | .sha")
  [ "$sha" = "$remote" ] && echo "ok $path" || echo "РАСХОЖДЕНИЕ $path"
done
```

## Как менять панель

1. Правь `addin/taskpane.js|css|html` (статика отдаётся с `Cache-Control: no-store`, перезагрузки хватает).
2. Прогони оффлайн-тест: `python scripts/make-pane-test.py` + headless Chrome (шаг 5 проверки).
   **Проверь сам тест**: временно верни баг — тест обязан покраснеть с тем же текстом ошибки Office.
3. Живая проверка: перезапусти панель в Excel (кнопка «Hermes» на «Главной») или открой `workspace/hermes-auto.xlsx`.

## Удаление

```bash
scripts\uninstall.ps1                 # снимает задачу, ярлык, запись в реестре, останавливает мост
scripts\uninstall.ps1 -Profile excel  # плюс бот вместе с его сессиями
scripts\uninstall.ps1 -Purge          # плюс workspace (книги, логи, файлы агента)
scripts\uninstall.ps1 -DryRun         # показать план, ничего не трогая
```

Идемпотентно: повторный запуск ничего не ломает. Сертификат `https://localhost` не удаляется — им могут
пользоваться другие надстройки. После этого повторная установка — `install.cmd`.
