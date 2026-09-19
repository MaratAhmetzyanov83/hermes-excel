# AGENTS.md — установка и работа с надстройкой «Hermes для Excel»

Этот файл — инструкция для ИИ-агента (Hermes, Claude Code, Codex и т.п.), который ставит надстройку
на машину пользователя. Выполняй шаги по порядку и **проверяй каждый** — процедура ниже проверена
на живой Windows-машине, включая все грабли из раздела «Грабли».

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
curl -sN --ssl-no-revoke -X POST -H "Content-Type: application/json" \
  -d '{"prompt":"Сколько строк данных?","model":"auto","context":{"csv":"a,b\n1,2\n3,4"}}' \
  https://localhost:3443/chat | grep -E "^event: (result|done)" -A1
hermes -p excel sessions list --limit 3     # та же сессия видна в сторе бота

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
| `scripts/autostart-install.ps1` | (опц.) запуск моста при входе в Windows |
| `scripts/make-icons.py` | иконки 16/32/64/80 (нужен Pillow) |
| `start-bridge.cmd` | `start-bridge.cmd` — старт, `start-bridge.cmd stop` — остановка |
| `workspace/` | рабочая папка агента и логи (в git не попадает) |

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

## Как менять панель

1. Правь `addin/taskpane.js|css|html` (статика отдаётся с `Cache-Control: no-store`, перезагрузки хватает).
2. Прогони оффлайн-тест: `python scripts/make-pane-test.py` + headless Chrome (шаг 5 проверки).
   **Проверь сам тест**: временно верни баг — тест обязан покраснеть с тем же текстом ошибки Office.
3. Живая проверка: перезапусти панель в Excel (кнопка «Hermes» на «Главной») или открой `workspace/hermes-auto.xlsx`.

## Удаление

```bash
powershell -NoProfile -Command "Remove-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer' -Name '<addin-id>'; Remove-Item -Recurse -Force 'HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\{9F5D2C41-7A83-4E1B-9C0D-1E2F3A4B5C6D}' -ErrorAction SilentlyContinue"
start-bridge.cmd stop
rm "$APPDATA/Microsoft/Windows/Start Menu/Programs/Startup/Hermes Excel Bridge.lnk"   # если ставили автозапуск
hermes profile delete excel                                                            # удалит бота и его сессии
```
