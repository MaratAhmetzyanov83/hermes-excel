# Установка надстройки «Hermes для Excel» одной командой. Идемпотентно: повторный запуск ничего не ломает.
#
# Запуск (из корня клонированного репозитория):
#   powershell -ExecutionPolicy Bypass -File scripts\install.ps1
#   install.cmd                       — то же двойным кликом
#   install.cmd -DryRun               — показать, что будет сделано, ничего не меняя
#   install.cmd -Profile excel2 -Port 3443
#
# Что делает по шагам (каждый шаг сначала проверяет, не сделано ли уже):
#   0. окружение: Windows, Excel, python (мимо заглушки WindowsApps), hermes
#   1. сертификаты для https://localhost (npx office-addin-dev-certs → openssl из git-bash → уже есть)
#   2. бот-профиль Hermes со своими сессиями + модель по умолчанию
#   3. регистрация надстройки в Excel + книга, которая сама открывает панель
#   4. сторож моста (задача раз в минуту + автозапуск) и поднятие моста
#   5. проверка: doctor + живой ход бота (чтобы «настроено» не значило «молчит»)
#
# Агенту на новой машине: клонировать репозиторий и выполнить `install.cmd`, затем `doctor.cmd`.
# Никаких pip/npm-зависимостей у моста нет; Node нужен только как один из способов выдать сертификат.

param(
    [string]$Profile = 'excel',
    [int]$Port = 3443,
    [switch]$DryRun,
    [switch]$NoBridge,          # не ставить сторожа и не поднимать мост
    [switch]$SkipCert,          # пропустить шаг сертификатов (если вы ими управляете сами)
    [string]$CaDir = (Join-Path $env:USERPROFILE '.office-addin-dev-certs')
)

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path -Parent $PSScriptRoot
$steps = @()
$fails = 0

function Say([string]$text, [string]$color = 'Gray') { Write-Host $text -ForegroundColor $color }
function Step([string]$name, [string]$status, [string]$detail) {
    $script:steps += [pscustomobject]@{ step = $name; status = $status; detail = $detail }
    $c = switch ($status) { 'ok' { 'Green' } 'skip' { 'DarkGray' } 'warn' { 'Yellow' } default { 'Red' } }
    Write-Host ("[{0}] {1,-44} {2}" -f $status.ToUpper(), $name, $detail) -ForegroundColor $c
    if ($status -eq 'fail') { $script:fails++ }
}
function Run([string]$exe, [string[]]$argv, [switch]$Show) {
    # ВАЖНО: возвращаем КОД, а не текст. Без захвата вывода PowerShell добавляет вывод команды
    # в возвращаемое значение функции, и «успех» выглядел как «вернул 0 … создан профиль …»
    if ($DryRun) { Say ("      (dry-run) " + $exe + ' ' + ($argv -join ' ')) 'DarkGray'; return 0 }
    $out = & $exe @argv 2>&1 | Out-String
    $code = [int]$LASTEXITCODE
    if ($Show -or $code -ne 0) { foreach ($l in ($out -split "`n" | Where-Object { $_.Trim() })) { Say ("      | " + $l.Trim()) 'DarkGray' } }
    return $code
}

Say '=== Hermes для Excel: установка ===' 'White'
Say ("репозиторий: $root")
Say ("профиль бота: $Profile | порт моста: $Port" + $(if ($DryRun) { ' | РЕЖИМ ПРОВЕРКИ (ничего не меняется)' } else { '' }))
Say ''

# ---------------------------------------------------------------- 0. окружение
if (-not $IsWindows -and $env:OS -notlike '*Windows*') {
    Step 'платформа' 'fail' 'нужна Windows (Excel desktop + реестровый сайлоад)'
    exit 1
}

$excel = @(
    (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)',
    (Join-Path ${env:ProgramFiles} 'Microsoft Office\root\Office16\EXCEL.EXE'),
    (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\EXCEL.EXE')
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
Step 'Excel (desktop)' $(if ($excel) { 'ok' } else { 'fail' }) $(if ($excel) { $excel } else { 'EXCEL.EXE не найден' })

function Resolve-Python {
    # В PATH на Windows часто первым стоит заглушка Microsoft Store: печатает «Python» и выходит.
    $cands = @()
    $cands += (Join-Path $env:LOCALAPPDATA 'hermes\hermes-agent\venv\Scripts\python.exe')
    $cands += (Join-Path $env:USERPROFILE 'hermes-excel\.venv\Scripts\python.exe')
    $fromPath = (Get-Command python -ErrorAction SilentlyContinue).Source
    if ($fromPath -and $fromPath -notlike '*WindowsApps*') { $cands += $fromPath }
    foreach ($c in $cands) {
        if (-not (Test-Path $c)) { continue }
        try {
            $v = & $c -c "import sys; print(sys.version.split()[0])" 2>$null
            if ($LASTEXITCODE -eq 0 -and $v) { return @{ exe = $c; ver = $v } }
        } catch { }
    }
    $py = (Get-Command py -ErrorAction SilentlyContinue).Source
    if ($py) {
        try {
            $v = & $py -3 -c "import sys; print(sys.version.split()[0])" 2>$null
            if ($LASTEXITCODE -eq 0 -and $v) { return @{ exe = $py; ver = $v; pre = @('-3') } }
        } catch { }
    }
    return $null
}

$python = Resolve-Python
Step 'python (мост и скрипты)' $(if ($python) { 'ok' } else { 'fail' }) `
    $(if ($python) { "$($python.exe) ($($python.ver))" } else { 'нет рабочего python; заглушка WindowsApps не считается' })

$hermes = (Get-Command hermes -ErrorAction SilentlyContinue).Source
if (-not $hermes) {
    $hermes = @(
        (Join-Path $env:LOCALAPPDATA 'hermes\bin\hermes.exe'),
        (Join-Path $env:LOCALAPPDATA 'hermes\hermes-agent\venv\Scripts\hermes.exe')
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}
Step 'hermes (агент внутри Excel)' $(if ($hermes) { 'ok' } else { 'fail' }) `
    $(if ($hermes) { $hermes } else { 'не найден: поставьте Hermes Agent и настройте модель' })

if ($fails -gt 0) {
    Say ''
    Say 'Не хватает обязательного — почините строки FAIL и запустите снова.' 'Red'
    exit 1
}

# ---------------------------------------------------------------- 1. сертификаты
$crt = Join-Path $CaDir 'localhost.crt'
$key = Join-Path $CaDir 'localhost.key'
function Test-Trusted([string]$crtPath) {
    # Проверяем ЦЕПОЧКУ: office-addin-dev-certs добавляет в доверенные корневой CA,
    # а сам лист лишь подписан им — поиск листа в CurrentUser\Root давал ложное «не доверен»
    # и заставлял переустанавливать сертификаты при каждом запуске.
    try {
        $c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($crtPath)
        $chain = New-Object System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.RevocationMode = 'NoCheck'      # локальный CA без CRL (та же причина, что у curl)
        return $chain.Build($c)
    } catch { return $false }
}

if ($SkipCert) {
    Step 'сертификат https://localhost' 'skip' 'пропущено по -SkipCert'
} elseif ((Test-Path $crt) -and (Test-Path $key) -and (Test-Trusted $crt)) {
    Step 'сертификат https://localhost' 'skip' 'уже есть и доверенный'
} else {
    $made = $false
    $npx = (Get-Command npx -ErrorAction SilentlyContinue).Source
    $opensslPath = (Get-Command openssl -ErrorAction SilentlyContinue).Source
    if ($npx) {
        Say '      генерирую сертификат через office-addin-dev-certs (npx)…'
        $code = Run $npx @('--yes', 'office-addin-dev-certs', 'install', '--days', '365')
        $made = ($code -eq 0) -and (Test-Path $crt) -and (Test-Path $key)
    }
    if (-not $made) {
        # Без Node: openssl из git-bash + импорт сертификата в доверенные корни текущего пользователя.
        $openssl = $opensslPath
        if ($openssl) {
            Say '      Node нет — генерирую самоподписанный сертификат через openssl…'
            if (-not (Test-Path $CaDir)) { New-Item -ItemType Directory -Path $CaDir -Force | Out-Null }
            # os=Windows_NT не даёт openssl подсунуть конфиг из MSYS; конфиг пишем сами (SAN обязателен)
            $cfg = Join-Path $CaDir 'openssl-localhost.cnf'
            @"
[req]
distinguished_name = dn
x509_extensions = v3
prompt = no
[dn]
CN = localhost
[v3]
subjectAltName = DNS:localhost,IP:127.0.0.1
basicConstraints = critical,CA:TRUE
keyUsage = critical,digitalSignature,keyCertSign
extendedKeyUsage = serverAuth
"@ | Set-Content -Path $cfg -Encoding ASCII
            if (-not $DryRun) {
                & $openssl req -x509 -newkey rsa:2048 -nodes -days 3650 -keyout $key -out $crt -config $cfg 2>$null
            }
            if ((Test-Path $crt) -and (Test-Path $key)) {
                if (-not $DryRun) {
                    Import-Certificate -FilePath $crt -CertStoreLocation 'Cert:\CurrentUser\Root' | Out-Null
                }
                $made = $true
            }
        }
    }
    if ($DryRun) {
        $how = if ($npx) { 'npx office-addin-dev-certs' } elseif ($openssl) { 'openssl' } else { 'нечем' }
        Step 'сертификат https://localhost' 'skip' "(dry-run) выдал бы через $how"
    } elseif ($made) {
        Step 'сертификат https://localhost' 'ok' $(if (Test-Trusted $crt) { 'создан, цепочка доверенная' } else { 'создан, но не в доверенных корнях' })
    } else {
        Step 'сертификат https://localhost' 'fail' `
            'нет ни npx, ни openssl. Поставьте Node (winget install OpenJS.NodeJS.LTS) или git-bash'
    }
}

# ---------------------------------------------------------------- 2. бот-профиль
$profilesJson = ''
if (-not $DryRun) {
    $profilesJson = (& $hermes profile list 2>&1 | Out-String)
}
$profileExists = $profilesJson -match "(^|\s)$([regex]::Escape($Profile))(\s|$)" -or (Test-Path (Join-Path $env:LOCALAPPDATA "hermes\profiles\$Profile"))
if ($profileExists) {
    Step "бот-профиль '$Profile'" 'skip' 'уже существует — сессии и память не трогаю'
} else {
    Say "      создаю профиль '$Profile' клоном основного (--clone переносит конфиг и ключи)…"
    $code = Run $hermes @('profile', 'create', $Profile, '--clone',
                          '--description', 'Excel-бот: живёт в надстройке Excel')
    if ($code -eq 0 -or $DryRun) {
        Step "бот-профиль '$Profile'" 'ok' 'создан'
    } else {
        Step "бот-профиль '$Profile'" 'fail' "hermes profile create вернул $code"
    }
}

# ---------------------------------------------------------------- 3. регистрация + книга
$manifest = Join-Path $root 'addin\manifest.xml'
$addinId = ''
if (Test-Path $manifest) {
    $m = Select-String -Path $manifest -Pattern '<Id>([0-9a-fA-F-]{36})</Id>' | Select-Object -First 1
    if ($m) { $addinId = $m.Matches[0].Groups[1].Value }
}
$reg = $null
try { $reg = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer' -ErrorAction Stop).$addinId } catch { }
if ($reg -and (Test-Path $reg)) {
    Step 'надстройка встроена в Excel' 'skip' "уже зарегистрирована ($addinId)"
} else {
    Say '      регистрирую надстройку и собираю книгу с авто-открытием панели…'
    $code = Run 'powershell' @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'sideload.ps1'))
    if ($code -eq 0 -or $DryRun) {
        Step 'надстройка встроена в Excel' 'ok' "Id $addinId → addin\manifest.xml"
    } else {
        Step 'надстройка встроена в Excel' 'fail' "sideload.ps1 вернул $code"
    }
}

# ---------------------------------------------------------------- 4. сторож и мост
if ($NoBridge) {
    Step 'сторож моста' 'skip' 'пропущено по -NoBridge'
} else {
    Say '      ставлю сторожа (задача раз в минуту + автозапуск) и поднимаю мост…'
    $argv = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'install-watchdog.ps1'))
    $code = Run 'powershell' $argv
    $health = $null
    try { $health = (Invoke-WebRequest -UseBasicParsing -TimeoutSec 8 "https://localhost:$Port/health").Content } catch { }
    Step 'сторож моста и запуск' $(if ($health) { 'ok' } else { 'warn' }) `
        $(if ($health) { 'мост отвечает на /health' } else { 'мост пока не отвечает — смотрите workspace\bridge.log' })
}

# ---------------------------------------------------------------- 5. проверка «а он живой?»
$botOk = $null
$botDetail = 'не проверял'
if (-not $DryRun) {
    $q = Join-Path $env:TEMP 'hermes-install-check.txt'
    Set-Content -Path $q -Value 'Ответь одним словом: ок' -Encoding UTF8
    $flagProfile = @()
    if ($Profile -and $Profile -ne 'default') { $flagProfile = @('-p', $Profile) }
    try {
        $out = & $hermes @flagProfile chat -Q --format stream-json --query-file $q --max-turns 2 2>&1 | Out-String
        $res = ($out -split "`n" | Where-Object { $_ -match '"type":\s*"result"' } | Select-Object -Last 1)
        if ($res) {
            $json = $res | ConvertFrom-Json
            if ($json.exit_code -eq 0) { $botOk = $true; $botDetail = "модель ответила за $($json.duration_ms) мс" }
            else { $botOk = $false; $botDetail = ('сбой провайдера: ' + (($json.text -split "`n")[0])) }
        } else { $botOk = $false; $botDetail = 'нет события result в ответе' }
    } catch { $botOk = $false; $botDetail = $_.Exception.Message }
    Remove-Item $q -ErrorAction SilentlyContinue
}
Step 'бот отвечает' $(if ($botOk -eq $true) { 'ok' } elseif ($botOk -eq $false) { 'warn' } else { 'skip' }) $botDetail
if ($botOk -eq $false) {
    Say "      проверьте модель и авторизацию: hermes -p $Profile config get model" 'Yellow'
    Say "      и при необходимости: hermes -p $Profile auth add <провайдер> --type oauth" 'Yellow'
}

# ---------------------------------------------------------------- итог
if (-not $DryRun) {
    Say ''
    Say '=== проверка всей надстройки ===' 'White'
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'doctor.ps1')
}

Say ''
Say '=== что делать дальше ===' 'White'
Say '  1. Открыть Excel и панель Hermes (или книгу workspace\hermes-auto.xlsx — панель откроется сама).'
Say '  2. Проверить: start-bridge.cmd status   (мост + сторож)'
Say '  3. Если что-то не так: doctor.cmd        (таблица проверок с командами-исправлениями)'
Say ''
if ($fails -gt 0) { Say "не выполнено шагов: $fails" 'Red'; exit 1 }
Say 'установка завершена' 'Green'
exit 0
