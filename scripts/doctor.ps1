# Доктор «Hermes для Excel»: одна команда вместо ручного обхода семи шагов проверки.
#
# Запуск:  powershell -ExecutionPolicy Bypass -File scripts\doctor.ps1
#          powershell -ExecutionPolicy Bypass -File scripts\doctor.ps1 -Json      (машиночитаемо)
#          powershell -ExecutionPolicy Bypass -File scripts\doctor.ps1 -Fix       (поднять мост, если лежит)
#
# Каждая строка — PASS/WARN/FAIL и готовая команда-исправление. Код возврата 1, если есть FAIL.

param(
    [switch]$Json,
    [switch]$Fix,
    [int]$Port = 3443
)

$ErrorActionPreference = 'Continue'
# Кириллица в консоли: без этого и таблица, и -Json выходят в cp866 и разбираются криво
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path -Parent $PSScriptRoot
$rows = @()

function Add-Row([string]$name, [string]$status, [string]$detail, [string]$fix) {
    $script:rows += [pscustomobject]@{ check = $name; status = $status; detail = $detail; fix = $fix }
}

function Test-Bridge([int]$p) {
    try {
        $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 6 "https://localhost:$p/health"
        if ($r.StatusCode -eq 200) { return $r.Content } else { return $null }
    } catch { return $null }
}

# 1. сертификаты для https://localhost
$certs = Join-Path $env:USERPROFILE '.office-addin-dev-certs'
$haveCrt = (Test-Path (Join-Path $certs 'localhost.crt')) -and (Test-Path (Join-Path $certs 'localhost.key'))
Add-Row 'сертификаты https://localhost' ($(if ($haveCrt) { 'PASS' } else { 'FAIL' })) `
    $(if ($haveCrt) { $certs } else { 'нет localhost.crt/localhost.key' }) `
    'npx --yes office-addin-dev-certs install --days 365'

# 2. мост
$health = Test-Bridge $Port
if (-not $health -and $Fix) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'bridge-watchdog.ps1') -Force | Out-Null
    Start-Sleep -Seconds 6
    $health = Test-Bridge $Port
}
$profileName = ''
if ($health) {
    try { $profileName = ($health | ConvertFrom-Json).profile } catch { }
}
Add-Row "мост на https://localhost:$Port" ($(if ($health) { 'PASS' } else { 'FAIL' })) `
    $(if ($health) { "профиль: $profileName" } else { 'не отвечает на /health' }) `
    'start-bridge.cmd   (или: powershell -File scripts\bridge-watchdog.ps1 -Force)'

# 3. бот (профиль) со своей моделью
$botHome = Join-Path $env:LOCALAPPDATA "hermes\profiles\$profileName"
$botOk = $profileName -and (Test-Path (Join-Path $botHome 'state.db'))
Add-Row "бот-профиль '$profileName'" ($(if ($botOk) { 'PASS' } else { 'FAIL' })) `
    $(if ($botOk) { $botHome } else { 'нет state.db у профиля' }) `
    "hermes profile create $profileName --clone"

# 4. сторож: задача планировщика (автозапуск при входе намеренно отключён)
$task = Get-ScheduledTask -TaskName 'Hermes Excel Bridge Watchdog' -ErrorAction SilentlyContinue
if ($task) {
    $info = $task | Get-ScheduledTaskInfo
    $st = if ($info.LastTaskResult -eq 0) { 'PASS' } else { 'WARN' }
    Add-Row 'сторож: задача раз в минуту' $st `
        ("состояние {0}, последний результат 0x{1:X}" -f $task.State, $info.LastTaskResult) `
        'powershell -File scripts\install-watchdog.ps1'
} else {
    Add-Row 'сторож: задача раз в минуту' 'WARN' 'задачи нет — мост не поднимется после падения' `
        'powershell -File scripts\install-watchdog.ps1'
}
$lnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'Hermes Excel Bridge.lnk'
Add-Row 'сторож: автозапуск при входе' ($(if (-not (Test-Path $lnk)) { 'PASS' } else { 'WARN' })) `
    $(if (-not (Test-Path $lnk)) { 'отключен — мост ждёт Excel' } else { 'старый ярлык запускает мост независимо от Excel' }) `
    'powershell -ExecutionPolicy Bypass -File scripts\install-watchdog.ps1'

# 5. сайлоад: запись в реестре указывает на существующий манифест с тем же Id
$manifest = Join-Path $root 'addin\manifest.xml'
$addinId = ''
if (Test-Path $manifest) {
    $m = Select-String -Path $manifest -Pattern '<Id>([0-9a-fA-F-]{36})</Id>' | Select-Object -First 1
    if ($m) { $addinId = $m.Matches[0].Groups[1].Value }
}
$reg = $null
try {
    $reg = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer' -ErrorAction Stop).$addinId
} catch { }
$sideloadOk = $reg -and (Test-Path $reg)
Add-Row 'надстройка зарегистрирована в Excel' ($(if ($sideloadOk) { 'PASS' } else { 'FAIL' })) `
    $(if ($sideloadOk) { "$addinId -> $reg" } else { "нет значения $addinId в WEF\Developer" }) `
    'powershell -ExecutionPolicy Bypass -File scripts\sideload.ps1'

# 6. книга, которая сама открывает панель
$wb = Join-Path $root 'workspace\hermes-auto.xlsx'
# Проверку книги делает Python: PowerShell 5.1 не видит тип ZipArchiveMode, а Python читает
# .xlsx даже когда файл открыт в Excel.
$wbOk = $false
$wbDetail = 'книги нет'
if (Test-Path $wb) {
    $pyForCheck = (Get-Command python -ErrorAction SilentlyContinue).Source
    if ($pyForCheck) {
        $out = & $pyForCheck (Join-Path $PSScriptRoot 'check-workbook.py') $wb 2>&1
        $wbOk = ($LASTEXITCODE -eq 0)
        $wbDetail = ($out | Select-Object -First 1)
    } else {
        $wbDetail = 'python не найден — проверить книгу нечем'
    }
}
Add-Row 'книга с авто-открытием панели' ($(if ($wbOk) { 'PASS' } else { 'WARN' })) $wbDetail `
    'python scriptsuto-open-workbook.py create'

# 7. панель действительно открывалась внутри Excel (по логу моста)
$log = Join-Path $root 'workspace\bridge.log'
$paneSeen = $false
if (Test-Path $log) {
    $paneSeen = (Select-String -Path $log -Pattern 'taskpane.html' -Quiet -ErrorAction SilentlyContinue)
}
Add-Row 'панель открывалась в Excel' ($(if ($paneSeen) { 'PASS' } else { 'WARN' })) `
    $(if ($paneSeen) { 'в логе есть запрос taskpane.html' } else { 'в логе нет запроса панели' }) `
    'откройте Excel и панель Hermes (или книгу workspace\hermes-auto.xlsx)'

# 8. python для сторожа: заглушка WindowsApps не считается
$py = (Get-Command python -ErrorAction SilentlyContinue).Source
$pyOk = $py -and ($py -notlike '*WindowsApps*')
Add-Row 'python в PATH' ($(if ($pyOk) { 'PASS' } else { 'WARN' })) `
    $(if ($py) { $py } else { 'python не найден' }) `
    'сторож сам найдёт интерпретатор Hermes (venv) — правка не нужна'

# 9. все .ps1 в UTF-8 с BOM (иначе PowerShell 5.1 давится кириллицей)
$bad = @()
Get-ChildItem -Path (Join-Path $root 'scripts') -Filter *.ps1 | ForEach-Object {
    $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
    if (-not ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) {
        $bad += $_.Name
    }
}
Add-Row 'кодировка .ps1 (UTF-8 BOM)' ($(if ($bad.Count -eq 0) { 'PASS' } else { 'FAIL' })) `
    $(if ($bad.Count -eq 0) { 'все скрипты с BOM' } else { 'без BOM: ' + ($bad -join ', ') }) `
    'python -c "перезапись с utf-8-sig"'

# 10. Excel запущен?
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
Add-Row 'Excel запущен' ($(if ($excel) { 'PASS' } else { 'WARN' })) `
    $(if ($excel) { "pid $($excel.Id -join ', ')" } else { 'Excel не запущен — панель некуда показать' }) `
    'start excel'

if ($Json) {
    $rows | ConvertTo-Json -Depth 4
} else {
    Write-Host ''
    Write-Host ("{0,-32} {1,-5} {2}" -f 'ПРОВЕРКА', 'ИТОГ', 'ЧТО ИМЕННО')
    Write-Host ('-' * 100)
    foreach ($r in $rows) {
        $color = switch ($r.status) { 'PASS' { 'Green' } 'WARN' { 'Yellow' } default { 'Red' } }
        Write-Host ("{0,-32} " -f $r.check) -NoNewline
        Write-Host ("{0,-5} " -f $r.status) -NoNewline -ForegroundColor $color
        Write-Host $r.detail
        if ($r.status -ne 'PASS' -and $r.fix) { Write-Host ("{0,-39}→ {1}" -f '', $r.fix) -ForegroundColor DarkGray }
    }
    $fails = ($rows | Where-Object { $_.status -eq 'FAIL' }).Count
    $warns = ($rows | Where-Object { $_.status -eq 'WARN' }).Count
    Write-Host ('-' * 100)
    Write-Host ("итог: {0} проверок, провалено {1}, предупреждений {2}" -f $rows.Count, $fails, $warns)
    if ($fails -or $warns) { Write-Host 'команды-исправления указаны в строках ниже соответствующих проверок' }
}

if (($rows | Where-Object { $_.status -eq 'FAIL' }).Count -gt 0) { exit 1 }
exit 0
