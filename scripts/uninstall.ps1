# Удаление надстройки «Hermes для Excel»: снимает всё, что поставил install.ps1. Идемпотентно.
#
# Запуск:  powershell -ExecutionPolicy Bypass -File scripts\uninstall.ps1
#          ... -Profile excel     — ещё и удалить бота вместе с его сессиями
#          ... -Purge             — ещё и очистить workspace (книги, логи, файлы агента)
#          ... -KeepBridge        — не останавливать мост (например, им пользуетесь сейчас)
#          ... -DryRun            — показать план, ничего не снимая
#
# Что снимается: задача планировщика + ярлык автозагрузки, значение надстройки в
# HKCU\...\WEF\Developer, ключ доверенного каталога (если он был), мост, при флагах — профиль и workspace.
# Сертификат https://localhost не трогаем: им могут пользоваться другие надстройки.

param(
    [string]$Profile = '',
    [switch]$Purge,
    [switch]$KeepBridge,
    [switch]$DryRun,          # показать, что будет снято, ничего не трогая
    [int]$Port = 3443
)

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$root = Split-Path -Parent $PSScriptRoot
$done = @()
function Note([string]$what) {
    $script:done += $what
    if ($DryRun) { Write-Host "  [план] снял бы: $what" -ForegroundColor DarkGray }
    else { Write-Host "  снято: $what" -ForegroundColor Gray }
}
function NoteSkip([string]$what) { Write-Host "  нечего снимать: $what" -ForegroundColor DarkGray }

Write-Host '=== Hermes для Excel: удаление ===' -ForegroundColor White

# 1. сторож: задача + ярлык автозагрузки
if (-not $DryRun) { & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'install-watchdog.ps1') -Remove | Out-Null }
Note 'задача планировщика и ярлык автозагрузки'

# 2. мост
if ($KeepBridge) {
    NoteSkip 'мост (по -KeepBridge)'
} else {
    $pids = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty OwningProcess | Sort-Object -Unique
    if ($pids) {
        if (-not $DryRun) { foreach ($p in $pids) { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue } }
        Note "мост (pid $($pids -join ', '))"
    } else { NoteSkip 'мост' }
}

# 3. регистрация надстройки
$manifest = Join-Path $root 'addin\manifest.xml'
$addinId = ''
if (Test-Path $manifest) {
    $m = Select-String -Path $manifest -Pattern '<Id>([0-9a-fA-F-]{36})</Id>' | Select-Object -First 1
    if ($m) { $addinId = $m.Matches[0].Groups[1].Value }
}
$key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
$has = $false
try {
    $props = (Get-ItemProperty $key -ErrorAction Stop).PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }
    if ($addinId -and ($props.Name -contains $addinId)) {
        if (-not $DryRun) { Remove-ItemProperty -Path $key -Name $addinId -Force }
        $has = $true
    }
    if (-not $props -or $props.Count -eq 0) { }        # ключ оставляем: им пользуются другие надстройки
} catch { }
if ($has) { Note "значение надстройки в WEF\Developer ($addinId)" } else { NoteSkip 'значение надстройки в реестре' }

foreach ($cat in @('{9F5D2C41-7A83-4E1B-9C0D-1E2F3A4B5C6D}', '{3C7E9A21-5D44-4B18-9E77-6F1A2B3C4D5E}')) {
    $p = "HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\$cat"
    if (Test-Path $p) { if (-not $DryRun) { Remove-Item -Recurse -Force $p -ErrorAction SilentlyContinue }; Note "ключ каталога $cat" }
}

# 4. бот-профиль (только по явному флагу: там сессии и память)
if ($Profile) {
    $home = Join-Path $env:LOCALAPPDATA "hermes\profiles\$Profile"
    if (Test-Path $home) {
        $hermes = (Get-Command hermes -ErrorAction SilentlyContinue).Source
        if (-not $DryRun) {
            # -y обязателен: без него hermes спрашивает подтверждение и ждёт ввода (агент повиснет).
            # purge-identity добивает «хвост» профиля (сессии/маршрутизацию), о котором Hermes предупреждает.
            if ($hermes) {
                & $hermes profile delete $Profile -y 2>&1 | Out-Null
                & $hermes profile purge-identity $Profile 2>&1 | Out-Null
            }
            if (Test-Path $home) { Remove-Item -Recurse -Force $home -ErrorAction SilentlyContinue }
        }
        Note "бот-профиль '$Profile' (вместе с его сессиями)"
    } else { NoteSkip "бот-профиль '$Profile'" }
} else {
    NoteSkip 'бот-профиль (укажите -Profile, если нужно удалить и его)'
}

# 5. рабочая папка
if ($Purge) {
    $ws = Join-Path $root 'workspace'
    if (Test-Path $ws) { if (-not $DryRun) { Remove-Item -Recurse -Force $ws -ErrorAction SilentlyContinue }; Note 'папка workspace (книги, логи, файлы агента)' }
} else {
    NoteSkip 'workspace (укажите -Purge, чтобы удалить книги и логи)'
}

Write-Host ''
Write-Host 'Сертификат https://localhost оставлен: им могут пользоваться другие надстройки.' -ForegroundColor DarkGray
Write-Host 'Повторная установка: install.cmd' -ForegroundColor Green
