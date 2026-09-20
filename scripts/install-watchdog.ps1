# Ставит сторожа моста «Hermes для Excel»: задача планировщика (раз в минуту) + ярлык в «Автозагрузке».
# После этого надстройка работает сама: открыли Excel — мост уже поднят, панель подключается к нему.
#
# Запуск:  powershell -ExecutionPolicy Bypass -File scripts\install-watchdog.ps1
# Снять:   powershell -ExecutionPolicy Bypass -File scripts\install-watchdog.ps1 -Remove
# Статус:  powershell -ExecutionPolicy Bypass -File scripts\install-watchdog.ps1 -Status

param(
    [switch]$Remove,
    [switch]$Status
)

$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $PSScriptRoot
$task    = 'Hermes Excel Bridge Watchdog'
$watch   = Join-Path $root 'scripts\bridge-watchdog.ps1'
$startup = [Environment]::GetFolderPath('Startup')
$lnk     = Join-Path $startup 'Hermes Excel Bridge.lnk'

if ($Status) {
    Write-Host "задача:  " -NoNewline; schtasks /Query /TN $task /FO LIST 2>$null | Select-String 'Status|Last Run|Next Run|Состояние|Последнее|Следующее'
    Write-Host "ярлык:   " -NoNewline; if (Test-Path $lnk) { Write-Host $lnk } else { Write-Host 'нет' }
    Write-Host "журнал:  " -NoNewline; Write-Host (Join-Path $root 'workspace\watchdog.log')
    exit 0
}

if ($Remove) {
    schtasks /Delete /TN $task /F 2>$null | Out-Null
    Remove-Item $lnk -ErrorAction SilentlyContinue
    Write-Host 'Сторож и автозапуск сняты. Мост можно поднять вручную: start-bridge.cmd'
    exit 0
}

# 1. Ярлык в «Автозагрузке»: поднимает мост сразу при входе в Windows, окно не показывается
$shell = New-Object -ComObject WScript.Shell
$sc = $shell.CreateShortcut($lnk)
$sc.TargetPath       = 'powershell.exe'
$sc.Arguments        = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$watch`" -Force"
$sc.WorkingDirectory = $root
$sc.WindowStyle      = 7
$sc.Description      = 'Hermes Excel Bridge: поднять мост при входе в Windows'
$sc.Save()
Write-Host "ярлык автозапуска: $lnk"

# 2. Задача планировщика: раз в минуту проверяет мост и поднимает его, если тот упал
$action = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$watch`""
schtasks /Create /TN $task /TR $action /SC MINUTE /MO 1 /F | Out-Null
Write-Host "задача планировщика: $task (раз в минуту)"

# 3. Сразу проверяем, что всё живо
schtasks /Run /TN $task | Out-Null
Start-Sleep -Seconds 6
try {
    $h = (Invoke-WebRequest -UseBasicParsing -TimeoutSec 5 'https://localhost:3443/health').Content
    Write-Host "мост отвечает: $h"
} catch {
    Write-Host "мост пока не отвечает — смотрите workspace\watchdog.log и workspace\bridge.log"
}

Write-Host ''
Write-Host 'Готово. Откройте Excel (или workbook workspace\hermes-auto.xlsx) — панель Hermes подключится сама.'
Write-Host 'Снять автозапуск: powershell -ExecutionPolicy Bypass -File scripts\install-watchdog.ps1 -Remove'
