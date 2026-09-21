# Ставит сторожа моста «Hermes для Excel»: только задачу планировщика (раз в минуту).
# Мост поднимается только пока запущен Excel и останавливается после его закрытия.
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
    Write-Host "автозапуск при входе: отключен" -NoNewline; if (Test-Path $lnk) { Remove-Item $lnk -Force -ErrorAction SilentlyContinue; Write-Host ' (старый ярлык удалён)' } else { Write-Host '' }
    Write-Host "журнал:  " -NoNewline; Write-Host (Join-Path $root 'workspace\watchdog.log')
    exit 0
}

if ($Remove) {
    schtasks /Delete /TN $task /F 2>$null | Out-Null
    Remove-Item $lnk -ErrorAction SilentlyContinue
    Write-Host 'Сторож и автозапуск сняты. Мост можно поднять вручную: start-bridge.cmd'
    exit 0
}

# Запуск без окна: powershell.exe — консольное приложение, и при запуске раз в минуту из планировщика
# или из «Автозагрузки» оно может мигнуть чёрным окном. wscript.exe живёт в GUI-подсистеме, консоли
# не создаёт вовсе, а run-watchdog.vbs просит скрытое окно и для дочернего процесса.
$vbs = Join-Path $PSScriptRoot 'run-watchdog.vbs'
$useVbs = Test-Path $vbs

# 1. Удаляем старый ярлык: он запускал мост независимо от Excel и блокировал обновления Hermes.
Remove-Item $lnk -Force -ErrorAction SilentlyContinue

# 2. Задача планировщика: раз в минуту проверяет мост и поднимает его только при работающем Excel
if ($useVbs) {
    $action = "wscript.exe `"$vbs`""
} else {
    $action = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$watch`""
}
schtasks /Create /TN $task /TR $action /SC MINUTE /MO 1 /F | Out-Null
Write-Host "задача планировщика: $task (раз в минуту, без окна)"

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
Write-Host 'Готово. Мост работает только пока запущен Excel; после закрытия Excel он будет остановлен.'
Write-Host 'Снять сторож: powershell -ExecutionPolicy Bypass -File scripts\install-watchdog.ps1 -Remove'
