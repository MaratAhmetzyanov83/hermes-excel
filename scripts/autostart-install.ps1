# Автозапуск моста Hermes при входе в Windows (ярлык в папке «Автозагрузка»).
# Включает локальный сервер с полным доступом к инструментам агента — включайте осознанно.
# Запуск:   powershell -ExecutionPolicy Bypass -File scripts\autostart-install.ps1
# Отключить: удалите ярлык "Hermes Excel Bridge.lnk" из папки Автозагрузка
#            (shell:startup).

$ErrorActionPreference = 'Stop'
$root  = Split-Path -Parent $PSScriptRoot
$py    = (Get-Command python).Source
$startup = [Environment]::GetFolderPath('Startup')
$lnk   = Join-Path $startup 'Hermes Excel Bridge.lnk'

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($lnk)
$shortcut.TargetPath = $py
$shortcut.Arguments  = '"' + (Join-Path $root 'bridge\bridge.py') + '"'
$shortcut.WorkingDirectory = $root
$shortcut.WindowStyle = 7                  # свёрнутое окно
$shortcut.Description = 'Hermes Excel Bridge (https://localhost:3443)'
$shortcut.Save()

Write-Host "Автозапуск создан: $lnk"
Write-Host "Отключить: удалите этот ярлык."
