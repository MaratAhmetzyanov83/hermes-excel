# Регистрация надстройки Hermes в Excel под Windows.
# Запуск:  powershell -ExecutionPolicy Bypass -File scripts\sideload.ps1
#
# Рабочий механизм на этой машине — (1) WEF\Developer: значение <Id надстройки> = путь к манифесту.
# Общий каталог (2) Office игнорировал: список доверенных каталогов он кэширует в
# Wef\AppCommands\11.0\TrustedCatalog и мои ключи туда не попадали (проверено и на локальном
# пути, и на UNC \\localhost\C$\..., и после перезапуска Excel). Поэтому (1) — основной путь.

$ErrorActionPreference = 'Stop'
$root     = Split-Path -Parent $PSScriptRoot          # ...\hermes-excel
$catalog  = Join-Path $root 'catalog'
$manifest = Join-Path $root 'addin\manifest.xml'
$id       = ([xml](Get-Content $manifest -Raw)).OfficeApp.Id
Write-Host "Id надстройки: $id"

# (1) сайлоад: Excel ищет манифест по этому значению
$dev = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
New-Item -Path $dev -Force | Out-Null
New-ItemProperty -Path $dev -Name $id -Value $manifest -PropertyType String -Force | Out-Null
Write-Host "WEF\Developer: $id -> $manifest"

# (2) каталог доверенных надстроек (нужен UNC-путь; локальный путь Office не принимает)
New-Item -ItemType Directory -Force -Path $catalog | Out-Null
Copy-Item $manifest (Join-Path $catalog 'manifest.xml') -Force
$guid = '{9F5D2C41-7A83-4E1B-9C0D-1E2F3A4B5C6D}'
$key  = "HKCU:\Software\Microsoft\Office\16.0\Wef\TrustedCatalogs\$guid"
$unc  = '\\localhost\C$' + ($catalog -replace '^[A-Za-z]:', '')
New-Item -Path $key -Force | Out-Null
New-ItemProperty -Path $key -Name 'Id'    -Value $guid  -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Url'   -Value $unc   -PropertyType String -Force | Out-Null
New-ItemProperty -Path $key -Name 'Flags' -Value 1      -PropertyType DWord  -Force | Out-Null
Write-Host "TrustedCatalogs: $key -> $unc"

# Книга, которая сама открывает панель: Excel не активирует надстройку, пока она
# ни разу не активирована в процессе, поэтому панель поднимаем из документа.
$py = (Get-Command python).Source
& $py (Join-Path $root 'scripts\auto-open-workbook.py') create (Join-Path $root 'workspace\hermes-auto.xlsx')

Write-Host "`nГотово." -ForegroundColor Green
Write-Host "1) запустите мост: start-bridge.cmd"
Write-Host "2) откройте workspace\hermes-auto.xlsx — панель Hermes поднимется сама, без кликов"
Write-Host "3) кнопка на ленте (вкладка «Главная» -> группа Hermes Agent) появляется после первой активации"
