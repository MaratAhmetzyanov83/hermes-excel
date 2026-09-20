@echo off
rem Hermes Excel Bridge — запуск, остановка и состояние локального моста надстройки.
rem   start-bridge.cmd          — поднять мост, если он не запущен (и проверить, что отвечает)
rem   start-bridge.cmd stop     — остановить мост
rem   start-bridge.cmd status   — состояние моста, сторожа и автозапуска
rem   start-bridge.cmd excel    — поднять мост и открыть Excel с книгой, которая сама открывает панель
setlocal
cd /d "%~dp0"

if /I "%~1"=="stop"   goto stop
if /I "%~1"=="status" goto status
if /I "%~1"=="excel"  goto excel

echo [Hermes] проверяю и поднимаю мост на https://localhost:3443 ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\bridge-watchdog.ps1" -Force
powershell -NoProfile -Command "try { (Invoke-WebRequest -UseBasicParsing https://localhost:3443/health -TimeoutSec 5).Content } catch { 'мост не отвечает: ' + $_.Exception.Message }"
goto :eof

:status
powershell -NoProfile -Command "try { 'мост: ' + (Invoke-WebRequest -UseBasicParsing https://localhost:3443/health -TimeoutSec 5).Content } catch { 'мост: не отвечает' }"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\install-watchdog.ps1" -Status
goto :eof

:excel
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\bridge-watchdog.ps1" -Force
if exist "%~dp0workspace\hermes-auto.xlsx" (
  echo [Hermes] открываю Excel с книгой hermes-auto.xlsx ...
  start "" "%~dp0workspace\hermes-auto.xlsx"
) else (
  echo [Hermes] книги авто-открытия нет — собираю ...
  python "%~dp0scripts\auto-open-workbook.py" create
  start "" "%~dp0workspace\hermes-auto.xlsx"
)
goto :eof

:stop
echo [Hermes] останавливаю мост...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3443 -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess | Sort-Object -Unique | ForEach-Object { Stop-Process -Id $_ -Force; 'остановлен pid ' + $_ }"
echo [Hermes] сторож поднимет мост снова при следующей проверке. Снять сторожа: scripts\install-watchdog.ps1 -Remove
goto :eof
