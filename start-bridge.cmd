@echo off
rem Hermes Excel Bridge — запуск/остановка локального моста
setlocal
cd /d "%~dp0"

if /I "%~1"=="stop" goto stop

echo [Hermes] стартую мост на https://localhost:3443 ...
start "Hermes Excel Bridge" /min cmd /c "cd /d %~dp0 && python bridge\bridge.py"
timeout /t 3 /nobreak >nul
powershell -NoProfile -Command "try { (Invoke-WebRequest -UseBasicParsing https://localhost:3443/health -TimeoutSec 5).Content } catch { 'мост не отвечает: ' + $_.Exception.Message }"
echo.
echo Открыть Excel:  start excel
goto :eof

:stop
echo [Hermes] останавливаю мост...
powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 3443 -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess | Sort-Object -Unique | ForEach-Object { Stop-Process -Id $_ -Force; 'остановлен pid ' + $_ }"
