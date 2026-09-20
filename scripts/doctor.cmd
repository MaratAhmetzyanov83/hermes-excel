@echo off
rem Быстрая проверка всей надстройки «Hermes для Excel»: мост, бот, сайлоад, сторож, панель.
rem   doctor.cmd          — таблица проверок
rem   doctor.cmd fix      — то же, но сначала поднять мост, если он лежит
setlocal
cd /d "%~dp0.."
if /I "%~1"=="fix" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0doctor.ps1" -Fix
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0doctor.ps1"
)
