@echo off
rem Установка надстройки «Hermes для Excel» одной командой (идемпотентно).
rem   install.cmd                 — настроить всё: сертификаты, бот, сайлоад, сторож, мост
rem   install.cmd -DryRun         — показать план, ничего не меняя
rem   install.cmd -Profile excel2 — другое имя бота
rem   install.cmd -NoBridge       — без сторожа и автозапуска
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\install.ps1" %*
