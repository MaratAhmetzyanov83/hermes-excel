#!/usr/bin/env sh
# Установка из git-bash / MSYS-оболочки.
#
# Зачем отдельный файл: в git-bash `cmd //c install.cmd` калечится (MSYS съедает `//c`) и вместо
# запуска открывает интерактивную консоль cmd — агент на этом виснет. Здесь путь ровный:
# PowerShell получает нативный путь и делает всё то же, что install.cmd.
#
#   ./install.sh                 — поставить всё
#   ./install.sh -DryRun         — показать план
#   ./install.sh -Profile excel2 — другое имя бота
set -e
DIR=$(cd "$(dirname "$0")" && pwd)
SCRIPT="$DIR/scripts/install.ps1"
WIN=$(cygpath -w "$SCRIPT" 2>/dev/null || echo "$SCRIPT")
exec powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$WIN" "$@"
