#!/usr/bin/env sh
# Проверка надстройки из git-bash / MSYS-оболочки (в cmd — doctor.cmd).
#
#   ./doctor.sh          — таблица проверок
#   ./doctor.sh -Json    — машиночитаемо
#   ./doctor.sh -Fix     — сначала поднять мост, если лежит
set -e
DIR=$(cd "$(dirname "$0")" && pwd)
SCRIPT="$DIR/scripts/doctor.ps1"
WIN=$(cygpath -w "$SCRIPT" 2>/dev/null || echo "$SCRIPT")
exec powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$WIN" "$@"
