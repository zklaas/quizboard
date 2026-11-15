#!/usr/bin/env bash
# Quizboard launcher.
# Usage: ./run.sh [board.xlsx]   -> the board (reopens the last workbook when none is given)
#        ./run.sh --check        -> regression checks
set -euo pipefail
HERE="$(dirname "$(readlink -f "$0")")"   # no cd: a relative workbook path stays relative to the caller

PY="${QUIZBOARD_PYTHON:-python3}"

case "${1:-}" in
  --check) shift; exec "$PY" "$HERE/_quizboard_check.py" "$@" ;;
  *)       exec "$PY" "$HERE/quizboard.py" "$@" ;;
esac
