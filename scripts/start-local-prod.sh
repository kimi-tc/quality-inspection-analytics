#!/bin/zsh
set -euo pipefail

PROJECT_DIR="${PROJECT_DIR:-$HOME/quality-inspection-analytics}"
PORT="${PORT:-3000}"
DATA_DIR="${DATA_DIR:-$PROJECT_DIR/data}"

export PATH="/opt/homebrew/bin:/usr/local/bin:/opt/anaconda3/bin:/Users/tangchao144522/miniconda3/bin:$PATH"

NPM_BIN="$(command -v npm || true)"

if [[ -z "$NPM_BIN" ]]; then
  echo "npm not found in PATH: $PATH" >&2
  exit 1
fi

cd "$PROJECT_DIR"
mkdir -p "$DATA_DIR"

"$NPM_BIN" install
"$NPM_BIN" run build

exec env PORT="$PORT" DATA_DIR="$DATA_DIR" "$NPM_BIN" run start:server
