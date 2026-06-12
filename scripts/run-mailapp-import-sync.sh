#!/bin/zsh
set -euo pipefail

PROJECT_DIR="${PROJECT_DIR:-$HOME/quality-inspection-analytics}"

export PATH="/Users/a144522/.nvm/versions/node/v25.9.0/bin:/opt/homebrew/bin:/usr/local/bin:/opt/anaconda3/bin:/Users/tangchao144522/miniconda3/bin:$PATH"

NPM_BIN="$(command -v npm || true)"

if [[ -z "$NPM_BIN" ]]; then
  echo "npm not found in PATH: $PATH" >&2
  exit 1
fi

cd "$PROJECT_DIR"

echo "[$(date '+%Y-%m-%d %H:%M:%S')] Starting mail report import sync..."
"$NPM_BIN" run mailapp:import:sync
echo "[$(date '+%Y-%m-%d %H:%M:%S')] Mail report import sync finished."
