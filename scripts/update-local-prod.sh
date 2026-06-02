#!/bin/zsh
set -euo pipefail

PROJECT_DIR="${PROJECT_DIR:-$HOME/quality-inspection-analytics}"
SERVICE_LABEL="${SERVICE_LABEL:-com.kimi.weekly-inspection-analytics}"

export PATH="/Users/a144522/.nvm/versions/node/v25.9.0/bin:/opt/homebrew/bin:/usr/local/bin:/opt/anaconda3/bin:/Users/tangchao144522/miniconda3/bin:$PATH"

NPM_BIN="$(command -v npm || true)"

if [[ -z "$NPM_BIN" ]]; then
  echo "npm not found in PATH: $PATH" >&2
  exit 1
fi

cd "$PROJECT_DIR"

echo "[1/4] Pulling latest code..."
git pull

echo "[2/4] Installing dependencies..."
"$NPM_BIN" install

echo "[3/4] Building frontend..."
"$NPM_BIN" run build

echo "[4/4] Restarting launchd service..."
launchctl kickstart -k "gui/$(id -u)/$SERVICE_LABEL"

echo "Done. Service restarted."
