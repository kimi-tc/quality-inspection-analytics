#!/bin/zsh
set -euo pipefail

PROJECT_DIR="${PROJECT_DIR:-$HOME/quality-inspection-analytics}"
SERVICE_LABEL="${SERVICE_LABEL:-com.kimi.weekly-inspection-analytics}"

cd "$PROJECT_DIR"

echo "[1/4] Pulling latest code..."
git pull

echo "[2/4] Installing dependencies..."
npm install

echo "[3/4] Building frontend..."
npm run build

echo "[4/4] Restarting launchd service..."
launchctl kickstart -k "gui/$(id -u)/$SERVICE_LABEL"

echo "Done. Service restarted."
