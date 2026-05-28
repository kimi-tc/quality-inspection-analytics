#!/bin/zsh
set -euo pipefail

SERVICE_LABEL="${SERVICE_LABEL:-com.kimi.weekly-inspection-analytics}"
PORT="${PORT:-3000}"

echo "[1/3] launchd service"
launchctl list | grep "$SERVICE_LABEL" || echo "Service not found in launchctl"

echo
echo "[2/3] Local HTTP check"
curl -I "http://127.0.0.1:${PORT}" || true

echo
echo "[3/3] Recent logs"
tail -n 20 "$HOME/Library/Logs/weekly-inspection-analytics.log" 2>/dev/null || echo "No stdout log yet"
