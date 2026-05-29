#!/bin/zsh
set -euo pipefail

ENV_FILE="${ENV_FILE:-$HOME/.weekly-inspection-tunnel.env}"

if [[ -f "$ENV_FILE" ]]; then
  # shellcheck disable=SC1090
  source "$ENV_FILE"
fi

VPS_USER="${VPS_USER:-}"
VPS_HOST="${VPS_HOST:-}"
VPS_PORT="${VPS_PORT:-22}"
REMOTE_PORT="${REMOTE_PORT:-13000}"
PUBLIC_URL="${PUBLIC_URL:-}"

echo "[1/4] local tunnel process"
ps -ef | grep "ssh .*127.0.0.1:${REMOTE_PORT}" | grep -v grep || echo "Tunnel process not found"

echo
echo "[2/4] local dashboard"
curl -I "http://127.0.0.1:${LOCAL_PORT:-3000}" || true

if [[ -n "$VPS_USER" && -n "$VPS_HOST" ]]; then
  echo
  echo "[3/4] VPS loopback tunnel endpoint"
  ssh -p "$VPS_PORT" "${VPS_USER}@${VPS_HOST}" "curl -I http://127.0.0.1:${REMOTE_PORT}" || true
else
  echo
  echo "[3/4] VPS loopback tunnel endpoint"
  echo "Skipped: VPS_USER or VPS_HOST is not configured in $ENV_FILE"
fi

echo
echo "[4/4] public URL"
if [[ -n "$PUBLIC_URL" ]]; then
  curl -I "$PUBLIC_URL" || true
else
  echo "Skipped: PUBLIC_URL is not configured in $ENV_FILE"
fi
