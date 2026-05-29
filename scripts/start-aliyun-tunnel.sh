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
LOCAL_HOST="${LOCAL_HOST:-127.0.0.1}"
LOCAL_PORT="${LOCAL_PORT:-3000}"

if [[ -z "$VPS_USER" || -z "$VPS_HOST" ]]; then
  echo "Missing VPS_USER or VPS_HOST. Configure $ENV_FILE first." >&2
  exit 1
fi

echo "Starting reverse tunnel: ${VPS_HOST}:127.0.0.1:${REMOTE_PORT} -> ${LOCAL_HOST}:${LOCAL_PORT}"

while true; do
  if ! curl -fsS "http://${LOCAL_HOST}:${LOCAL_PORT}" >/dev/null; then
    echo "Local dashboard is not ready at http://${LOCAL_HOST}:${LOCAL_PORT}; retrying in 10s..."
    sleep 10
    continue
  fi

  ssh \
    -N \
    -T \
    -p "$VPS_PORT" \
    -o ExitOnForwardFailure=yes \
    -o ServerAliveInterval=30 \
    -o ServerAliveCountMax=3 \
    -o StrictHostKeyChecking=accept-new \
    -R "127.0.0.1:${REMOTE_PORT}:${LOCAL_HOST}:${LOCAL_PORT}" \
    "${VPS_USER}@${VPS_HOST}" || true

  echo "Tunnel disconnected; reconnecting in 5s..."
  sleep 5
done
