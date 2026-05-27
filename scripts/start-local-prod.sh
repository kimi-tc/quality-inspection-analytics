#!/bin/zsh
set -euo pipefail

PROJECT_DIR="${PROJECT_DIR:-$HOME/quality-inspection-analytics}"
PORT="${PORT:-3000}"
DATA_DIR="${DATA_DIR:-$PROJECT_DIR/data}"

cd "$PROJECT_DIR"
mkdir -p "$DATA_DIR"

npm install
npm run build

exec env PORT="$PORT" DATA_DIR="$DATA_DIR" npm run start:server
