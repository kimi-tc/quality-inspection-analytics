#!/bin/zsh
set -euo pipefail

PROJECT_DIR="${PROJECT_DIR:-$HOME/quality-inspection-analytics}"
DATA_FILE="${DATA_FILE:-$PROJECT_DIR/data/shared-dataset.json}"
BACKUP_DIR="${BACKUP_DIR:-$PROJECT_DIR/backups}"
TIMESTAMP="$(date +%Y%m%d-%H%M%S)"

mkdir -p "$BACKUP_DIR"

if [[ ! -f "$DATA_FILE" ]]; then
  echo "Data file not found: $DATA_FILE" >&2
  exit 1
fi

cp "$DATA_FILE" "$BACKUP_DIR/shared-dataset-$TIMESTAMP.json"
echo "Backup created: $BACKUP_DIR/shared-dataset-$TIMESTAMP.json"
