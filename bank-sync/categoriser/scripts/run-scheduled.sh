#!/bin/bash
# Installed as ~/categoriser/run-scheduled.sh and called by launchd.
# Runs the categoriser once from the prebuilt image. Skips if it already
# succeeded today, so RunAtLoad + the 06:15 slot don't double-run.
set -euo pipefail

export PATH="/usr/local/bin:/opt/homebrew/bin:/Applications/Docker.app/Contents/Resources/bin:$PATH"
cd "$(dirname "$0")"
mkdir -p logs
STAMP=logs/.last-success
TODAY=$(date +%F)

if [ -f "$STAMP" ] && [ "$(cat "$STAMP")" = "$TODAY" ]; then
  echo "$(date '+%F %T') already ran today, skipping"
  exit 0
fi

if ! docker info >/dev/null 2>&1; then
  echo "$(date '+%F %T') starting Docker Desktop"
  open -ga Docker
  for _ in $(seq 1 60); do docker info >/dev/null 2>&1 && break; sleep 5; done
  docker info >/dev/null 2>&1 || { echo "Docker did not start"; exit 1; }
fi

if ! docker image inspect categoriser:local >/dev/null 2>&1; then
  echo "Image categoriser:local missing: run scripts/install-schedule.sh from the repo"
  exit 1
fi

echo "$(date '+%F %T') run start"
docker compose run --rm categoriser
echo "$TODAY" > "$STAMP"
echo "$(date '+%F %T') run ok"
