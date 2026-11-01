#!/bin/bash
# Installed as ~/trade-republic/run-scheduled.sh and called by launchd (daily 05:40, before the Bank Sync
# daily sync between 06:00 and 07:00). Skips if it already succeeded today. When the Trade Republic
# session has expired (exit 3) it shows a macOS notification once a day: run ~/trade-republic/login.sh.
# No auto-update: after a change to bank-sync/trade-republic is merged, re-run scripts/install-schedule.sh.
set -uo pipefail

export PATH="/usr/local/bin:/opt/homebrew/bin:/Applications/Docker.app/Contents/Resources/bin:$PATH"
cd "$(dirname "$0")"
mkdir -p logs
STAMP=logs/.last-success
TODAY=$(date +%F)

log() { echo "$(date '+%F %T') $*"; }

notify() {
  /usr/bin/osascript -e "display notification \"$1\" with title \"Bank Sync: Trade Republic\"" >/dev/null 2>&1 || true
}

if [ -f "$STAMP" ] && [ "$(cat "$STAMP")" = "$TODAY" ]; then
  log "already ran today, skipping"
  exit 0
fi

if ! docker info >/dev/null 2>&1; then
  log "starting Docker Desktop"
  open -ga Docker
  for _ in $(seq 1 60); do docker info >/dev/null 2>&1 && break; sleep 5; done
  docker info >/dev/null 2>&1 || { log "Docker did not start"; exit 1; }
fi

log "run start"
docker compose run --rm -T tr-export
code=$?
case $code in
  0)
    echo "$TODAY" > "$STAMP"
    rm -f logs/.login-needed
    log "run ok"
    ;;
  3)
    log "login needed: run $PWD/login.sh"
    if [ "$(cat logs/.login-needed 2>/dev/null)" != "$TODAY" ]; then
      notify "Session expired: run ~/trade-republic/login.sh in Terminal"
      echo "$TODAY" > logs/.login-needed
    fi
    ;;
  *)
    log "run failed (exit $code)"
    ;;
esac
exit $code
