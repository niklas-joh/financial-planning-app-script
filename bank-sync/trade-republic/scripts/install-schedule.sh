#!/bin/bash
# Run from Terminal in the repo (launchd cannot read ~/Documents). Needed once, and again after a change
# to bank-sync/trade-republic is merged (there is no auto-update). Safe to re-run.
#  1. Creates ~/trade-republic (outside Documents and iCloud) with .env (from .env.example, if missing),
#     credentials/, data/, session/ and logs/; never overwrites what is there.
#  2. Builds the tr-export:local image from the repo.
#  3. Installs run-scheduled.sh, login.sh, the runtime compose file and the launchd job.
set -euo pipefail

SRC="$(cd "$(dirname "$0")/.." && pwd)"
DEST="${TR_RUNTIME:-$HOME/trade-republic}"
LABEL=lu.niklas.trade-republic
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"

mkdir -p "$DEST"/{credentials,data,session,logs}
chmod 700 "$DEST" "$DEST/credentials" "$DEST/session"
[ -f "$DEST/.env" ] || cp "$SRC/.env.example" "$DEST/.env"
chmod 600 "$DEST/.env"
# The categoriser's service account can be reused: copy its key if there is none here yet.
if [ ! -f "$DEST/credentials/service-account.json" ] && [ -f "$HOME/categoriser/credentials/service-account.json" ]; then
  cp "$HOME/categoriser/credentials/service-account.json" "$DEST/credentials/service-account.json"
fi
find "$DEST/credentials" -type f -exec chmod 600 {} +

COMMIT=local
if git -C "$SRC" rev-parse --git-dir >/dev/null 2>&1 && [ -z "$(git -C "$SRC" status --porcelain -- .)" ]; then
  COMMIT=$(git -C "$SRC" rev-parse HEAD)
fi
echo "Building tr-export:local ($COMMIT) ..."
docker build --label "lu.niklas.trade-republic.commit=$COMMIT" -t tr-export:local "$SRC"

for f in run-scheduled.sh login.sh; do
  cp "$SRC/scripts/$f" "$DEST/$f"
  chmod +x "$DEST/$f"
done
cp "$SRC/deploy/docker-compose.runtime.yml" "$DEST/docker-compose.yml"

mkdir -p "$(dirname "$PLIST")"
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
sed "s#__TR_DIR__#$DEST#g" "$SRC/launchd/$LABEL.plist.template" > "$PLIST"
plutil -lint "$PLIST" >/dev/null
launchctl bootstrap "gui/$(id -u)" "$PLIST"

echo "Installed. Runtime folder: $DEST"
echo "Next: set TR_PHONE in $DEST/.env, then run $DEST/login.sh"
echo "Logs: cat $DEST/logs/launchd.out.log $DEST/logs/launchd.err.log"
