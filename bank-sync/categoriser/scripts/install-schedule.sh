#!/bin/bash
# Run from Terminal/SSH (which can read ~/Documents). Safe to re-run, e.g.
# after code changes to rebuild the image.
#  1. Creates ~/categoriser (outside Documents and iCloud) and MOVES .env,
#     credentials/, data/, models/ there from the repo folder if present.
#  2. Builds the categoriser:local image from the repo.
#  3. Installs run-scheduled.sh, the runtime compose file and the launchd job.
set -euo pipefail

SRC="$(cd "$(dirname "$0")/.." && pwd)"
DEST="${CATEGORISER_HOME:-$HOME/categoriser}"
LABEL=lu.niklas.categoriser
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"

mkdir -p "$DEST"/{credentials,data,models,logs}
chmod 700 "$DEST" "$DEST/credentials"

# Move secrets and data out of the repo (never overwrite what DEST has).
if [ -f "$SRC/.env" ] && [ ! -f "$DEST/.env" ]; then mv "$SRC/.env" "$DEST/.env"; fi
[ -f "$DEST/.env" ] || cp "$SRC/.env.example" "$DEST/.env"
for d in credentials data models; do
  if [ -d "$SRC/$d" ]; then
    find "$SRC/$d" -mindepth 1 -maxdepth 1 | while read -r f; do
      [ -e "$DEST/$d/$(basename "$f")" ] || mv "$f" "$DEST/$d/"
    done
  fi
done
chmod 600 "$DEST/.env"
find "$DEST/credentials" -type f -exec chmod 600 {} +

echo "Building categoriser:local ..."
docker compose -f "$SRC/docker-compose.yml" build

cp "$SRC/scripts/run-scheduled.sh" "$DEST/run-scheduled.sh"
chmod +x "$DEST/run-scheduled.sh"
cp "$SRC/deploy/docker-compose.runtime.yml" "$DEST/docker-compose.yml"
cp "$SRC/deploy/RUNTIME_README.md" "$DEST/README.md"

mkdir -p "$(dirname "$PLIST")"
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
sed "s#__CATEGORISER_DIR__#$DEST#g" "$SRC/launchd/$LABEL.plist.template" > "$PLIST"
plutil -lint "$PLIST" >/dev/null
launchctl bootstrap "gui/$(id -u)" "$PLIST"

echo "Installed. Runtime folder: $DEST"
echo "Check: launchctl print gui/\$(id -u)/$LABEL | grep -E 'state|last exit code'"
echo "Logs:  cat $DEST/logs/launchd.out.log $DEST/logs/launchd.err.log"
