#!/bin/bash
# Installed as ~/categoriser/run-scheduled.sh and called by launchd.
# Runs the categoriser once from the categoriser:local image. Skips if it already
# succeeded today, so RunAtLoad + the 06:15 slot don't double-run.
#
# Auto-update (on unless CATEGORISER_AUTO_UPDATE=0 in .env): before running, it fetches
# `main` of the public repo into ~/categoriser/repo (only bank-sync/categoriser) and, if
# that folder changed since the image was built, rebuilds the image and refreshes the
# runtime files. Any failure (offline, build error) keeps the current image and the run
# goes ahead. A changed launchd plist template is only reported: re-run install-schedule.sh.
set -euo pipefail

export PATH="/usr/local/bin:/opt/homebrew/bin:/Applications/Docker.app/Contents/Resources/bin:$PATH"
cd "$(dirname "$0")"
mkdir -p logs
STAMP=logs/.last-success
TODAY=$(date +%F)

REPO_URL=${CATEGORISER_REPO_URL:-https://github.com/niklas-joh/financial-planning-app-script.git}
BRANCH=main
SUB=bank-sync/categoriser
IMAGE=categoriser:local
TREE_LABEL=lu.niklas.categoriser.tree
COMMIT_LABEL=lu.niklas.categoriser.commit
LABEL=lu.niklas.categoriser

log() { echo "$(date '+%F %T') $*"; }

# Value of KEY in .env ('' if unset). Not `source`: .env is data for the container.
env_value() { sed -n "s/^$1=//p" .env 2>/dev/null | tail -n 1 | tr -d '"'"'"' \r'; }

if [ -f "$STAMP" ] && [ "$(cat "$STAMP")" = "$TODAY" ]; then
  log "already ran today, skipping"
  exit 0
fi

if ! docker info >/dev/null 2>&1; then
  log "starting Docker Desktop"
  open -ga Docker
  for _ in $(seq 1 60); do docker info >/dev/null 2>&1 && break; sleep 5; done
  docker info >/dev/null 2>&1 || { echo "Docker did not start"; exit 1; }
fi

# Replace a runtime file only when it changed; mv gives it a new inode, so this
# script can replace itself while bash is still reading the old copy.
refresh_file() {
  local src=$1 dest=$2
  cmp -s "$src" "$dest" && return 1
  cp -p "$src" "$dest.new" && mv -f "$dest.new" "$dest"
}

auto_update() {
  if [ "$(env_value CATEGORISER_AUTO_UPDATE)" = "0" ]; then
    log "auto-update off (CATEGORISER_AUTO_UPDATE=0)"
    return 0
  fi
  if [ ! -d repo/.git ]; then
    rm -rf repo
    git clone --quiet --depth 1 --filter=blob:none --sparse --branch "$BRANCH" "$REPO_URL" repo || return 1
    git -C repo sparse-checkout set "$SUB" || return 1
  fi
  git -C repo fetch --quiet --depth 1 origin "$BRANCH" || return 1
  local commit want have
  commit=$(git -C repo rev-parse FETCH_HEAD) || return 1
  want=$(git -C repo rev-parse "FETCH_HEAD:$SUB") || return 1
  have=$(docker image inspect -f "{{ index .Config.Labels \"$TREE_LABEL\" }}" "$IMAGE" 2>/dev/null || true)
  git -C repo checkout --quiet --force --detach FETCH_HEAD || return 1

  if [ "$want" != "$have" ]; then
    log "code changed on $BRANCH (${commit:0:7}): rebuilding $IMAGE"
    docker build --quiet --label "$TREE_LABEL=$want" --label "$COMMIT_LABEL=$commit" -t "$IMAGE" "repo/$SUB" >/dev/null \
      || { log "build failed: keeping the current image"; return 1; }
    log "image updated to ${commit:0:7}"
  fi

  refresh_file "repo/$SUB/deploy/docker-compose.runtime.yml" docker-compose.yml && log "docker-compose.yml updated" || true
  refresh_file "repo/$SUB/deploy/RUNTIME_README.md" README.md || true
  refresh_file "repo/$SUB/scripts/run-scheduled.sh" run-scheduled.sh \
    && log "run-scheduled.sh updated (used from the next run)" || true

  local plist="$HOME/Library/LaunchAgents/$LABEL.plist"
  if ! sed "s#__CATEGORISER_DIR__#$PWD#g" "repo/$SUB/launchd/$LABEL.plist.template" | cmp -s - "$plist"; then
    log "WARNING: the launchd schedule changed on $BRANCH; run $PWD/repo/$SUB/scripts/install-schedule.sh to apply it"
  fi
}

auto_update || log "auto-update did not complete: running the current image"

if ! docker image inspect "$IMAGE" >/dev/null 2>&1; then
  echo "Image $IMAGE missing: run scripts/install-schedule.sh from the repo"
  exit 1
fi

log "run start ($(docker image inspect -f "{{ index .Config.Labels \"$COMMIT_LABEL\" }}" "$IMAGE" 2>/dev/null | cut -c1-7))"
docker compose run --rm categoriser
echo "$TODAY" > "$STAMP"
log "run ok"
