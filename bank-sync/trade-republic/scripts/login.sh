#!/bin/bash
# Installed as ~/trade-republic/login.sh. Run it in Terminal when the scheduled run says the
# Trade Republic session expired (or once after installing). It asks for your PIN (not stored),
# then you confirm the login in the Trade Republic app. Your phone stays logged in.
set -euo pipefail
export PATH="/usr/local/bin:/opt/homebrew/bin:/Applications/Docker.app/Contents/Resources/bin:$PATH"
cd "$(dirname "$0")"
docker compose run --rm tr-export python -m tr_export login
rm -f logs/.login-needed
