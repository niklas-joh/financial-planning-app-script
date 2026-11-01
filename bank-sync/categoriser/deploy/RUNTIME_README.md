# ~/categoriser — runtime folder (do not commit, do not edit code here)

This folder is what the daily scheduled job runs from. The **code** lives in the repo:
`~/Documents/Coding/appscript/financial-planning-app-script/bank-sync/categoriser/`
(full docs: its README.md). Everything here was created by its `scripts/install-schedule.sh`.

| File | What it is | Edit here? |
|---|---|---|
| `.env` | sheet ID and sheet names | yes, next run picks it up |
| `credentials/service-account.json` | Google key (secret) | replace to rotate the key |
| `data/`, `models/` | categoriser data and trained models | written by the job |
| `logs/` | `launchd.out.log`, `launchd.err.log`, `.last-success` | read only |
| `docker-compose.yml`, `run-scheduled.sh`, `README.md` | copies from the repo | **no**, refreshed from `main` |
| `repo/` | sparse clone of `main` used to rebuild the image | **no** |

The job runs the Docker image `categoriser:local`. Before each daily run it fetches `main` and, if
`bank-sync/categoriser/` changed, rebuilds the image. So **merging to `main` is the deploy**; it goes live
at the next 06:15 run once GitHub shows that commit's `Tests (categoriser)` check passed. If the check is
pending or red, or the fetch or build fails, the current image runs as before.

To run a branch instead, add `CATEGORISER_AUTO_UPDATE=0` to `.env` and run `install-schedule.sh` from that
branch. Remove the line to go back to `main`. A changed schedule (plist) still needs `install-schedule.sh`;
the log warns when that is the case.

Problems the job skipped (e.g. duplicate rows) are listed in the `categoriser_issues`
sheet of Bank Sync, not only in the logs.

Status: `launchctl print gui/$(id -u)/lu.niklas.categoriser | grep -E 'state|last exit code'`
Run now: `rm -f logs/.last-success && launchctl kickstart gui/$(id -u)/lu.niklas.categoriser`
Stop: `launchctl bootout gui/$(id -u)/lu.niklas.categoriser`
