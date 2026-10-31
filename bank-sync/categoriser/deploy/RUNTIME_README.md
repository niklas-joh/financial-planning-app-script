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
| `docker-compose.yml`, `run-scheduled.sh` | copies from the repo | **no**, overwritten on reinstall |

The job runs the Docker image `categoriser:local`, a snapshot of the categoriser code
taken when `install-schedule.sh` was last run. **Changing the code in the repo does nothing
until you re-run:**

```bash
~/Documents/Coding/appscript/financial-planning-app-script/bank-sync/categoriser/scripts/install-schedule.sh
```

Problems the job skipped (e.g. duplicate rows) are listed in the `categoriser_issues`
sheet of Bank Sync, not only in the logs.

Status: `launchctl print gui/$(id -u)/lu.niklas.categoriser | grep -E 'state|last exit code'`
Run now: `rm -f logs/.last-success && launchctl kickstart gui/$(id -u)/lu.niklas.categoriser`
Stop: `launchctl bootout gui/$(id -u)/lu.niklas.categoriser`
