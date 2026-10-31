# categoriser

Python job (inside `bank-sync/`) that reads the **Bank Sync** Google Sheet and writes categories to a
separate `categorised` sheet. Categorisation logic (rules, model) is not
implemented yet. This is the I/O scaffold.

> **This repo is public.** Keys, `.env`, CSVs, models and logs are git-ignored
> and blocked by the pre-commit hook (`../../scripts/hooks/pre-commit`). Never use
> `--no-verify` to get round it.

## Output contract

- Reads `transactions`. **Never writes to it** (enforced in `sheets._assert_writable`).
- Writes only to `categorised`, keyed by `transactionId`:

  | transactionId | owner | type | category | subCategory | method | confidence | ruleId | reviewed | updatedAt |
  |---|---|---|---|---|---|---|---|---|---|

  `method` is `rule`, `model` or `manual`; `reviewed` is a checkbox; `updatedAt` is ISO 8601 UTC;
  `confidence` is blank when not applicable.
- Rows with `reviewed = TRUE` are never overwritten. Other existing rows are updated in place; new IDs are appended.
  Just before writing, the key and `reviewed` columns are re-read: a row ticked or moved while the run was
  planning is skipped (`skipped_reviewed` / `skipped_moved` in the log).
- A `transactionId` on more than one row is ambiguous: every copy is left unchanged, the rest of the run
  carries on, and the problem goes to the issues log. Delete the extra rows to resolve it.

### Issues log (`categoriser_issues` sheet)

Created automatically. One row per issue, never repeated:

| issueKey | status | type | transactionId | detail | source | firstSeen | lastSeen | resolvedAt |
|---|---|---|---|---|---|---|---|---|

- New issue → appended as `open`. Still there next run → only `lastSeen` and `detail` are refreshed.
- Gone → `status` becomes `resolved` with `resolvedAt`. Comes back → reopened on the same row.
- Issue types today: `duplicate_transaction_id`. The job writes only to `categorised` and this sheet.
- Missing sheet → created with headers, frozen header row, and checkbox validation on `reviewed`.
  Existing sheet with different headers → the run fails rather than touching it.

## Layout

```
bank-sync/categoriser/
  src/categoriser/  config.py  schema.py  merge.py (pure upsert plan)  sheets.py (gspread I/O)
                    smoke.py   run.py (entry point)
  tests/            pytest, no network
  Dockerfile  docker-compose.yml (dev)  deploy/docker-compose.runtime.yml
  scripts/install-schedule.sh  scripts/run-scheduled.sh  launchd/*.plist.template
```

Secrets and data live **outside the repo** in the runtime folder `~/categoriser/`
(override with `CATEGORISER_HOME`):

```
~/categoriser/
  .env  credentials/service-account.json  data/  models/  logs/
  docker-compose.yml  run-scheduled.sh       (installed by install-schedule.sh)
```

Why: macOS blocks launchd jobs from reading `~/Documents`, and this keeps keys
and transaction data out of iCloud and the public repo.

## 1. Google access (one-off)

1. Go to <https://console.cloud.google.com/>, create a project (e.g. `bank-sync-categoriser`).
2. **APIs & Services → Library**: enable **Google Sheets API**. (Drive API isn't needed: we open by ID.)
3. **IAM & Admin → Service Accounts → Create service account**, e.g. `categoriser`. Skip the optional role steps.
4. Open the account → **Keys → Add key → Create new key → JSON**. A file downloads.
5. Move it into the runtime folder:
   ```bash
   mkdir -p ~/categoriser/credentials && chmod 700 ~/categoriser ~/categoriser/credentials
   mv ~/Downloads/<project>-*.json ~/categoriser/credentials/service-account.json
   chmod 600 ~/categoriser/credentials/service-account.json
   ```bash
cd bank-sync/categoriser
docker compose run --rm --build categoriser python -m categoriser.smoke
```
Or a venv:
```bash
cd bank-sync/categoriser
python3.12 -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
python -m categoriser.smoke
```
Expected: the `transactions` header row, then `categorised header: [...]` (the sheet is created on first run).

## 3. Run once / tests

```bash
docker compose run --rm categoriser      # python -m categoriser.run
pytest                                   # in the venv
```

## 4. Scheduling (launchd, macOS)

Daily at 06:15 with `RunAtLoad`. launchd runs a missed calendar slot once after wake. If the Mac
was off, `RunAtLoad` runs it at login. `run-scheduled.sh` starts Docker Desktop if needed and skips if
today's run already succeeded. It uses the prebuilt `categoriser:local` image and never reads the repo.

Install or update (from Terminal or SSH, re-run after code changes to rebuild the image):
```bash
bank-sync/categoriser/scripts/install-schedule.sh
```
It creates `~/categoriser/`, moves `.env`/`credentials/`/`data/`/`models/` there from the repo folder,
builds the image, installs the runtime files and (re)loads `~/Library/LaunchAgents/lu.niklas.categoriser.plist`.
No Full Disk Access needed.

Check:
```bash
launchctl print gui/$(id -u)/lu.niklas.categoriser | grep -E 'state|last exit code'
cat ~/categoriser/logs/launchd.out.log ~/categoriser/logs/launchd.err.log
launchctl kickstart gui/$(id -u)/lu.niklas.categoriser   # run now (skips if it ran today)
```
Uninstall: `launchctl bootout gui/$(id -u)/lu.niklas.categoriser`.

## 5. After you change something: what to re-run

The scheduled job runs the Docker image `categoriser:local`, a **snapshot of the categoriser code**
taken the last time `scripts/install-schedule.sh` ran. Pull or edit all you like; the job keeps
running the old snapshot until you re-run the script. Run it from a checkout of the branch you
want live (normally `main`).

| You changed | Re-run `install-schedule.sh`? |
|---|---|
| `bank-sync/categoriser/src/**`, `pyproject.toml`, `Dockerfile` | **Yes** (rebuilds the image) |
| `scripts/run-scheduled.sh`, `deploy/*`, `launchd/*.plist.template` | **Yes** (re-copies them to `~/categoriser`) |
| `bank-sync/categoriser/tests/**`, this README | No |
| `~/categoriser/.env`, or the key in `~/categoriser/credentials/` | No (read on every run) |
| Apps Script code (`src/`, `bank-sync/src/`) or anything else in the repo | No (not part of the categoriser; deploy with clasp as usual) |
| Column names in the `transactions` sheet | Check the code still matches, then yes if you changed it |

Re-running is safe: it never overwrites `.env`, keys, data or models already in `~/categoriser`.
