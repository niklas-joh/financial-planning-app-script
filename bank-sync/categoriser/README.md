# categoriser

Python job (inside `bank-sync/`) that reads the **Bank Sync** Google Sheet and writes categories to a
separate `categorised` sheet: rules you keep in the `categoriser_rules` sheet, internal transfer pairing,
then a model for the rest, which labels a row only when it is confident. Everything else is left blank.

> **This repo is public.** Keys, `.env`, CSVs, models and logs are git-ignored
> and blocked by the pre-commit hook (`../../scripts/hooks/pre-commit`). Never use
> `--no-verify` to get round it.

## Output contract

- Reads `transactions`. **Never writes to it** (enforced in `sheets._assert_writable`).
- Writes only to `categorised`, `categoriser_issues`, `categoriser_rules` (only `hits`/`lastMatched`, and
  rows appended by `categoriser.mine`) and `categoriser_taxonomy` (created, and seeded once). This list is
  `Settings.writable_sheets`. The legacy sheet is only opened with a read-only token.
- `categorised` is keyed by `transactionId`:

  | transactionId | owner | type | category | subCategory | method | confidence | ruleId | reviewed | updatedAt |
  |---|---|---|---|---|---|---|---|---|---|

  `method` is `rule`, `model` or `manual`; `reviewed` is a checkbox; `updatedAt` is ISO 8601 UTC;
  `confidence` is blank when not applicable.
- Rows with `reviewed = TRUE` are never overwritten. Other existing rows are updated in place when something
  changed (so `updatedAt` is the last change); new IDs are appended below the last row with an ID.
- A `rule` or `model` row that the run no longer produces (rule switched off or changed, model no longer
  confident) is cleared: only its `transactionId` and `updatedAt` stay (`cleared` in the log). `model` rows
  are kept as they are when the model could not be loaded.
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
- Issue types: `duplicate_transaction_id`; `invalid_rule` (an active rule that was skipped, keyed by
  ruleId); `rule_disagrees_with_review` (a reviewed row a rule would label differently, so the rule can be
  fixed).
- Missing job sheet → created with headers, frozen header row, and its checkboxes/dropdowns.
  Existing sheet with different headers → the run fails before writing anything.

## How rows are categorised

1. **Rules** (`categoriser_rules`, active ones). Tried by `priority` (lowest first), then by how many match
   columns are set (most first), then `ruleId`. The first match wins. Output: `method` = `rule`, `ruleId` set.
2. **Internal transfers**, for rows no rule covered. Both legs of a pair are labelled
   `Internal Transfer / Internal Transfer / <sub>`, `ruleId` = `transfer:pair`. A pair is an outgoing and an
   incoming row on two of our accounts, same currency and amount, at most 3 days apart; card rows are never
   paired. Ties: a leg whose `counterpartyAccount` is the other's IBAN first, then the same description, then the
   closest date. A single leg sent to one of our own IBANs is `transfer:own-account`. `<sub>` is
   `From <Person>` when exactly one person (not `JOINT_OWNER`) is involved and the taxonomy has it, else
   `Internal Transfer`.
3. **Model**, for rows neither covers, except card rows: `method` = `model`, `confidence` = its probability
   (two decimals), `ruleId` blank. Written only when `confidence` ≥ `MODEL_MIN_CONFIDENCE` (`.env`, default
   0.9). See [Model](#model).
4. Everything else is left out.

**Owner** = the `owner` column on `transactions` (bank-sync fills it from `account_owners`), unless the rule
sets one.

### Rules sheet (`categoriser_rules`)

| column | meaning |
|---|---|
| `ruleId` | unique; mined rules are `m-<hash>` |
| `active` | checkbox; only active rules categorise |
| `priority` | number, lower first; blank = 100 |
| `account` | `bank / accountName`, exact (case-insensitive) |
| `direction` | `in`, `out` or `any` |
| `counterpartyContains`, `descriptionContains` | whole words, case and punctuation ignored (`corner cafe` matches `Shop 12 - Corner_Cafe Ltd`) |
| `textRegex` | Python regex, case-insensitive, searched in `counterparty \| description \| reference` |
| `providerCategory` | the provider's `category` on `transactions`, exact (e.g. `SPENDING_CASHBACK`) |
| `amountMin`, `amountMax` | on the EUR magnitude (`abs(amountEur)`) |
| `owner` | blank = the account's owner |
| `type`, `category`, `subCategory` | must be a row of `categoriser_taxonomy` |
| `notes`, `source` | free text; `manual` or `mined` |
| `hits`, `lastMatched` | written by the job: rows the rule matches on its own (active or not), and the latest `bookingDate` |

Blank match columns mean "any", but at least one must be set. An active rule that breaks a check is skipped and
logged as `invalid_rule`; the run carries on. The job never edits your rules, only `hits` and `lastMatched`.

### Taxonomy sheet (`categoriser_taxonomy`)

`type | category | subCategory`, one allowed combination per row. Seed it once from the legacy sheet's
`Dropdowns` tab with `python -m categoriser.seed_taxonomy` (does nothing if it already has rows), then edit it
there. The rules sheet's dropdowns read from it.

### Mining and holdout accuracy (run by hand, never by the schedule)

These (and `categoriser.train`) need `LEGACY_SPREADSHEET_ID` in `~/categoriser/.env` and the legacy sheet shared with the service account
as Viewer.

- `python -m categoriser.mine [--dry-run]` appends suggestions to `categoriser_rules` with `active` = FALSE and
  `source` = mined. Labelled examples: legacy rows joined to a Bank Sync row (same amount, ≤ 3 days, ties broken
  by shared words), legacy rows that join to nothing, and rows you marked reviewed (these win). Per kind of
  transaction it tries counterparty words, then description words, then the provider category, and keeps the
  first needle with at least 3 examples of which 90% agree. Rows an existing rule already names get nothing new.
- `python -m categoriser.evaluate [--cutoff 2026-03-01]` mines only from examples before the cutoff and scores
  on labelled rows from the cutoff on. The report (coverage, exact and type accuracy, misses) goes to
  `~/categoriser/data/eval/`, never into the repo or the logs.

With Docker: `docker compose run --rm categoriser python -m categoriser.mine --dry-run`.

### Model

Logistic regression (scikit-learn) over the words and letter groups of counterparty, description and reference
(numbers dropped), plus direction and the size of the amount. It predicts a `type / category / subCategory` that
must be in the taxonomy. Owner is the account's owner.

- **Train** by hand, after reviewing rows or approving rules: from the runtime folder,
  `docker compose run --rm categoriser python -m categoriser.train`. Use the scheduled image: the model is saved
  with its scikit-learn version and the daily run skips a model from another version (`model=incompatible` in the
  log). scikit-learn is pinned in `pyproject.toml`; after bumping it, re-run `install-schedule.sh` and retrain.
- Examples are the same as for mining: legacy rows, Bank Sync rows joined to them and rows you reviewed
  (these count three times). Labels with fewer than 2 examples are not learnt.
- It first runs the holdout below, then trains on everything and writes `models/categoriser-model.joblib` and
  `models/categoriser-model.json` (when, how many examples per label, holdout scores). Never in the repo.
- **Holdout:** `categoriser.evaluate` (and `train`) also trains a model on rows before the cutoff and reports, per
  account, how many of the rows the rules leave it labels and how many correctly, for confidence 0.5 to 0.95.
  Pick `MODEL_MIN_CONFIDENCE` from that table; `--min-confidence` tries another value.
- **Card rows are skipped.** The hand-labelled history has few card rows like Trading 212's, and on them the model
  was often confidently wrong. Review a few rows per merchant instead: mining turns them into rules.
- **Daily run:** no model file → `model=missing`, unreadable or failing → `model=unreadable` / `failed`; the run
  carries on without it. `model_reviewed_checked` / `model_reviewed_wrong` count rows you reviewed that the model
  would label, and how many it gets wrong (reviewed rows are part of training, so this is optimistic; the
  holdout is the real test). Retraining is manual.

## Layout

```
bank-sync/categoriser/
  src/categoriser/  config.py  schema.py  merge.py (pure upsert plan)  sheets.py (gspread I/O)
                    rules.py (parse, validate, match)  transfers.py  engine.py (rules + transfers)
                    taxonomy.py  legacy.py (read + join)  mine.py  evaluate.py  seed_taxonomy.py
                    model.py (features, train, predict, save/load)  train.py
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
| The model in `~/categoriser/models/` (after `categoriser.train`) | No (loaded on every run) |
| Apps Script code (`src/`, `bank-sync/src/`) or anything else in the repo | No (not part of the categoriser; deploy with clasp as usual) |
| Column names in the `transactions` sheet | Check the code still matches, then yes if you changed it |

Re-running is safe: it never overwrites `.env`, keys, data or models already in `~/categoriser`.
