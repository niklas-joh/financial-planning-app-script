# categoriser

Python job (inside `bank-sync/`) that reads the **Bank Sync** Google Sheet and writes categories to a
separate `categorised` sheet: rules you keep in the `categoriser_rules` sheet, internal transfer pairing,
then a model for the rest, which labels a row only when it is confident. Everything else is left blank,
and the rows that most need you are listed in a review inbox (`categoriser_review`) you answer in the sheet.

> **This repo is public.** Keys, `.env`, CSVs, models and logs are git-ignored
> and blocked by the pre-commit hook (`../../scripts/hooks/pre-commit`). Never use
> `--no-verify` to get round it.

## Output contract

- Reads `transactions`. **Never writes to it** (enforced in `sheets._assert_writable`).
- Writes only to `categorised`, `categoriser_issues`, `categoriser_rules` (only `hits`, `wins`, `lastMatched`
  and `sample`, and rows appended by `categoriser.mine` or made in the review inbox), `categoriser_taxonomy`
  (created, and seeded once) and `categoriser_review` (the review inbox, rewritten each run). This list is
  `Settings.writable_sheets`. The legacy sheet is only opened with a read-only token.
- `categorised` is keyed by `transactionId`:

  | transactionId | owner | type | category | subCategory | method | confidence | ruleId | reviewed | updatedAt | bookingDate | counterparty | description | amountEur |
  |---|---|---|---|---|---|---|---|---|---|---|---|---|---|

  `method` is `rule`, `model` or `manual`; `reviewed` is a checkbox; `updatedAt` is ISO 8601 UTC;
  `confidence` is blank when not applicable. The last four columns are copied from `transactions` so that
  filtering on `ruleId` shows what a rule caught. A sheet from before these columns existed gets them added
  at the end on the next run (its rule and model rows are rewritten once, so their `updatedAt` moves). Rows the
  run doesn't rewrite (reviewed, manual) get only these four cells filled while they are all blank; their labels
  and `updatedAt` are never touched (`context_filled` in the log).
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
  fixed); `invalid_review_answer` (a ticked inbox answer that can't be applied); `review_conflicts_with_reviewed`
  (an inbox answer that differs from a row already reviewed on `categorised`).
- After writing, each run checks `categorised` against `transactions` (`coverage.py`) and logs
  `transactions`, `categorised` (rows with a full label), `uncategorised`, `orphaned`, `not_written` and
  `ids_not_unique` counts. It opens: `orphaned_categorised_row` (a `categorised` row whose transactionId is no
  longer on `transactions`; never deleted by the job, since it may hold your answer: delete it yourself);
  `categorised_row_not_written` (a row the run labelled that is missing after the write); and
  `transaction_id_not_unique` (one transactionId on several `transactions` rows, e.g. identical BIL CSV rows on
  two accounts: only the first is categorised until bank-sync gives them different ids).
- Rows of `categorised` below the data read as `reviewed` = FALSE (an unticked checkbox) and are not rows: the
  sheet grows past its first 1000 rows on its own when more are appended.
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
| `notes`, `source` | free text; `manual`, `mined` or `review` (made with "all like this" in the review inbox) |
| `hits` | written by the job: rows on `transactions` the rule matches **on its own**, ignoring other rules and `active` |
| `lastMatched` | written by the job: the latest `bookingDate` among those rows |
| `wins` | written by the job: rows the rule actually categorised (it was the first active rule to match) |
| `sample` | written by the job: the 3 most common counterparties among the `hits`, with counts, e.g. `Acme Market (12); ACME MKT ONLINE (3); +2 more` (description when the counterparty is blank) |

Blank match columns mean "any", but at least one must be set. An active rule that breaks a check is skipped and
logged as `invalid_rule`; the run carries on. The job never edits your rules, only the four job columns.

**Reading the job columns.** `hits` is not an exact-text count: one rule can catch several spellings and
merchants (`sample` shows which), and a row can count towards several rules. On an inactive (e.g. mined) rule,
`hits` is what it would catch if switched on; `wins` is 0. On an active rule, `hits - wins` is what a rule
earlier in the matching order took instead. To see every row a rule labelled, filter `categorised` on its `ruleId`.

**Layout.** Columns are found by header name: reorder them, or add your own, freely. A missing job column is added
at the end; a missing or repeated rule column stops the run before anything is written. The job adds a filter to
the sheet (once; an existing one is left as is), so it can be sorted, e.g. by `hits`. Sorting is safe: rules are
tried by `priority`, not by row.

### Taxonomy sheet (`categoriser_taxonomy`)

`type | category | subCategory`, one allowed combination per row. Seed it once from the legacy sheet's
`Dropdowns` tab with `python -m categoriser.seed_taxonomy` (does nothing if it already has rows), then edit it
there. The rules sheet's dropdowns read from it.

### Review inbox (`categoriser_review`)

Each daily run fills this tab with up to `REVIEW_SIZE` rows (default 25) that most need you: one row per
merchant with no category, merchants with the most such rows first, plus `REVIEW_SPOT_CHECKS` (default 3) rows
the model labelled, as spot checks of how often it is right. Card rows are included.

| column | meaning |
|---|---|
| `transactionId` … `description` | the transaction (`account` is `bank / accountName`) |
| `why`, `similar` | why it is here; how many rows of this merchant have no category |
| `confidence` | the model's probability for the pre-filled `answer` (blank without a model) |
| `answer` | one dropdown, `type › category › subCategory`; pre-filled with the model's best guess |
| `owner` | pre-filled with the account's owner |
| `accept` | tick when `answer` and `owner` are right |
| `allLikeThis` | also tick to make a rule for every row of this merchant |
| `matchText` | the words that rule matches (shorten them to widen it) |
| `suggested`, `kind` | the model's first guess, and `merchant` or `spot check` (filled by the job) |

The next run (answers are not picked up before then):

- Each ticked row becomes a `categorised` row with `method` = `manual` and `reviewed` = TRUE, and leaves the inbox.
- With `allLikeThis`, an **active** rule is added to `categoriser_rules`: `ruleId` `r-…`, priority 50 (ahead of
  mined rules), `counterpartyContains` = `matchText` (or `descriptionContains` when the row has no counterparty),
  the row's direction, any account, `source` = `review`. It applies in the same run.
- An answer that can't be applied (not in the taxonomy, unknown owner, row gone, on several `categorised` rows,
  different from an already reviewed row, or `allLikeThis` with a `matchText` shorter than 3 characters or not
  made of whole words of the row's counterparty/description) stays ticked in the inbox with an issue until you
  fix it.
- Rows you have not ticked stay as they are, including an answer you typed; rows that got a label drop out.
- If you tick or edit a row (`answer`, `owner`, `accept`, `allLikeThis`, `matchText`) while the job is running,
  it leaves the inbox as it is that day (`review=skipped_busy` in the log), so the edit is not lost.
- Each applied answer adds one line to `~/categoriser/data/review-history.jsonl` (what the model suggested and
  your answer), used later to measure the model. It stays in the runtime folder.

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
  log). scikit-learn is pinned in `pyproject.toml`; after a bump reaches `main`, the next run rebuilds the
  image and skips the old model (`model=incompatible`) until you retrain with the new image.
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
                    review.py (review inbox: answers and refill, pure)
                    coverage.py (after-write check of categorised vs transactions, pure)
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
was off, `RunAtLoad` runs it at login. `run-scheduled.sh` starts Docker Desktop if needed, skips if
today's run already succeeded, updates itself from `main` (below) and runs the `categoriser:local` image.
It never reads your checkout in `~/Documents`.

Install once (from Terminal or SSH):
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

## 5. After you change something: auto-update from `main`

Merging to `main` is the deploy. Before each daily run, `run-scheduled.sh` fetches `main` of this (public)
repo into `~/categoriser/repo` (only `bank-sync/categoriser/`). If that folder's git tree differs from the
one the image was built from (an image label), it rebuilds `categoriser:local` and refreshes
`docker-compose.yml`, `README.md` and `run-scheduled.sh` itself (used from the next run). So a change
merged today goes live at the next 06:15 run. The log line `run start (<commit>)` says which commit ran.

- **Only green commits deploy.** A new `main` commit is used once GitHub shows its `Tests (categoriser)` check
  passed (CI runs on every push to `main`). Pending, red or unreachable: the current image runs and the log says
  `check has not passed`; the next run tries again. A merge just before 06:15 may therefore go live a day later.
- Offline, fetch or build failure: the current image runs as before (`auto-update did not complete` in the log).
- Dependencies are capped below their next major version in `pyproject.toml`, so a rebuild picks up bug fixes but
  not a breaking release CI never saw.
- Changes elsewhere in the repo don't trigger a rebuild.
- A changed `launchd/*.plist.template` (e.g. the schedule time) is only reported (`WARNING: the launchd schedule
  changed`): re-run `install-schedule.sh` for that.
- To run a branch instead (testing): set `CATEGORISER_AUTO_UPDATE=0` in `~/categoriser/.env`, run
  `install-schedule.sh` from that branch, and remove the line again when done. Without it, the next run goes
  back to `main`.
- To go live now instead of tomorrow: `rm -f ~/categoriser/logs/.last-success && launchctl kickstart gui/$(id -u)/lu.niklas.categoriser`.

| You changed | Goes live |
|---|---|
| `bank-sync/categoriser/**` merged to `main` (code, `Dockerfile`, `deploy/*`, `run-scheduled.sh`) | Next daily run |
| `launchd/*.plist.template` | After re-running `install-schedule.sh` |
| `~/categoriser/.env`, or the key in `~/categoriser/credentials/` | Next run (read on every run) |
| The model in `~/categoriser/models/` (after `categoriser.train`) | Next run (loaded on every run) |
| Apps Script code (`src/`, `bank-sync/src/`) | Never via this job; deploy with clasp as usual |
| Column names in the `transactions` sheet | Check the code still matches |

Re-running is safe: it never overwrites `.env`, keys, data or models already in `~/categoriser`.
