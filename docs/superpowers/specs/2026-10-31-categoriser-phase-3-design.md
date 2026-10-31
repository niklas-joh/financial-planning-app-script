# Categoriser phase 3: review inbox and a model that improves itself

Date: 2026-10-31. Status: design agreed with the owner (all four decisions as proposed); 3a is built first.

## Why

Phases 1 and 2 (#104, #105) label rows with rules, transfer pairing and a model, but nothing new ever
reaches them:

- 0 of the 999 rows in `categorised` are reviewed, and 0 of the 80 rules are active (73 mined, 7 manual).
- 972 of the 1,667 transactions have no label, 756 of them card rows, which the model skips.
- A row without a label is not on `categorised` at all, so correcting it means adding it by hand.

Retraining every night on the same data gives the same model. The owner asked for a categoriser that
improves itself. That needs feedback first, so phase 3 has two parts:

- **3a:** a review inbox that makes answering a row take seconds on a phone.
- **3b:** the daily job learns from those answers on its own.

## Goals

- **3a:** each morning a short list of the rows that most need the owner. Answering one is a tick, or one
  dropdown and a tick. One answer can cover a whole merchant ("all like this").
- **3a:** every answer becomes ground truth: a reviewed `manual` row in `categorised`.
- **3a:** card rows (Trading 212) are reviewable, so they get covered.
- **3b:** the model retrains on its own when there is something new to learn, and only replaces the current
  model when it is not worse.
- **3b:** the owner sets an accuracy target and the job picks the confidence cut-off from it.
- **3b:** card rows are labelled by the model once reviews show it is accurate on them.
- **3b:** rule suggestions are mined on their own each week.

## Non-goals

- No change to Apps Script (`src/`, `bank-sync/src/`).
- No instant processing: answers are picked up by the next daily run (06:15).
- No new model type or feature work; 3b retrains the phase 2 model as it is.
- No UI beyond the Google Sheet (works in the Sheets phone app).

## Anti-goals (a failure even if it works)

- Writing to `transactions`, the legacy sheet, or a reviewed row.
- An answer the owner gave being lost or silently ignored.
- A worse model replacing a better one without the owner seeing it.
- The daily run stopping because of a bad answer, a bad rule or a failed training.
- Personal financial data in logs or the repo.

## Constraints (from CLAUDE.md and the existing code)

- The job writes only to `Settings.writable_sheets`; the review tab is added to that list.
- The legacy sheet is read only by hand-run commands. 3b's daily retraining therefore reads a **local
  snapshot** of it in `~/categoriser/data/`, written by a hand-run command (CLAUDE.md is updated to say so).
- Problems are logged once in `categoriser_issues` and never stop the run.
- Tests use synthetic data only (`tests/builders.py`); `pytest` without network.
- Deploying = re-running `install-schedule.sh` after the PR merges.

---

## 3a: review inbox (first PR)

### The `categoriser_review` tab

Created by the job (setting `REVIEW_SHEET`, default `categoriser_review`). One row per transaction to
review, sorted by `similar` (most first), spot checks last:

| column | filled by | meaning |
|---|---|---|
| `transactionId` | job | key |
| `bookingDate` | job | |
| `account` | job | `bank / accountName` |
| `amountEur` | job | signed, as on `transactions` |
| `counterparty` | job | |
| `description` | job | |
| `why` | job | e.g. `14 rows like this have no category`, `spot check: model 96% sure` |
| `similar` | job | rows with the same merchant key and no label (incl. this one) |
| `confidence` | job | the model's probability for `answer` as filled in (blank without a model) |
| `answer` | job, then owner | one dropdown: `type › category › subCategory`, pre-filled with the model's best guess |
| `owner` | job, then owner | pre-filled with the account's owner; dropdown from `account_owners` |
| `accept` | owner | checkbox: "this answer is right" |
| `allLikeThis` | owner | checkbox: also make a rule for every row like this |
| `matchText` | job, then owner | the words the rule will match (pre-filled; shorten it to widen the rule) |

- **One `answer` dropdown instead of three.** On a phone, one pick from the full list is quicker than three
  dropdowns that don't depend on each other. The list is every taxonomy triple joined with ` › ` (not `/`,
  because `Wants/Pleasure` holds a slash), rebuilt each run. Not strict: the job validates. *(Builder's
  call; if the taxonomy ever outgrows a dropdown list, switch the validation to a helper range.)*
- The model's guess is shown for **every** row, card rows included, even below the cut-off. It is only a
  suggestion; nothing is written until the owner ticks `accept`.

### What the daily run does with answers (before categorising)

Answers are read first, so a rule made with "all like this" already applies in the same run.

For each row with `accept` = TRUE:

1. **Check it.** The answer must be a taxonomy triple, the owner a known owner, and the transactionId on
   `transactions`. If not, the row stays in the inbox, still ticked, and an `invalid_review_answer` issue is
   logged (keyed by transactionId). It resolves on its own once fixed.
2. **Ground truth.** Write a `categorised` row: `method` = `manual`, `reviewed` = TRUE, `confidence` blank,
   `ruleId` blank (or the new rule's id, see 3), the chosen owner and triple. A row already reviewed on
   `categorised` is never overwritten. If it has the same owner and triple, the answer is simply done. If it
   differs, the inbox row stays, still ticked, with a `review_conflicts_with_reviewed` issue, until the owner
   deletes or unticks it or changes the reviewed row (the issue then resolves on its own).
3. **All like this.** If ticked and `matchText` is not blank, append an **active** rule to
   `categoriser_rules`:
   - `ruleId` = `r-<hash of field, matchText, direction>`
   - `active` = TRUE, `priority` = 50 (ahead of mined rules at the default 100)
   - `counterpartyContains` = `matchText` when the row has a counterparty, else `descriptionContains`
   - `direction` = the row's direction; no account (a merchant is the same on every account)
   - outputs = the answer and owner (owner left blank when it equals the account's owner)
   - `source` = `review` (a new value for the `source` dropdown), `notes` = `from review <date>`
   - If a rule with the same ruleId already exists, nothing is added.

   The rule then goes through the normal rule checks, so a bad `matchText` shows up as `invalid_rule`.
4. Remove the row from the inbox.

Then the normal run (rules, transfers, model) runs. A transaction with a manual answer is taken out of the
rule and model output, so it is written once, as the manual row.

### Refilling the inbox (after categorising)

The inbox is rebuilt each run to `REVIEW_SIZE` rows (default 25):

1. **Keep** ticked rows that could not be applied (invalid or conflicting answers), and unanswered rows
   unless the transaction now has a label (a new rule covers it, or it was reviewed on `categorised`). An
   answer typed but not ticked survives.
2. **Candidates:** transactions with no row on `categorised`, or a cleared one (no label), and not reviewed.
   Low-confidence model guesses are candidates too, since they are not written.
3. **Group by merchant key:** the normalised counterparty without numbers, or the description when there is
   no counterparty, plus the direction. `similar` = the group's size.
4. **Pick one row per group**, biggest groups first (most recent row of the group), until the inbox has
   `REVIEW_SIZE - REVIEW_SPOT_CHECKS` rows from this step. A group already in the inbox is not added again.
5. **Spot checks:** fill the remaining `REVIEW_SPOT_CHECKS` slots (default 3) with random rows the model
   labelled this run (`method` = `model`), from groups not in the inbox. `why` says it is a spot check. They
   measure the model on rows it has not seen.
6. **Pre-fill** `answer` and `confidence` with the model's best guess (if a model is loaded), `owner` with the
   account's owner and `matchText` with the merchant key.

**Writing safely:** the inbox is rewritten as a whole (values only; the dropdowns and checkboxes are
re-applied). Just before writing, the `transactionId` and `accept` columns are read again. If any row was
ticked since the run read the tab, the rewrite is skipped this run (`review=skipped_busy` in the log), so a
tick made while the job runs is never lost; it is picked up the next morning.

### Review history (for 3b)

Each processed answer appends one line to `~/categoriser/data/review-history.jsonl` (outside the repo):
date, transactionId, the suggested triple and confidence, the answer triple, whether it was a spot check and
whether the account is a card account. 3b uses it for spot-check accuracy. This data cannot be rebuilt
later, so 3a writes it even though 3a does not read it.

### Settings and logging

- New settings: `REVIEW_SHEET` (`categoriser_review`), `REVIEW_SIZE` (25), `REVIEW_SPOT_CHECKS` (3).
- `Settings.writable_sheets` gains the review sheet; all sheet names must still differ.
- Log line gains counts only: `review_answers`, `review_invalid`, `review_rules_added`, `review_changed`
  (answer differs from the suggestion), `review_spot_checks_right`, `review_spot_checks_wrong`, `review_size`,
  `review=written|skipped_busy`.
- New issue types: `invalid_review_answer`, `review_conflicts_with_reviewed`.

### Tests (synthetic)

Pure planning code (no gspread) with tests for: answer checks, manual rows, the "all like this" rule and its
id, manual answers winning over rule and model rows, keeping unanswered rows, grouping and picking, spot
checks, pre-fill, the busy-skip, and a full run on the fakes (two runs: the second changes nothing).

---

## 3b: the model improves itself (second PR)

### Retraining in the daily run

- **When:** after the answers are applied and before categorising, if any of:
  - no model, or an incompatible one;
  - at least `RETRAIN_MIN_NEW_REVIEWS` (10) reviewed rows more than the current model was trained with
    (`reviewedExamples` in the model's JSON);
  - the model is older than `RETRAIN_MAX_AGE_DAYS` (7).
- **Data:** the same labelled examples as `categoriser.train` (`mine.labelled`), with legacy rows read from
  the local snapshot instead of the legacy sheet. No snapshot: the step is skipped (`retrain=no_snapshot`).
- **Never fatal:** any failure is logged (`retrain=failed`) and the run carries on with the current model.

### Legacy snapshot

- New hand-run command `python -m categoriser.snapshot_legacy` reads the legacy sheet (read-only token, as
  today) and writes the parsed labelled rows to `~/categoriser/data/legacy-snapshot.json`.
- `mine`, `evaluate` and `train` refresh the snapshot when they read the legacy sheet anyway.
- The daily run only ever reads the file. CLAUDE.md and the README are updated to say the daily job uses a
  local copy of the legacy rows and never opens the legacy sheet.

### Accepting a new model

1. Holdout: train on labelled rows before a cutoff, score on rows after it. The cutoff rolls: 90 days before
   the newest labelled row, so new reviews are in the test set.
2. **Calibrate:** the cut-off is the lowest confidence (0.50 to 0.99, steps of 0.01) at which at least
   `MODEL_TARGET_ACCURACY` (0.97) of the holdout rows at or above it are exactly right, with at least 30 such
   rows. If none qualifies, the cut-off is 1.0 (the model labels nothing) and an issue is logged.
3. **Gate:** the candidate replaces the current model only if its holdout accuracy at its cut-off is at most
   1 point below the current model's recorded score and it covers at least as many holdout rows. *(Builder's
   call: the two holdout sets differ by the newest rows, so an exact "at least as accurate" would reject
   good models on noise.)* A rejected candidate is logged as a `model_rejected` issue with both scores.
4. Then train on all labelled rows and save it with the cut-off(s) and scores in the model's JSON. The daily
   run uses the saved cut-off; `MODEL_MIN_CONFIDENCE` is only the fallback for a model without one.
5. **Keep the last 3 models** in `models/archive/<trainedAt>/`. Roll back by copying one back (documented in
   the README; no command).

### Card rows

The same calibration runs on reviewed card rows alone. When there are at least 30 holdout card rows and a
card cut-off qualifies, the model labels card rows at that cut-off (`cardThreshold` in the model's JSON).
Until then card rows stay skipped, as today.

### Spot-check accuracy

From `review-history.jsonl`: of the last 30 spot checks, how many the model got right. Logged each run
(`spot_check_accuracy`). Below the target by more than 5 points, a `model_accuracy_low` issue is opened.
*(Builder's call: report it, don't act on it automatically.)*

### Weekly mining

Once a week (date of the last mining kept in `~/categoriser/data/state.json`), the daily run mines rule
suggestions from the snapshot and reviewed rows, appended inactive with `source` = `mined`, as the hand-run
command does today. Rules made with "all like this" are active already.

### Settings

`MODEL_TARGET_ACCURACY` (0.97), `RETRAIN_MIN_NEW_REVIEWS` (10), `RETRAIN_MAX_AGE_DAYS` (7).

---

## Left to the builder

- Module layout (e.g. `review.py` pure planning, I/O in `sheets.py`), names of helpers, exact `why` texts.
- How the inbox's dropdown list is built and re-applied.
- Test structure, within the repo's conventions.

## Calls made without the owner (check while reading)

1. One `answer` dropdown (`type › category › subCategory`) instead of three dropdowns.
2. Rules from "all like this": priority 50, no account, `source` = `review`.
3. An answer that differs from an already-reviewed row is not applied; it stays in the inbox with an issue
   so the owner chooses.
4. The inbox rewrite is skipped for a run when a row is ticked while the job runs.
5. 3a already records review history, because it can't be rebuilt later.
6. The 3b gate allows 1 point of noise; spot-check accuracy only raises an issue.
7. Rollback is a documented file copy, not a command.
