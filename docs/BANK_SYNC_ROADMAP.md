# Bank Sync: roadmap and working state

Single source of truth for the `bank-sync/` project, so work can resume in a fresh
session. Read `CLAUDE.md` first (binding), then `docs/HANDOFF.md` (architecture rules:
SRP/DRY, JSDoc, provider abstraction, normalized models, FX design, idempotent
imports, Node-testable pure logic). This file records what has been built, what was
decided, and what comes next. Update it in the same PR that changes the state.

Last updated: 2026-10-25 (after phase 2.5 went live).

## Setup in one paragraph

Two independent Apps Script projects, each bound to its own Sheet:

- **Legacy** (`src/`, `.clasp.json` at repo root): sheet "Financial Planning". Plaid and
  SaltEdge code. **Not being developed further**; the smoke test of it was dropped
  because its `Transactions` sheet is empty. Retirement is a pending decision (below).
- **bank-sync** (`bank-sync/`, own `.clasp.json`, namespace `BankSync`): sheet "Bank Sync".
  Enable Banking is the data source. Europe/Luxembourg time zone, EUR reporting currency.

## Status

| Phase | Scope | State |
|---|---|---|
| 0 | Setup, folder layout, config, secrets handling | Merged (#40) |
| 1 | Enable Banking auth (RS256 JWT, redirect web app), Revolut accounts and balances to raw tabs | Merged (#42), live |
| 2 | Transactions: paginated fetch, normalization, idempotent upsert, `transactions` tab | Merged (#44), live-tested on Revolut (sync twice gives 0 added/updated/removed). Hardened in #48 after the first Claude review |
| 2.5 | Connections, Connect flow with live bank list, labels, `raw_*` tabs with every API field | Merged (#49), live on the Bank Sync sheet (Revolut and Danske Bank SE) |
| 3 | FX (ECB via Frankfurter) | Planned |
| 4 | Automation and consent-expiry warnings | Planned |
| 5 | Trading212 provider | Planned |
| 6 | CSV import provider (Avanza, BIL) | Planned |
| 7 | Danske Bank (Sweden) | Done with no code change: connected through *Connect a bank...*; more banks follow the same route |
| 8 | Wife's accounts (second holder at the same bank) | Last; already accounted for in the 2.5 design |

Rules for every phase: branch per task, one small PR per phase, squash merge, never
commit to `main`, conventional-commit PR title with a lowercase subject (e.g.
`feat(bank-sync): add transactions writer (phase 2)`), summary and test plan in the PR
body, open as draft and mark ready when CI is green (that starts the review), ask before merging, and stop after each phase for review.

## Live state (2026-10-25)

- Connections: `revolut-niklas` (EUR and SEK accounts on one masked IBAN, consent until
  2027-01-22) and `danske-bank-niklas` (one SEK account, consent until 2027-01-23). The Danske
  account sends a real name ("Danske Konto"); Revolut sends none, so its accounts show
  "Current account EUR" and "Current account SEK" (fallback: type and currency).
- Tabs on the Bank Sync sheet: `raw_accounts`, `raw_balances`, `raw_transactions` (every API
  field as `api.*` columns; transactions show only the latest fetch per connection) and
  `transactions` (the full deduplicated history, signed amounts, label columns, empty
  `amountEur`/`fxRate`/`fxRateDate` waiting for Phase 3).
- Menu: *Connect a bank...*, *Sync all*, *Show status*. Sync all is idempotent: a second run
  reports 0 added, 0 updated, 0 removed.
- Apps Script: web-app deployment (the redirect `/exec`) is pinned to version 3; the script
  itself is pushed from `main` after #49. The web app only changes when `auth.complete`,
  `doGet` or what they call change: then run `clasp deploy -i <deployment id>`.
- Known data quirks: Revolut returns one transaction in 90 days (a zero-amount card check);
  its SEK balance has a reference date in 2024 (dormant). Danske history starts at the 90-day
  lookback; older history may need a longer `INITIAL_LOOKBACK_DAYS` (check what the bank serves).
- `clasp push` asks about the manifest every time because the live manifest says
  Europe/Brussels; the repo says Europe/Luxembourg. Diff it, then `clasp push --force`.

## Current architecture (bank-sync)

Load order is `filePushOrder` in `bank-sync/.clasp.json` (entries include `src/`).

- `core/`: `config` (static, non-secret), `errors`, `redact`, `legacy-migration` (one-time, removable), `connections` (bank picker
  helpers and the connection store in Script Properties), `account-labels` (labels joined from
  the accounts list), `raw-sync` (collect and write the three raw tabs), `date-range` (pure
  windows), `txn-merge` (pure upsert), `txn-sync`, `sync-all` (all connections, shared time
  budget, failures isolated).
- `lib/`: `jwt`, `http` (retry/backoff, carries the API error code), `hash` (FNV-1a),
  `flatten` (API object to dotted keys, account numbers masked), `gas-adapters` (the only
  place touching Apps Script services directly).
- `providers/bank-provider.js`: the contract (`authorize`, `listAccounts`, `getBalances`,
  `getTransactions`) and the `Account`, `Balance`, `Transaction` typedefs.
- `providers/enable-banking/`: `eb-client` (one method per endpoint), `eb-auth`,
  `eb-token`, `eb-session-store`, `eb-normalize`, `eb-provider`.
- `sheets/`: `raw-writer` (fixed columns plus one `api.*` column per API field, per-connection
  replace, never a whole-tab clear), `transactions-store` (loads by header name).
- `entry/`: `container` (composition root), `menu`, `web-app`. Only these touch globals.
- Tests: `bank-sync/test/*.test.js` (`npm test` in `bank-sync/`), using the real modules
  loaded in push order by `test/load.js`.

### Behaviour worth knowing (Phase 2)

- Amounts are signed: the API gives a positive amount plus `credit_debit_indicator`;
  DBIT becomes negative.
- Transaction id: `entry_reference` (documented unique and immutable per account), else
  `transaction_id` (can be null and unstable), else a content hash plus an occurrence
  counter. Dedupe key: provider + accountId + transactionId.
- Pending rows inside a re-fetched window that the bank no longer returns are removed
  (booked under a new id, or cancelled). A booked row is never downgraded to pending.
- Statuses: BOOK is booked; PDNG and HOLD are pending; CNCL, RJCT, SCHD, OTHR are skipped.
- No resume cursor: each account restarts at its latest booked date minus 7 days; the first
  run reads 90 days back; fetches are 31-day windows, oldest first, saved after each;
  time budget 4.5 minutes.
- Text columns are forced to plain-text format so bank strings cannot become formulas.
- `amountEur`, `fxRate`, `fxRateDate` stay empty until Phase 3 and are preserved on re-sync
  when amount and currency are unchanged.

## Enable Banking facts (verified against the API reference)

- Base `https://api.enablebanking.com`; every call carries a short-lived JWT signed RS256
  with the application's private key; the application id is the `kid`.
- `GET /accounts/{uid}/transactions`: query `date_from`, `date_to` (inclusive, UTC),
  `continuation_key`, `transaction_status`, `strategy` (`default` | `longest`). Response
  `transactions` and `continuation_key` (null on the last page; only valid in-session).
- `entry_reference` is documented unique and immutable per account; `transaction_id` may be
  null and may change between fetches.
- Consent expires per bank (90-180 days) and must be renewed by logging in again.
  Revolut's current consent lapses 2027-01-22.
- Not yet verified: whether restricted mode lets a second person authorize their own login
  under this application (matters for Phase 8), and how far back Revolut serves history.

## Phase 2.5: connections, generic menu, labels

Motivation, from live use of Phase 2:

1. The menu is hardwired to Revolut (`BANK = BANKS.REVOLUT`). It must be config-driven.
2. Sessions are stored by bank key, so a second login at the same bank would overwrite
   the first. Introduce a **connection**: key = bank + holder (e.g. `revolut-niklas`), each
   with its own session, consent expiry and accounts. Adding a holder or bank becomes a
   config entry.
3. Tabs do not show which bank, connection or account a row belongs to.

As built (see `bank-sync/README.md` for use): connections are created by the *Connect a bank...*
menu flow (country, bank picked from Enable Banking's live `GET /aspsps` list, holder label)
and kept in Script Properties, not in config; one property per session; sessions keyed by
connection (the old `revolut` session is moved to `revolut-niklas` automatically, once);
menu = *Connect a bank...*, *Sync all*, *Show status*; labels `bank`, `connection`,
`holderName`, `accountName`, `ibanMasked` on all three tabs, joined from the accounts list and
stored on each transaction row so one connection's sync never blanks another's; `raw_accounts`
and `raw_balances` carry every API field as `api.*` columns (IBAN-looking values masked);
account details are read automatically when missing or older than 7 days; the tab loads by
header name so older layouts still load.

Verified in the API reference: `name` is the account **holder's** name (stored as
`holderName`), `details` is the user/bank-set label (`accountName`), `product` the product type,
and `GET /sessions/{id}` returns only account ids, so details come from
`GET /accounts/{id}/details`. Enable Banking has no endpoint listing the banks or accounts you
linked (`GET /application` returns name, countries, services only) and no hosted bank picker
(`aspsp` is required in `POST /auth`); "not enabled" is detected from `NO_ACCOUNTS_ADDED` /
`ACCESS_DENIED` or an empty account list. Not yet seen live: which of these a restricted
application actually returns for an unlinked bank.

Original scope notes:

- Config: `CONNECTIONS` (key, bank, holder label), replacing the single-bank assumption.
  Session store and pending-auth state keyed by connection key (migrate the existing
  `revolut` session to `revolut-niklas` without forcing a re-login).
- Menu: **Connect a bank...** (picker of configured connections), **Sync all** (one lock,
  one time budget, per-connection result so one lapsed consent does not block the rest),
  **Show status** (every connection with consent expiry).
- Labels on `raw_accounts`, `raw_balances` and `transactions`: `bank`, `connection`
  (including holder), `accountName`, `ibanMasked`. Bank and connection come from the
  provider; name and masked IBAN are joined from the accounts list at write time (no
  duplicated logic). Last four IBAN characters only.
- Custom account names: carry whatever the API sends (`name`, plus `product` and
  `details` when present) into the account model and the labels, with a clear fallback
  order. Verify on the live Revolut response which fields actually hold a user-given name.
- The `transactions` layout changes, so the tab is rebuilt on the next sync. That is safe
  because data is re-fetchable and the dedupe key is unchanged.
- Design for Phase 8 now: the holder is part of the connection key, labels show it, and
  nothing assumes one holder per bank.

## Phase 3: FX

Injectable `FxService`; ECB reference rates via Frankfurter (verify the current endpoint
before coding). Each transaction converts at its booking-date rate, falling back to the
previous business day for weekends and holidays; balances use the latest rate. Cache in a
`Rates` tab plus `CacheService`; one fetch per date range, never per row. Fills `amountEur`,
`fxRate`, `fxRateDate`. Pure conversion logic is tested in Node.

Starting points for the next session:
- `Transaction.amountEur/fxRate/fxRateDate` already exist; `txn-merge` preserves them while
  amount and currency are unchanged, and `transactions-store` already writes the columns.
- Balances: add the EUR value as derived columns on `raw_balances` (latest rate); the reporting
  currency is `Config.REPORTING_CURRENCY` ('EUR'); accounts are EUR and SEK so far.
- Verify the current Frankfurter endpoint and its date semantics before coding; fetch one date
  range per currency per run, never per row; keep `FxService` injectable (rates fetcher and
  cache passed in) so the conversion and fallback logic are pure and Node-tested.
- Wire it into `TxnSync` (convert new/changed rows before saving) and `SyncAll`; keep every
  Apps Script service behind `gas-adapters`/`container`.

Observed on live Danske data: cross-border credits (EUR sent to the SEK account) state the bank's
own conversion only inside the description text ("Amount received: EUR 500.00 ... Exchange rate:
EUR/SEK ..."); the API's `exchange_rate` field is empty. Phase 3 keeps the planned ECB rate per
booking date; the bank's stated rate is a possible cross-check, not a source.

## Phase 4: automation and consent expiry

Time-driven refresh guarded by `LockService`, resumable under the 6-minute limit. Track
consent expiry per connection and email a warning well before it lapses (Revolut:
2027-01-22). Bank logins and consent renewals are done by the owner; the code only warns.

## Phase 5: Trading212 provider

Own API key in Script Properties (pasted by the owner). Holdings and cash are modeled
explicitly, not as transactions. Adding it must not touch existing providers.

## Phase 6: CSV import provider

`CsvImportProvider` for banks Enable Banking does not cover: Avanza (not supported) and
BIL (not supported, parked). Same normalized `Transaction` model and the same upsert.

## Phase 7: Danske Bank (Sweden): done

Danske connected through *Connect a bank...* (country SE, bank picked from the live list,
holder label) with no code or config change, and synced with SEK data. BGL, Spuerkeess,
Raiffeisen and SEB follow the same route once they are linked in the Enable Banking Control
Panel (the Connect flow explains it when a bank is not yet enabled).

## Phase 8: second holder at the same bank (last)

A partner's own accounts at a bank already connected. With the 2.5 connection model this
is a new connection and a separate consent (their own login). Blocked on confirming that
Enable Banking's restricted mode allows it; if not, it needs its own application or
another route. Decide only when this phase starts.

## Starting a new session

1. Read `CLAUDE.md`, this file, then `docs/HANDOFF.md`.
2. `git fetch`, branch off `origin/main` (never commit to `main`), run `npm test` in
   `bank-sync/` and `npm run check` at the root to confirm a green baseline.
3. Do the next phase in one small PR, opened as a draft; mark it ready when CI is green, read the
   review comment and its `code-review` issues, fix the real ones on the branch, and stop after the
   phase. Do not re-run the review yourself.
4. Deploy only when asked: `clasp push --force` from `bank-sync/` after diffing the manifest,
   and `clasp deploy -i <deployment id>` only if the web-app path changed.

## Review workflow

Process set in `CLAUDE.md` (#51): open every PR as a **draft**, mark it **ready for review** only
when CI is green; that runs the checks on the PR head and starts the Claude review once. The review
posts one comment and files each bug and improvement as its own `code-review` issue (nits in one
issue); fix them on the same PR branch and close the issues. Do not re-run the review yourself
(no `workflow_dispatch`, no draft/ready toggling) unless the owner asks. Read the review before
merging: it found real bugs on #44 and #49. A PR that edits the workflow itself is skipped by design.

## Follow-ups from reviews (small, not blocking)

- Remove `LEGACY_CONNECTION`, `core/legacy-migration.js` and the single-map fallback in
  `EbSessionStore.loadSession` once the live Revolut session has been migrated (it is repeat-safe
  and repairs a half-finished run, so there is no urgency).
- `refreshStale` still runs inside `listAccounts` (it now receives the shared time check through
  `listAccounts({outOfTime})`); moving it to its own step in `SyncAll` would be cleaner but is not needed.
- Enable Banking keeps no list of enabled banks, so "not enabled" is detected after the fact;
  confirm live which error a restricted application gets for an unlinked bank.
- `raw_transactions` accumulates one row per transaction id (a refetch replaces the older copy). A pending row
  that later books under a new id stays as its own pending row there; the `transactions` tab drops it.
  Cancelled, rejected and scheduled rows are skipped by the normalizer, so they do not appear.

## Open decisions (raise, do not act)

- **One sheet:** move toward a single sheet. Options once Phase 2 is in: migrate the legacy
  views into Bank Sync, or split the repo.
- **Retire Plaid and SaltEdge** from `src/`: propose a removal PR.
- **`deploy.yml`** pushes only the legacy `src/` project. Do not extend it to bank-sync
  without asking.

## Deploy rules (learned the hard way)

- Use the globally installed `clasp`, not `npx clasp`.
- Run clasp from the folder of the project you mean: repo root for legacy, `bank-sync/`
  for bank-sync. Never from the other.
- Before every push, check the checked-out branch contains the code you intend to
  deploy; pushing from a branch without it overwrites the live script.
- After a bank-sync change that affects the web app, redeploy to the existing
  deployment with `clasp deploy -i <deployment id>` (see `clasp deployments`), otherwise
  `/exec` keeps serving the old version. Phase 2 did not change the web app.
- Run `npm run check` (repo root) and `npm test` (`bank-sync/`) before every push and PR.
- `clasp push` may ask to overwrite a changed manifest: diff the remote one first
  (`clasp clone` into a scratch folder) and only force if the difference is understood.
  The first Phase 2 push only corrected the time zone to Europe/Luxembourg.
- After a push the user reloads the sheet (menus are built in `onOpen`) and may need to
  re-authorize.

## Secrets and boundaries

- Never read, print or copy `.pem` files, Script Property values, `~/.clasprc.json` or
  tokens. The owner pastes secrets. Never put them in the repo, logs, sheets or chat.
- Bank logins and consent renewals belong to the owner: give the link and wait.
- No personal financial data in logs, commits, PR text or chat; IBANs stay masked; tests
  and docs use fake data.
- Ignore "[github assistant] ..." errors in the Apps Script editor (a Chrome extension).
- A `client_secret_*.json` sits (gitignored) in the repo root. The owner chose to leave it
  there for now; do not open it.

## Housekeeping done

Removed the merged remote branches `add-handoff-prompt` and `feat/enable-banking-auth-v2`,
turned on automatic deletion of merged branches, and dropped the untracked SaltEdge
account-filtering note.
