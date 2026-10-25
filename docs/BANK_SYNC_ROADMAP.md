# Bank Sync: roadmap and working state

Single source of truth for the `bank-sync/` project, so work can resume in a fresh
session. Read `CLAUDE.md` first (binding), then `docs/HANDOFF.md` (architecture rules:
SRP/DRY, JSDoc, provider abstraction, normalized models, FX design, idempotent
imports, Node-testable pure logic). This file records what has been built, what was
decided, and what comes next. Update it in the same PR that changes the state.

Last updated: 2026-10-25 (phase 4, daily sync and consent warnings, in review).

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
| 3 | FX (ECB via Frankfurter): `amountEur`/`fxRate`/`fxRateDate` on transactions and balances, `rates` tab | Merged (#56), live: first sync converted the full history, a second converted 0 |
| 4 | Automation and consent-expiry warnings | In review (draft PR): daily trigger, follow-up runs, last-run record, consent warning emails. Not pushed to Apps Script yet |
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
  field as `api.*` columns; `raw_transactions` accumulates, see below), `transactions` (the full deduplicated
  history, signed amounts, label columns, `amountEur`/`fxRate`/`fxRateDate` filled), and `rates` (ECB rates per
  currency and date). `raw_balances` also carries `amountEur`/`fxRate`/`fxRateDate` at the latest rate.
- Menu: *Connect a bank...*, *Sync all*, *Show status* (Phase 4 adds *Enable daily sync* and *Disable daily sync* once pushed). Sync all is idempotent: a second run
  reports 0 added, 0 updated, 0 removed, and 0 transactions converted to EUR.
- Apps Script: web-app deployment (the redirect `/exec`) is pinned to version 3; the script
  itself is pushed from `main` after #56 (Phase 3 did not change the web app, so no redeploy). The web app only changes when `auth.complete`,
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
  budget, failures isolated), `schedule` (pure trigger de-duplication and follow-up decisions),
  `consent-warnings` (which warnings are due, the email text, the sent record), `last-run` (outcome
  record for Show status), `auto-sync` (daily trigger on/off and the triggered run).
- `lib/`: `jwt`, `http` (retry/backoff, carries the API error code), `hash` (FNV-1a),
  `flatten` (API object to dotted keys, account numbers masked), `gas-adapters` (the only
  place touching Apps Script services directly, including ScriptApp triggers, MailApp and Session).
- `fx/`: `fx` (pure: rate lookup with previous-business-day fallback, conversion, missing
  ranges), `frankfurter` (rates client), `fx-service` (injectable: rates tab + cache + fetch,
  converts transactions and balances). `sheets/rates-store` reads and writes the `rates` tab.
- `providers/bank-provider.js`: the contract (`authorize`, `listAccounts`, `getBalances`,
  `getTransactions`) and the `Account`, `Balance`, `Transaction` typedefs.
- `providers/enable-banking/`: `eb-client` (one method per endpoint), `eb-auth`,
  `eb-token`, `eb-session-store`, `eb-normalize`, `eb-provider`.
- `sheets/`: `raw-writer` (fixed columns plus one `api.*` column per API field, per-connection
  replace, never a whole-tab clear), `transactions-store` (loads by header name).
- `entry/`: `container` (composition root; `syncAllDeps` is shared by the menu and the trigger), `menu`, `web-app`,
  `triggers` (the two global trigger handlers). Only these touch globals.
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

## Phase 3: FX (live)

As built (see `bank-sync/README.md`, "EUR conversion"): at the end of *Sync all*, `SyncAll` converts the fresh
balances (latest rate) before writing the raw tabs, then loads `transactions`, converts only the rows with
no `amountEur` at their booking-date rate (else value date, else latest) and saves only if something changed.
Converted rows are never recalculated; `txn-merge` clears the EUR columns when a re-fetch changes the amount,
currency or date, so such a row counts as new. `FxService` gets the rates
source, the `rates` tab store, the cache and the clock injected; the pure logic is in `fx/fx.js` and Node-tested
(`test/fx.test.js`). `amountEur = amount / fxRate`, rounded to cents; EUR rows get rate 1 and their own date.

- Fallback: the latest rate on or before the date, at most `FX.MAX_FALLBACK_DAYS` (7) older; else the row stays
  as it was and counts as "without a rate". Easter (Good Friday and Easter Monday) is the longest regular gap (4 days).
- Fetching: the stored series per currency is one contiguous block; a run fetches only its missing edges (older
  than the first stored date, or newer than the last up to today), one call each. A range asked for in the last
  6 hours is not asked again (CacheService marker), so weekends and runs before the afternoon publication do not
  refetch. Only the dates of rows being converted count, so old converted rows never widen a fetch.
- A row dated today (or later) is stored only with that day's own rate; until it is published it stays empty and
  counts as "waiting". A past date with no rate (weekend, holiday) is final with the previous business day.
- Cache: the whole rate table is cached as JSON for 6 hours (skipped above 90,000 characters); the `rates` tab is
  the durable copy.
- A failed fetch (e.g. HTTP 422 for a currency the ECB does not publish) is reported per currency in the summary;
  an FX error never stops the rest of the sync.
- The web app is unchanged (no `auth.complete`/`doGet` change): after `clasp push` no `clasp deploy` is needed.
  No new OAuth scope (`script.external_request` already covers UrlFetchApp). The first *Sync all* after the push
  fills the EUR columns for the whole history and adds three columns to `raw_balances`.

Frankfurter facts (verified live on 2026-10-25):
- Current API is v2: `GET https://api.frankfurter.dev/v2/rates?from=&to=&base=EUR&quotes=SEK&providers=ecb`
  returns a JSON array of `{date, base, quote, rate}`, one entry per ECB business day, `rate` = units of the quote
  per 1 base. v1 (`/v1/...`, ECB only) is deprecated but still works; `api.frankfurter.app` redirects.
- Without `providers=ecb`, v2 blends rates from many central banks and returns weekend dates, so it is required.
- A `from` on a weekend or holiday starts at the previous business day; a `to` after the latest publication is
  clipped; a range with no data (e.g. before 1999) returns `[]`; an unknown currency returns HTTP 422.
- ECB rates are published once a day around 16:00 CET; before that, today's date has no rate yet.

Observed on live Danske data: cross-border credits (EUR sent to the SEK account) state the bank's
own conversion only inside the description text ("Amount received: EUR 500.00 ... Exchange rate:
EUR/SEK ..."); the API's `exchange_rate` field is empty. Phase 3 uses the ECB rate per booking date;
the bank's stated rate is a possible cross-check, not a source.

## Phase 4: automation and consent expiry (in review)

As built (see `bank-sync/README.md`, "Daily sync and consent warnings"):

- *Enable daily sync* / *Disable daily sync* create or remove one daily time-driven trigger
  (`triggerDailySync`, 06:00-07:00 script time zone). `Schedule.plan` keeps exactly one per handler,
  so enabling twice or a stray duplicate never leaves two.
- The trigger runs the same `SyncAll.run` (deps built by `Container.syncAllDeps`, shared with the menu)
  under `Lock.run`. An unfinished run (time budget, or the lock busy) schedules one one-off follow-up
  (`triggerSyncContinue`, 5 minutes later), at most 3 per daily run; old one-off triggers are deleted
  at the start of each triggered run because Apps Script keeps them listed after they fire.
- Each run's outcome (counts and redacted errors only) is stored in `BS_LAST_RUN` and shown by
  *Show status*; a manual run is recorded too (also when it fails) and keeps the follow-up chain. A run that throws (other than a busy lock) is recorded and rethrown so Apps Script's
  own failure email fires too.
- Consent warnings at 30, 14, 7 and 1 days before expiry and once after it lapses, one email per
  threshold. Dedupe: `BS_CONSENT_WARNED_<connection>` = `{validUntil, sent: [..]}`; a new expiry date
  (renewed consent) resets it; an email is recorded only after it was sent. Reaching several thresholds
  at once (e.g. the daily sync turned on late) sends one email. Sent only by the daily sync, also when its
  sync fails; errors are collected per connection so one failure never silences the others.
- Recipient: `Session.getEffectiveUser().getEmail()` at run time (the trigger owner), never stored.
- New scopes, verified in the Apps Script reference: `script.scriptapp` (ScriptApp triggers),
  `script.send_mail` (MailApp.sendEmail), `userinfo.email` (Session user email). Added to
  `bank-sync/src/appsscript.json` and `scripts/allowed-scopes.json`; `check:manifest` now also checks
  bank-sync's manifest against the allow-list (before, it only checked `src/`).
- Deploy: `clasp push --force` from `bank-sync/` after diffing the manifest (time zone and scopes),
  then reload the sheet and re-authorize. The web-app path (`auth.complete`, `doGet`) is unchanged, so
  no `clasp deploy`. Then *Enable daily sync* once.

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
- `raw_transactions` accumulates one row per transaction id (a refetch replaces the older copy). Rows older than
  365 days (`RAW_RETENTION_DAYS`) are dropped on each sync to keep the tab bounded. A pending row
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
