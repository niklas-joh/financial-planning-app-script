# Bank Sync: roadmap and working state

Single source of truth for the `bank-sync/` project, so work can resume in a fresh
session. Read `CLAUDE.md` first (binding), then `docs/HANDOFF.md` (architecture rules:
SRP/DRY, JSDoc, provider abstraction, normalized models, FX design, idempotent
imports, Node-testable pure logic). This file records what has been built, what was
decided, and what comes next. Update it in the same PR that changes the state.

Last updated: 2026-10-25 (phase 6, the CSV import provider with BIL, as it stands once merged and deployed).

## Setup in one paragraph

Two independent Apps Script projects, each bound to its own Sheet:

- **Legacy** (`src/`, `.clasp.json` at repo root): sheet "Financial Planning". Plaid and
  SaltEdge code. **Not being developed further**; the smoke test of it was dropped
  because its `Transactions` sheet is empty. Retirement is a pending decision (below).
- **bank-sync** (`bank-sync/`, own `.clasp.json`, namespace `BankSync`): sheet "Bank Sync".
  Enable Banking (banks), the Trading 212 public API (one API key per account) and CSV exports in a
  Drive folder per holder (BIL) are the data sources.
  Europe/Luxembourg time zone, EUR reporting currency.

## Status

| Phase | Scope | State |
|---|---|---|
| 0 | Setup, folder layout, config, secrets handling | Merged (#40) |
| 1 | Enable Banking auth (RS256 JWT, redirect web app), Revolut accounts and balances to raw tabs | Merged (#42), live |
| 2 | Transactions: paginated fetch, normalization, idempotent upsert, `transactions` tab | Merged (#44), live-tested on Revolut (sync twice gives 0 added/updated/removed). Hardened in #48 after the first Claude review |
| 2.5 | Connections, Connect flow with live bank list, labels, `raw_*` tabs with every API field | Merged (#49), live on the Bank Sync sheet (Revolut and Danske Bank SE) |
| 3 | FX (ECB via Frankfurter): `amountEur`/`fxRate`/`fxRateDate` on transactions and balances, `rates` tab | Merged (#56), live: first sync converted the full history, a second converted 0 |
| 4 | Automation and consent-expiry warnings | Merged (#58), live: pushed, re-authorized, daily sync enabled. First daily run and first warning email not seen yet (see Live state) |
| 5 | Trading212 provider: cash, investments, holdings, daily values, cash movements and dividends | Merged (#64), live since 2026-10-25 (first sync: 336 transactions, 19 holdings, 1 daily value). Fix (#65): daily total from Trading 212's own total, one-time full history backfill |
| 6 | CSV import provider: Drive folder per holder, BIL export | Merged, live once deployed (see Phase 6). Avanza follows after B1 |
| 7 | Danske Bank (Sweden) | Done with no code change: connected through *Connect a bank...*; more banks follow the same route |
| 8 | Wife's accounts (second holder at the same bank) | Last; already accounted for in the 2.5 design |

Rules for every phase: branch per task, one small PR per phase, squash merge, never
commit to `main`, conventional-commit PR title with a lowercase subject (e.g.
`feat(bank-sync): add transactions writer (phase 2)`), summary and test plan in the PR
body, open as draft and mark ready when CI is green (that starts the review), ask before merging, and stop after each phase for review.
**Update this file in the same PR as the phase**, written as the state after merge and deploy (status
row, live state, deploy steps), so no follow-up docs PR is needed. Anything only seen live (counts,
quirks) goes into the next phase's PR.

## Live state (2026-10-25)

- Connections: `revolut-niklas` (EUR and SEK accounts on one masked IBAN, consent until
  2027-01-22) and `danske-bank-niklas` (one SEK account, consent until 2027-01-23). The Danske
  account sends a real name ("Danske Konto"); Revolut sends none, so its accounts show
  "Current account EUR" and "Current account SEK" (fallback: type and currency).
- Tabs on the Bank Sync sheet: `raw_accounts`, `raw_balances`, `raw_transactions` (every API
  field as `api.*` columns; `raw_transactions` accumulates, see below), `transactions` (the full deduplicated
  history, signed amounts, label columns, `amountEur`/`fxRate`/`fxRateDate` filled), and `rates` (ECB rates per
  currency and date). `raw_balances` also carries `amountEur`/`fxRate`/`fxRateDate` at the latest rate.
- Menu: *Connect a bank...*, *Sync all*, *Import CSV files*, *Show status*, *Enable daily sync*, *Disable daily sync*. Sync all is idempotent: a second run
  reports 0 added, 0 updated, 0 removed, and 0 transactions converted to EUR.
- Apps Script: web-app deployment (the redirect `/exec`) is pinned to version 3; the script
  itself is pushed from `main` after #58 (Phases 3 and 4 did not change the web app, so no redeploy). The three
  Phase 4 scopes (`script.scriptapp`, `script.send_mail`, `userinfo.email`) are authorized.
- Daily sync: on (one trigger, 06:00-07:00 Europe/Luxembourg, owned by the owner's account). On 2026-10-25
  *Show status* said "Last sync: none recorded yet"; the first daily run is due 2026-10-26. Both consents had
  89 days left; the first warning emails (30 days) are due around 2026-12-23 (Revolut) and 2026-12-24 (Danske).
  Check after the first daily run that *Show status* shows `(daily): OK` with 0 added/updated/removed. The web app only changes when `auth.complete`,
  `doGet` or what they call change: then run `clasp deploy -i <deployment id>`.
- Trading 212 (Phase 5): `trading212-niklas` is live (key and secret pasted 2026-10-25); the partner's
  pair is added once she has her own key. The first *Sync all* added the `holdings` and `daily_values`
  tabs and read 90 days of history; with the fix (#65) the rest of the history backfills over the next
  syncs (*Show status* shows "history: complete" when done). Seen on the first live sync:
  - `investments.currentValue` already contains the cash in pies (the holdings sum plus `cash.inPies`), so
    the first version counted pie cash twice; and `totalValue` was slightly above available cash plus
    investments (an amount the summary does not break down, possibly payments still being processed). The
    fix takes `totalValue` as the total and cash as total minus investments.
  - History types: `INTEREST_ON_FREE_CASH` (about two a day), `DEPOSIT`, `WITHDRAW`, `LENDING_INTEREST`, and
    dividends. Card payments do appear, but only as plain `WITHDRAW` rows (cashback as small `DEPOSIT` rows) with
    no card type, merchant or category, so they cannot be told apart from other movements (owner, checked
    against the app on 2026-10-25; this corrects the first reading that they were missing). See B3.
  - A key without its secret shows as "not connected" (both are needed).
  - The backfill (#65) loaded the full history in two syncs. Every re-read Trading 212 row was reported as
    "updated" (the values were identical; only the key order differed from the copy read back from the sheet,
    and `TxnMerge.changed` compared JSON text). Fixed in #67: rows are compared field by field. After deploying
    it, a repeated sync reports 0 added and 0 updated for Trading 212 too.
  - `totalValue` stays a constant small amount above available cash plus investments over several syncs, so it
    is not a payment in flight; it counts as cash (total minus investments).
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
  `getTransactions`, plus the optional `maxWindowDays` and `getHoldings`) and the `Account`, `Balance`,
  `Transaction`, `Holding`, `DailyValue` typedefs.
- `core/provider-registry.js`: one entry per provider (its connections, whether one can sync, how to build
  its provider, its Show status lines and expiring consents). Sync all, Show status and the consent warnings
  loop over it; a new provider is a new entry. `core/daily-values.js`: pure daily value per account from the
  balances tagged `cash`/`investments`.
- `providers/trading212/`: `t212-normalize` (pure), `t212-rate-limit` (pure 429 wait from
  `x-ratelimit-reset`, plugged into `lib/http` as `retryDelayMs`), `t212-credentials` (key pairs in Script
  Properties), `t212-client` (one method per endpoint, Basic auth), `t212-provider`.
- `providers/csv/`: `csv-parse` (pure: rows, header below a preamble, numbers and dates in a declared layout),
  `csv-formats` (one entry per export, written from a real sample), `csv-normalize` (pure: format detection,
  account, stable ids, number masking), `csv-folders` (folder per holder and the last-read record in Script
  Properties), `csv-provider`. `lib/gas-adapters` `driveFolder` is the only DriveApp use.
- `providers/enable-banking/`: `eb-client` (one method per endpoint), `eb-auth`,
  `eb-token`, `eb-session-store`, `eb-normalize`, `eb-provider`.
- `sheets/`: `raw-writer` (fixed columns plus one `api.*` column per API field, per-connection
  replace or accumulate by key, never a whole-tab clear; also the `holdings` and `daily_values` columns),
  `transactions-store` (loads by header name).
- `entry/`: `container` (composition root; `syncAllDeps` is shared by the menu and the trigger, and hands
  providers the Apps Script helpers: HTTP, Base64, local dates, time left), `menu`, `web-app`,
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

## Phase 4: automation and consent expiry (live)

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
- Deployed 2026-10-25: `clasp push --force` from `bank-sync/` after diffing the manifest (time zone and
  scopes), sheet reloaded and re-authorized, *Enable daily sync* run once. The web-app path
  (`auth.complete`, `doGet`) is unchanged, so no `clasp deploy`.
- Not yet seen live: a daily run's record, a follow-up run, and a warning email.

## Phase 5: Trading212 provider (live once deployed)

As built (see `bank-sync/README.md`, "Trading 212"):

- One connection per API key pair in Script Properties, `T212_API_KEY_<HOLDER>` and `T212_API_SECRET_<HOLDER>`
  (found with `getKeys()`, holder slugged: `trading212-niklas`, `trading212-niklas-isa`). Pairs are never
  stored elsewhere; *Show status* shows `set`/`MISSING`; `Redact` scrubs `Basic ...` headers.
- `listAccounts` builds one account from `GET /equity/account/summary` (its `id` is the `accountId`). The
  summary is read once per run (the limit is 1 per 5 s) and also gives the five balances:
  `cash.availableToTrade`, `cash.inPies`, `cash.reservedForOrders` (detail, no kind),
  `investments.currentValue` (`valueKind: investments`; it includes the cash in pies) and `totalValue`
  (`valueKind: total`).
- `getHoldings` (new optional contract method) maps `GET /equity/positions` to `Holding` rows; `amount` is the
  value in the account currency so the unchanged `FxService.convertBalances` converts it.
- `daily_values` is built in `RawSync.write` from the converted balances by `DailyValues.fromBalances` (no second
  FX path); accumulated by connection, account and date, so a same-day re-run replaces the row. Rule: with a
  `total` balance, total is that and cash = total - investments (so it matches the provider); without one,
  cash = sum of `cash` balances and total = cash + investments. Banks can be added later by tagging their
  booked balance `cash` in the Enable Banking normalizer.
- `getTransactions` reads `/equity/history/transactions` and `/equity/history/dividends` newest first
  (`limit=50`, following `nextPagePath`) until a page reaches back past `dateFrom`; `maxWindowDays: Infinity`
  makes `TxnSync` ask for one window, so its lookback, restart and idempotent upsert are reused unchanged.
  Withdrawals and fees negative; deposits, interest and dividends positive; TRANSFER and unknown types keep the
  API's sign. Dividend ids are `div:<reference>`. Orders are not imported.
- One-time full history backfill (`t212-history`): on top of the incremental read, each feed continues from its
  saved `nextPagePath` (Script Property `T212_HISTORY_<HOLDER>`, `{transactions|dividends: {next, done}}`) for up to
  `TRADING212.BACKFILL_PAGES_PER_RUN` (40) pages per run; a rate limit, HTTP error or the time budget pauses it
  until the next sync, and the positions are saved only after every feed's items are read and normalized (a
  failure saves nothing, so the same pages are read again rather than skipped). When both feeds reach
  their last page it stops; deleting the property starts it again. Older rows go through the same upsert.
- Rate limits: no client-side pacer (a run makes about four calls per account, far below the burst limits).
  `lib/http` gained an optional `retryDelayMs` hook: for Trading 212 a 429 waits until `x-ratelimit-reset`
  (seconds or ms accepted), at most `TRADING212.MAX_RETRY_WAIT_MS` (65 s) and never past the run's time budget;
  Enable Banking keeps the plain backoff.
- No new OAuth scope (`script.external_request` covers UrlFetchApp; `Utilities.base64Encode` needs none). The
  web-app path (`auth.complete`, `doGet`) is unchanged, so `clasp push --force` without `clasp deploy`.
- Deploy steps: the owner creates each key (permissions `account`, `portfolio`, `history:transactions`,
  `history:dividends`, no IP restriction), pastes the pairs, `clasp push --force` from `bank-sync/` after diffing
  the manifest, reloads the sheet, runs *Show status* and *Sync all*.

Trading 212 facts (verified against docs.trading212.com/api and its OpenAPI file, 2026-10-25):

- Base URLs: live `https://live.trading212.com/api/v0`, demo `https://demo.trading212.com/api/v0`. Beta; Invest and
  Stocks ISA accounts only; multi-currency accounts not supported; amounts in the primary account currency
  (instrument prices in the instrument currency).
- Auth: HTTP Basic, API key as user and API secret as password. A legacy scheme (`Authorization: <key>`) is still
  listed, which is why the old portfolio sheet works. Keys are scoped: a missing scope gives
  `403 Scope(<name>) missing`; optional IP restriction.
- Endpoints used and limits: `GET /equity/account/summary` 1/5 s; `GET /equity/positions` 1/1 s;
  `GET /equity/history/transactions` (`cursor`, `time`, `limit` <= 50) 20/min; `GET /equity/history/dividends`
  (`cursor`, `ticker`, `limit`) 20/min. Limits count per account whatever the key or IP; bursts allowed; every
  response carries `x-ratelimit-limit/-period/-remaining/-reset/-used`. `equity/account/cash` and
  `equity/portfolio`, used by the old sheet, are no longer documented.
- History types: WITHDRAW, DEPOSIT, FEE, TRANSFER, INTEREST_ON_FREE_CASH, LENDING_INTEREST. No card type and no
  merchant field; the CSV report (`POST /equity/history/exports`, 1/30 s) is the only other history source.

## Phase 6: CSV import provider (live once deployed)

As built (see `bank-sync/README.md`, "CSV import"):

- Owner decisions (2026-10-25): files come from a **Drive folder** per holder (`CSV_FOLDER_<HOLDER>` = folder id,
  connection `csv-<holder>`), read by every *Sync all* (so also the daily sync) and on demand by *Import CSV
  files*. Compared before coding: pasting into an import tab (no scope, but Sheets re-parses decimal commas and
  dates by the sheet locale) and an upload dialog (no scope, one manual step per file) were the alternatives; the
  folder needs `drive.readonly` and no manual step after the upload. **Avanza waits for B1**, so its movements go
  straight to `investment_transactions`; it becomes one more entry in `csv-formats`.
- One registry entry (`csv`): connections from the folder properties, no consents, Show status = folder
  `set`/`MISSING` and the last read (`CSV_IMPORTS_<HOLDER>`: time, per source files, transactions and date range,
  skipped files with the reason; no amounts or text). The Enable Banking and Trading 212 providers are untouched;
  `SyncAll`, `TxnSync`, `TxnMerge`, `RawSync` and the FX step are reused as they are (the provider returns all
  of the folder's rows in one window, `maxWindowDays: Infinity`).
- Each run reads every `.csv` in the folder (Google-format files are ignored), finds its format by the header
  (trying UTF-8, then Windows-1252 if UTF-8 decodes with invalid bytes), and normalizes it. A file that matches no
  format or has an unreadable row is skipped whole and named; the others still import. Nothing is guessed.
- Ids without a bank id: `h<date>-<fnv1a of the parsed identifying columns>-<occurrence>`, the counter per
  identical row within a file. Exports cover whole days, so the same or an overlapping file yields the same ids:
  0 added, 0 updated. Rows in two overlapping files count once. The date in the id confines a 32-bit hash
  collision to one day.
- `accountId` = `<source>-` + hash of the account number; `ibanMasked` = its last four. IBANs and 12-19 digit runs
  (card, account and personal id numbers) are masked to their last four characters in `description`,
  `counterparty` and the `api.*` columns.
- BIL export (from the owner's sample, 2026-10-25): `;`-separated; title lines, then `<IBAN> <product>`, a blank
  line, then the header `Date;Value date;Name;Message;Amount`; dates `DD/MM/YYYY`; amounts signed (out negative),
  `.` decimal mark and `,` thousands separator (`8,828.03`, `-1,500`); some messages are in double quotes, some
  not (both are read); no id column (identifying columns: both dates, `Name`, `Message`, `Amount`); no currency
  column (taken as EUR). Opening the file in a spreadsheet app as comma-separated splits amounts with a thousands
  comma across two columns: an artefact of the viewer, not of the file. `Name` is the movement type (`Payment
  order`, `Credit`, `Card payment`); the counterparty comes from `Beneficiary:` / `By order of:` in the message,
  or the merchant text before the date for card payments.
- New OAuth scope `drive.readonly` (in `bank-sync/src/appsscript.json` and `scripts/allowed-scopes.json`): DriveApp
  needs a Drive scope to read files the script did not create (`drive.file` does not cover files uploaded by hand).
  Read-only: files are never changed, moved or deleted. The web-app path is unchanged.
- Deploy steps: `clasp push --force` from `bank-sync/` after diffing the manifest (time zone and the new scope),
  reload the sheet and re-authorize (the daily trigger fails until then), create the folder, set
  `CSV_FOLDER_NIKLAS`, upload a BIL export, run *Import CSV files* twice (the second reports 0 added, 0 updated)
  and check *Show status*. No `clasp deploy`.
- Not yet seen live: a real BIL file through the import, its encoding, and whether BIL changes a movement's text
  between exports (that would add a second row).

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
3. Do the next phase in one small PR, opened as a draft, with this file updated in the same PR (as the
   state after merge and deploy); mark it ready when CI is green, read the review comment and its
   `code-review` issues, fix the real ones on the branch, and stop after the phase. Do not re-run the
   review yourself.
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

- **B1. Investment-account transactions in their own tab** (planned; next before Avanza). Move Trading 212
  (and then Avanza) cash movements and dividends out of `transactions` into `investment_transactions`, so
  `transactions` stays current-account cash flow. What it touches:
  - the transactions store becomes a property of the registry entry (`SyncAllDeps.store` is one store today;
    `TxnSync` gets the entry's store, and the FX step converts both tabs);
  - labels: unchanged (`AccountLabels.stamp` works on any row), same columns on both tabs;
  - FX: `FxService.convertTransactions` runs once per tab, rates fetched once for the run;
  - migration: a one-time, repeat-safe step that moves the existing `trading212` rows (by provider) from
    `transactions` to the new tab, keeping their EUR columns, so nothing is re-fetched or re-converted;
  - the dedupe key stays provider + accountId + transactionId, so the T212 history backfill and later syncs land
    on the moved rows unchanged;
  - card spending from Trading 212 (B3) is current-account-like spending and would stay in `transactions`.
- **B2. Internal transfers** (open decision). A bank-to-Trading 212 transfer appears twice in `transactions` (the
  bank's outflow, e.g. a BIL standing order to "Trading 212 Markets Limited", and the Trading 212 `DEPOSIT`); the
  same will hold for Avanza and for transfers between the owner's own accounts. Options: (a) match pairs (same
  holder, opposite amounts, a few days apart) and mark both `internal` so analysis skips them; (b) let B1 settle
  the investment side (the deposit leaves `transactions`; the bank outflow stays as money moved to savings). B1
  covers investment accounts only; transfers between bank accounts would still need (a).
- **B3. Card spending** (open decision). The Trading 212 API returns card payments only as plain `WITHDRAW` rows
  and cashback as `DEPOSIT` rows, without merchant or type. The Trading 212 CSV report (`POST
  /equity/history/exports`, or exported from the app) has them. Header from the owner's export (2026-10-25):
  `Action, Time (UTC), Notes, ID, Total, Currency (Total), Currency conversion fee, Currency (Currency conversion
  fee), Merchant name, Merchant category`; actions seen: `Card debit` (negative `Total`, merchant name and
  category), `Spending cashback`, `Deposit` (note `Bank Transfer`); time `YYYY-MM-DD HH:MM:SS+00:00` in UTC;
  `.` decimal mark; one id per row (a UUID). This header has no instrument columns (ISIN, ticker, shares),
  presumably because that export held no orders or dividends, so a format must match on the columns it needs,
  not the whole header. The raw separator is not confirmed yet (the sample was copied from a spreadsheet app). Because the API already has these rows,
  importing them as new rows would count each card payment twice. Open: whether the API `reference` equals the
  report's `ID` (then the report can add merchant and category to the API rows instead of adding rows), and
  where the enrichment lives (columns on `transactions`, or a small `card_details` tab joined on the id).
- **Daily values for banks:** tag the booked balance `cash` in the Enable Banking normalizer.
- **Old portfolio sheet:** it calls Trading 212 with the legacy key and shares the per-account rate limit; retire
  it or move what is still used.

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
