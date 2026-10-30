# Bank Sync: roadmap and working state

Single source of truth for the `bank-sync/` project, so work can resume in a fresh
session. Read `CLAUDE.md` first (binding), then `docs/HANDOFF.md` (architecture rules:
SRP/DRY, JSDoc, provider abstraction, normalized models, FX design, idempotent
imports, Node-testable pure logic). This file records what has been built, what was
decided, and what comes next. Update it in the same PR that changes the state.

Last updated: 2026-10-30 (account numbers stored in full by default, `MASK_ACCOUNT_NUMBERS` and *Re-process stored
rows*; Trading 212 report export live (#83, #88, #94); 6b Avanza CSV import live;
Avanza deposits and withdrawals stay on `investment_transactions`; B2 moved into a later classification engine;
remaining work in order).

## Setup in one paragraph

Two independent Apps Script projects, each bound to its own Sheet:

- **Legacy** (`src/`, `.clasp.json` at repo root): sheet "Financial Planning". Plaid and
  SaltEdge code. **Not being developed further**; the smoke test of it was dropped
  because its `Transactions` sheet is empty. Retirement is a pending decision (below).
- **bank-sync** (`bank-sync/`, own `.clasp.json`, namespace `BankSync`): sheet "Bank Sync".
  Enable Banking (banks), the Trading 212 public API (one API key per account) and CSV exports in a
  Drive folder per holder (BIL, Avanza; the Trading 212 report for card details) are the data sources.
  Transactions are kept on three tabs (B1): `transactions` (cash flow and card spending),
  `investment_transactions` and `unclassified_transactions`. Europe/Brussels time zone (the same clock as Luxembourg), EUR reporting currency.

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
| 6 | CSV import provider: Drive folder per holder, BIL export, Trading 212 card details (B3) | Merged (#68), live since 2026-10-25 (first import: 159 BIL rows; report details on 8 Trading 212 rows; a second run changed nothing). BIL message fields and per-bank format files: #72, pushed 2026-10-26 |
| B1 | Investment transactions: `investment_transactions` and `unclassified_transactions` tabs, rows routed by provider and report label | Merged (#76), live since 2026-10-26 (see Live state). Trading 212 bank transfers to `transactions`: #79 |
| 6b | Avanza CSV export as one more CSV format, rows to `investment_transactions` | Merged (#80), live since 2026-10-30 (see Live state) |
| 9 | Avanza unofficial API (values, positions, cash) | Roadmap item only, not planned: see "Avanza" for the findings and what it would take |
| 7 | Danske Bank (Sweden) | Done with no code change: connected through *Connect a bank...*; more banks follow the same route |
| 8 | Wife's accounts (second holder at the same bank) | Last; already accounted for in the 2.5 design |
| M1 | Account numbers in full: `MASK_ACCOUNT_NUMBERS` (default off), one masking helper, *Re-process stored rows* | Merged; live once deployed and re-processed (see "Account numbers") |
| B3+ | Trading 212 report export: monthly request through the API, saved to an exports folder, applied by the CSV import | Merged (#83, #88, #94), live since 2026-10-30 (see Live state) |

Rules for every phase: branch per task, one small PR per phase, squash merge, never
commit to `main`, conventional-commit PR title with a lowercase subject (e.g.
`feat(bank-sync): add transactions writer (phase 2)`), summary and test plan in the PR
body, open as draft and mark ready when CI is green (that starts the review), ask before merging, and stop after each phase for review.
**Update this file in the same PR as the phase**, written as the state after merge and deploy (status
row, live state, deploy steps), so no follow-up docs PR is needed. Anything only seen live (counts,
quirks) goes into the next phase's PR.

## Remaining work (owner, 2026-10-30)

In this order; each item still gets its own PR and stops for review.

1. Owner action: drop an Avanza "Sedan start" export in the CSV folder. (The Trading 212 report export is live and named
   795 of the 817 unclassified rows; see Live state. The 22 left are still to be looked at.)
2. Bank balances in `daily_values` (tag the booked balance `cash`; BIL has no balance in its export) and Avanza cash from
   a "Sedan start" export (open decision "Daily values for banks").
3. Classification engine: categories for all spending, with B2 (internal-transfer pairs) as part of it. Design first.
4. Analysis on the classified data, toward one sheet (open decision "One sheet").
5. When convenient: Phase 8 (partner), B4 (Drive scope), retiring Plaid/SaltEdge, the old portfolio sheet and the
   one-time migration code. Phase 9 (Avanza API) stays unplanned.

## Account numbers (owner, 2026-10-30)

Owner decision: store IBANs and other account, card and reference numbers in full on the Bank Sync sheet.
`Config.MASK_ACCOUNT_NUMBERS` (default `false`) switches the old masking (`****8707`) back on. As built:

- Every masking site calls `Redact.maskAccount(value)`, which returns the value unchanged when the flag is off and
  calls the old `maskIban` when it is on: `Flatten` (IBAN-shaped values and `iban`/`identification`/... keys, so all
  `api.*` columns, e.g. `creditor_account.other.identification`, `debtor_account_additional_identification[0]
  .identification`), `CsvNormalize.maskNumbers` (IBANs and 12-19 digit runs in CSV texts, `counterpartyAccount`,
  `reference`, `api.*`, report details), the CSV account's `ibanMasked`, and Enable Banking's `ibanMasked`.
  `Redact.text` (logs and error text) is not affected.
- Column names are unchanged (`ibanMasked` stays) so sheets and formulas keep working; its JSDoc says it holds the
  full value when masking is off.
- Ids never used masked text (CSV ids hash the raw columns and the raw account number; Enable Banking ids are
  `entry_reference`, `transaction_id` or a hash of unmasked fields), so the flag changes no id. Tested both ways.
- *Re-process stored rows* (menu) = Sync all over the Enable Banking and CSV connections with `fullHistory`: every
  account's details are read again (`refreshStale` with `force`), each account is re-fetched from its oldest stored
  row (a window the bank refuses with a non-consent error is counted as "not served by the source" and skipped), and
  every CSV file still in the folder is re-imported. Stored `raw_transactions` rows also take their account's fresh
  labels now. Afterwards it lists rows still holding a masked value (`Redact.isMasked`) on the three transaction
  tabs and the raw tabs: "cannot be recovered" (source no longer has them) or "run again" (time budget, failure).
  Tested end to end on a fake sheet: two runs leave every tab's row count and every id unchanged; the second
  updates nothing.
- Deploy: push `bank-sync/` from `main`, reload the sheet, run *Re-process stored rows*, then once more to see
  0 added / 0 updated. Record the counts and the "cannot be recovered" list (counts only) here in the next PR.
  Not known yet: how far back Revolut and Danske serve history through Enable Banking; rows older than that keep
  their masked API fields (their account label is still rewritten).

## Live state (2026-10-30)

- Trading 212 report export (#83, #88, #94) deployed on 2026-10-30, `CSV_FOLDER_T212_NIKLAS` set, monthly trigger on.
  The first full-history request (2016 to today) got HTTP 400 (fixed in #88: a year per report); the next saved three
  reports (2024-08-12 to 2026-10-30) unattended, but the import skipped them (dividend rows have no `ID`; fixed in
  #94). After #94, *Import CSV files*: `csv-t212-niklas` 0 added, 0 updated; "Details from reports: 2559 transactions
  updated (2586 read)"; "Moved between tabs: 791 to transactions, 4 to investment_transactions". A second import: details
  0 updated (2586 read), nothing moved. *Show status*: "trading212-report: 3 files, 2578 transactions, 2024-08-12 to
  2026-10-30", "report export - CSV_FOLDER_T212_NIKLAS: set; last saved 2026-10-30 (2026-07-27 to 2026-10-30)",
  "Monthly Trading 212 report: on", "Rows: transactions 971, investment_transactions 2107, unclassified_transactions 22";
  daily sync of 2026-10-30 06:20 OK; Danske and Revolut 84 consent days left.
- 22 rows stay in `unclassified_transactions`: Trading 212 `WITHDRAW`/`DEPOSIT` rows no saved report names (to look
  at: e.g. rows whose report `ID` differs from the API `reference`, or newer than the last report; counts only so far).

- 6b (#80) deployed on 2026-10-30. First *Sync all*: `csv-niklas` 27 added, 0 updated; "EUR conversion: 27 transactions
  converted, 2 waiting for today's rate"; *Show status* "avanza: 1 file, 27 transactions, 2025-10-27 to 2026-10-22",
  "Rows: transactions 180, investment_transactions 2103, unclassified_transactions 817"; consents: Danske 85 and Revolut
  84 days left. The uploaded export covers one year, not "Sedan start", and has no `Insättning` or `Uttag` rows.
- A second *Sync all* after the next daily run was OK (owner): `csv-niklas` 0 added, 0 updated, nothing moved;
  *Show status* "Last sync: ... (daily): OK".
- Owner decision (2026-10-30): Avanza `Insättning` and `Uttag` stay on `investment_transactions` for now (moving them to
  `transactions`, like Trading 212 bank transfers, was proposed and dropped). A bank-to-Avanza transfer is therefore in
  `transactions` once, as the bank's outflow.

## Live state (2026-10-26)

- B1 (#76) and the time zone (#78) deployed on 2026-10-26 (#72 just before). First run after the push (*Import CSV
  files*): "Moved between tabs: 6 to transactions, 2070 to investment_transactions, 824 to unclassified_transactions",
  "Details from reports: 8 transactions updated (8 read)" (the report filled `classification`; 6 card rows came back to
  `transactions`, 2 bank transfers went on to `investment_transactions`). Second run (*Sync all*): nothing moved,
  0 added/updated, details 0 updated. *Show status*: "Rows: transactions 177, investment_transactions 2070,
  unclassified_transactions 816"; daily sync "between 6:00 and 7:00 (Europe/Brussels)". The 816 unclassified rows are
  Trading 212 withdrawals and deposits older than the one report in the folder (2026-10-19 to 10-24): a report for the
  whole history names them. With #79 the 2 bank transfers move to `transactions` on the next sync.
- Lesson: `clasp push` from the repo root pushes the legacy project, not bank-sync (the first B1 push did that; it was
  harmless because `main` only changed the legacy manifest's time zone). Run it from `bank-sync/`.
- First daily sync (2026-10-26, 06:20 Europe/Luxembourg): *Show status* says `Last sync: 2026-10-26 06:20 (daily): OK`;
  Danske, Revolut and CSV 0 added/updated/removed, Trading 212 2 added (that day's interest), 2 converted to EUR.
  Both consents had 88 days left; Trading 212 history complete. So the trigger, the lock and the last-run record
  work unattended; a follow-up run and a warning email are still not seen.
- `daily_values` holds Trading 212 only, by design: bank balances are not tagged `cash` yet (open decision below).
  Its `cash` is Trading 212's total minus investments, so it moves overnight with exchange rates (non-EUR cash, such
  as the GBP the interest is partly paid in, is valued in EUR by Trading 212) even when no money moved.
- #72 (BIL message fields) was pushed on 2026-10-26 after that daily run, so the `transactions` header gains its three
  columns at the next save that changes something (the next import updates the 159 BIL rows once).

Earlier (2026-10-25):

- Connections: `revolut-niklas` (EUR and SEK accounts on one masked IBAN, consent until
  2027-01-22) and `danske-bank-niklas` (one SEK account, consent until 2027-01-23). The Danske
  account sends a real name ("Danske Konto"); Revolut sends none, so its accounts show
  "Current account EUR" and "Current account SEK" (fallback: type and currency).
- Tabs on the Bank Sync sheet: `raw_accounts`, `raw_balances`, `raw_transactions` (every API
  field as `api.*` columns; `raw_transactions` accumulates, see below), `transactions` (the full deduplicated
  history, signed amounts, label columns, `amountEur`/`fxRate`/`fxRateDate` filled), and `rates` (ECB rates per
  currency and date). `raw_balances` also carries `amountEur`/`fxRate`/`fxRateDate` at the latest rate.
- Menu: *Connect a bank...*, *Sync all*, *Import CSV files*, *Show status*, *Enable daily sync*, *Disable daily sync*
  (since then also *Request Trading 212 report* and *Re-process stored rows*). Sync all is idempotent: a second run
  reports 0 added, 0 updated, 0 removed, and 0 transactions converted to EUR.
- Apps Script: web-app deployment (the redirect `/exec`) is pinned to version 3; the script
  itself is pushed from `main` after #58 (Phases 3 and 4 did not change the web app, so no redeploy). The three
  Phase 4 scopes (`script.scriptapp`, `script.send_mail`, `userinfo.email`) are authorized.
- Daily sync: on (one trigger, 06:00-07:00 Europe/Brussels, owned by the owner's account). On 2026-10-25
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
    against the app on 2026-10-25; this corrects the first reading that they were missing). Phase 6 adds the
    merchant and category from the Trading 212 report (B3).
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
- Time zone: since tzdata 2022b `Europe/Luxembourg` is only an alias of `Europe/Brussels` (identical since 1944), and
  Apps Script stores the alias as Brussels, so a repo saying Luxembourg made `clasp push` ask about the manifest every
  time. All three manifests and `Config.TIMEZONE` say `Europe/Brussels` since 2026-10-26; dates, rates and the daily
  trigger hour are unchanged. If `clasp push` still asks about the manifest, diff it: the time zone is no longer the
  reason. A push left waiting at that prompt pushes nothing (the first Phase 6 test ran the old code until it was answered).
- CSV import (Phase 6): `csv-niklas` is live (folder set 2026-10-25). The first *Import CSV files* added 159 BIL
  rows (converted to EUR) and gave 8 Trading 212 rows their report details; a repeat run reported 0 added, 0
  updated and "0 transactions updated (8 read)". The summary alert kept the run open until it was clicked, which
  is why the summaries now open in a modeless window. BIL descriptions were the whole message; the follow-up
  splits them on BIL's labels (below), so the next import reports the 159 BIL rows as updated once (ids unchanged).

## Current architecture (bank-sync)

Load order is `filePushOrder` in `bank-sync/.clasp.json` (entries include `src/`).

- `core/`: `config` (static, non-secret), `errors`, `redact` (log scrubbing, and `maskAccount`: the one place stored
  account numbers are masked, per `MASK_ACCOUNT_NUMBERS`), `reprocess` (*Re-process stored rows*), `legacy-migration` (one-time, removable), `connections` (bank picker
  helpers and the connection store in Script Properties), `account-labels` (labels joined from
  the accounts list), `raw-sync` (collect and write the three raw tabs), `date-range` (pure
  windows), `txn-merge` (pure upsert), `txn-routing` (B1: the tab rule and the routed store over the three
  transaction tabs), `txn-sync`, `sync-all` (all connections, shared time
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
  Properties), `t212-client` (one method per endpoint, Basic auth), `t212-provider`, `t212-history` (one-time
  backfill), `t212-export` (monthly CSV report: request, check, save to the exports folder).
- `providers/csv/`: `csv-parse` (pure: rows, header below a preamble, numbers and dates in a declared layout),
  `formats/` (one file per export, written from a real sample: `bil`, `trading212-report`, `avanza`; all write to the
  same tabs), `csv-formats` (the shared shapes and the list; `kind` 'transactions' or 'details'; the account from the
  lines above the header, or per row), `csv-normalize` (pure: format detection, accounts, stable ids, details, number
  masking), `csv-folders` (folder per holder and the
  last-read record in Script Properties), `csv-provider`. `lib/gas-adapters` `driveFolder` is the only DriveApp use.
- `providers/enable-banking/`: `eb-client` (one method per endpoint), `eb-auth`,
  `eb-token`, `eb-session-store`, `eb-normalize`, `eb-provider`.
- `sheets/`: `raw-writer` (fixed columns plus one `api.*` column per API field, per-connection
  replace or accumulate by key, never a whole-tab clear; also the `holdings` and `daily_values` columns),
  `transactions-store` (one transaction tab, loads by header name; used once per tab by the routed store).
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

## Phase 5: Trading212 provider (live)

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
- Reports (verified in the OpenAPI file, 2026-10-30): `POST /equity/history/exports` (1/30 s) takes `dataIncluded`
  (`includeDividends`, `includeInterest`, `includeOrders`, `includeTransactions`), `timeFrom`, `timeTo` and returns
  `reportId`; `GET /equity/history/exports` (1/min) lists every report with `status` (`Queued`, `Processing`,
  `Running`, `Canceled`, `Failed`, `Finished`) and `downloadLink`. Not stated: the permission needed (the 403 names
  none), a longest period, the CSV columns, whether the link needs the key and how long it lasts. No endpoint for
  card payments exists.
- History types: WITHDRAW, DEPOSIT, FEE, TRANSFER, INTEREST_ON_FREE_CASH, LENDING_INTEREST. No card type and no
  merchant field; the CSV report (`POST /equity/history/exports`, 1/30 s) is the only other history source.

## Phase 6: CSV import provider (live)

As built (see `bank-sync/README.md`, "CSV import"):

- Owner decisions (2026-10-25): files come from a **Drive folder** per holder (`CSV_FOLDER_<HOLDER>` = folder id,
  connection `csv-<holder>`), read by every *Sync all* (so also the daily sync) and on demand by *Import CSV
  files*. Compared before coding: pasting into an import tab (no scope, but Sheets re-parses decimal commas and
  dates by the sheet locale) and an upload dialog (no scope, one manual step per file) were the alternatives; the
  folder needs `drive.readonly` and no manual step after the upload. **Avanza waited for B1**, so its movements land
  where B1 decided (`investment_transactions`) from the start; it becomes one more entry in `csv-formats` (6b).
- One registry entry (`csv`): connections from the folder properties, no consents, Show status = folder
  `set`/`MISSING` and the last read (`CSV_IMPORTS_<HOLDER>`: time, per source files, transactions and date range,
  skipped files with the reason; no amounts or text). The Enable Banking and Trading 212 providers are untouched;
  `SyncAll`, `TxnSync`, `TxnMerge`, `RawSync` and the FX step are reused (the provider returns all of the folder's
  rows in one window, `maxWindowDays: Infinity`). For details files (B3) `SyncAll` gained one step and `TxnMerge`
  one rule, described under B3.
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
  column (the owner's BIL accounts are EUR, confirmed 2026-10-25). Opening the file in a spreadsheet app as
  comma-separated splits amounts with a thousands comma across two columns: an artefact of the viewer, not of the
  file. `Name` is BIL's movement type (`Payment order`, `Credit`, `Card payment`). The `Message` is
  `<kind> Remittance Info: <text> [Beneficiary: ..] [Beneficiary account: ..] [At: <BIC - bank>] [By order of: ..]
  [Paying bank: ..] [End-to-end Identification: ..] [Payment initiation reference: ..]` (kinds seen: `Standing
  order`, `Instant Payment`, `Credit`); card payments are `<merchant> <country> DD/MM/YYYY HH:MM <card>`. It is
  split on those labels only (a remittance may contain its own "LABEL: value" text, e.g. from the health fund),
  matched from the end in each kind's label order, so even a label name typed inside the remittance stays in it:
  `description` = remittance text, `counterparty` = beneficiary or payer (card: merchant), `category` = the kind as
  a code (card: `CARD_PAYMENT`; no label: the `Name`), `counterpartyAccount` (masked), `counterpartyBank`,
  `reference` (end-to-end id, else payment reference); all parts also as `api.parsed.*` on `raw_transactions`. The
  id uses the raw columns, so changing how the message is split never changes ids.
- `transactions` gained optional last columns `counterpartyAccount`, `counterpartyBank` and `reference` (after
  `category`). `TxnMerge.changed` treats a field a provider does not set as equal to the empty cell read back, so
  providers that do not fill them never count rows as updated.
- New OAuth scope `drive.readonly` (in `bank-sync/src/appsscript.json` and `scripts/allowed-scopes.json`): DriveApp
  needs a Drive scope to read files the script did not create (`drive.file` covers only files the script created
  or the user picked in a Google Picker, not a folder filled by hand). There is no per-folder scope, so it grants
  read access to **all** of Drive; the code opens only the `CSV_FOLDER_<HOLDER>` folders (`Gas.driveFolder` is the
  only DriveApp use) and never writes. Accepted by the owner for the hands-off daily import (review #69); narrowing
  it is open decision B4. Replaced by `drive` on 2026-10-30 so the report export can save into the owner's folder
  (see "Trading 212 report export").
  The web-app path is unchanged.
- Deploy steps: `clasp push --force` from `bank-sync/` after diffing the manifest (time zone and the new scope),
  reload the sheet and re-authorize (the daily trigger fails until then), create the folder, set
  `CSV_FOLDER_NIKLAS`, upload a BIL export, run *Import CSV files* twice (the second reports 0 added, 0 updated)
  and check *Show status*. No `clasp deploy`.
- Trading 212 report (from the owner's export, 2026-10-25): comma-separated, UTF-8, text values in double quotes,
  `.` decimal mark, header `Action,Time (UTC),Notes,ID,Total,Currency (Total),Currency conversion fee,Currency
  (Currency conversion fee),Merchant name,Merchant category`, time `YYYY-MM-DD HH:MM:SS+00:00` (UTC), signed
  `Total`, one `ID` per row. Its `ID` equals the API `reference` (owner checked a card debit). Required columns are
  `Action`, `Time (UTC)`, `ID`, `Total` (columns empty for a period may be missing). See B3 for what it does.
- Not yet seen live: a real BIL file through the import, its encoding, and whether BIL changes a movement's text
  between exports (that would add a second row); a report with orders, dividends, interest or a card refund
  (they are coded from their action names, which were not in the sample).

## B1: investment transactions in their own tabs (live since 2026-10-26)

Owner decision (2026-10-26): `transactions` = current-account cash flow plus card spending. Amended the same day: bank
transfers to and from Trading 212 count as `transactions` too (its cash is used like a current account; both sides of a
transfer stay there for B2 to pair). Every Trading 212 API row
(cash movements, interest, dividends, and the card WITHDRAW/DEPOSIT rows) is kept apart, because the API cannot tell
card spending from a transfer; only once a Trading 212 report (CSV, about once a year) labels a row can card spending
move to `transactions`. Avanza lands on the investment side too. The owner chose three tabs (2026-10-26) so that
`investment_transactions` holds only rows known to be investment movements.

As built (see `bank-sync/README.md`, "Transaction tabs"):

- Tabs, same columns: `transactions`, `investment_transactions`, `unclassified_transactions`. Rule (pure,
  `TxnRouting.tabFor`): `classification = card` or `transfer` goes to `transactions`, `classification = investment` to
  `investment_transactions`; without one, rows of a provider in `Config.ROUTING.INVESTMENT_PROVIDERS` (Trading 212) go to `investment_transactions`, except
  its `unclearTypes` (`WITHDRAW`, `DEPOSIT`), which go to `unclassified_transactions`; everything else to `transactions`.
  Interest, lending interest, dividends, fees and transfers are known from the API alone.
- `classification` is a new optional last column, filled by the Trading 212 report (`CsvT212Report.classify`: a merchant,
  an action starting `CARD_`, or `SPENDING_CASHBACK` is `card`; `DEPOSIT` and `WITHDRAWAL` are `transfer`; every other
  action is `investment`). A re-read without it keeps it (like `counterparty` and `category`); unlike them, the report's
  classification replaces a stored one, because the report is its only source, so a rule change reaches stored rows on
  the next sync. B2 can reuse the column.
- One routed store (`TxnRouting.createStore`) replaces the single store in `Container.syncAllDeps`: `load` reads the
  three tabs as one list, `save` writes each row to its tab. So `TxnSync`, `TxnMerge.merge`/`enrich` and the FX step are
  unchanged and see every row: a re-read updates a moved row in place and never adds it twice (key unchanged), the
  report reaches rows on any tab (also from *Import CSV files*, which runs only the CSV connections), and the EUR
  conversion runs once over all three tabs, with rates fetched once per run.
- Writes: only tabs whose content changed (compared cell for cell with what the tab held) are written. A row that changes
  tab is first written to its new tab while the old one keeps it, and removed from the old tab after; a save that stops in
  between leaves it on two tabs, and the next load joins the copies (the first copy, with empty fields filled from the
  second). A tab without rows is not created.
- Migration: no separate step. The first save after the push applies the rule to the rows on `transactions`, keeping
  their EUR columns and details; the 8 rows the report already labelled get their `classification` from the report in the same run
  (the folder is read every run), so they end on `transactions`. Repeat-safe: a second run moves nothing.
- Summary: "Moved between tabs: N to transactions, M to investment_transactions, K to unclassified_transactions." (only
  when something moved; also in the last-run record). *Show status*: "Rows: transactions N, investment_transactions M,
  unclassified_transactions K".
- No new OAuth scope; the web app is unchanged (no `clasp deploy`). Deploy: `clasp push --force` from `bank-sync/` after
  diffing the manifest, reload the sheet, *Sync all*, check the summary and *Show status*, then *Sync all* again (nothing
  moves). Expected on the first run: every Trading 212 row leaves `transactions`, and the 8 labelled card rows return.
- Trade-off: card payments stay unclassified until a report covering them is in the folder. Since the report export
  (see "Trading 212 report export") the API builds that report every month, so card rows are named without a manual
  export.

## Phase 6b: Avanza CSV import (live since 2026-10-30)

As built (see `bank-sync/README.md`, "CSV import"), with the owner's decisions of 2026-10-26:

- `formats/avanza.js` plus its entry in `csv-formats`, detected by header. Required columns are the ones every
  version has (`Datum`, `Konto`, `Typ av transaktion`, `Värdepapper/beskrivning`, `Antal`, `Kurs`, `Belopp`,
  `Transaktionsvaluta`, `ISIN`); `Courtage` (`Courtage (SEK)` in an older export), `Valutakurs`, `Instrumentvaluta`
  and `Resultat` are optional. UTF-8 only.
- Several accounts per file: a format without a file-level `account` (BIL keeps its account line above the header)
  names each row's account (Avanza: `Konto`), and `CsvNormalize.file` returns `accounts` (one per account seen).
  `accountId` = source + hash of the number, or of the name when there is none (`avanza-<fnv1a(Konto)>`);
  `ibanMasked` empty; the account's currency is the one most of its rows have (rows keep their own). The identical-row
  counter is per account (for BIL, one account per file, so its ids are unchanged).
- Rename: accepted and documented (owner). A renamed `Konto` gets a new `accountId`; the next import adds its rows again
  and the old ones stay; the README gives the fix (rename back, or delete the old rows and exports). A name mapping
  (new name to old) can be added later without migration if it ever happens.
- Routing: `Config.ROUTING.INVESTMENT_CSV_SOURCES` (`['avanza']`); `TxnRouting.tabFor` sends a `csv` row whose
  `accountId` prefix is listed to `investment_transactions`, after the report label and before the provider rule. BIL
  stays on `transactions`; no data migration. Not chosen: the format setting `classification = investment`, since that
  column is the report's label and B2 may write `internal` there. Deposits and withdrawals go there too (owner,
  2026-10-30: kept there for now; moving them to `transactions` would be a list of Avanza transfer types, no migration).
- Mapping, the same for every type (so unknown types are imported, with the type as written): amount = `Belopp` as
  signed, currency = `Transaktionsvaluta`, description = the type, category = the type as a code (`SALJ`,
  `ATERINV_UTDELNING`, ...; `CsvParse.code` now drops accents, which changes nothing for the ASCII BIL and Trading 212
  labels), counterparty = `Värdepapper/beskrivning`, reference = `ISIN`, no value date. `Antal`, `Kurs`, `Courtage`,
  `Valutakurs`, `Instrumentvaluta`, `Resultat` only as `api.*` on `raw_transactions` (kept 365 days; older ones are in
  the export). A blank `Belopp` (not seen) is amount 0. An unreadable date, number, account or currency skips the file.
- Trades (`Köp`/`Sälj`) are imported (owner): unlike Trading 212 (whose orders show in `daily_values`), they are
  Avanza's only record, and with them an account's rows sum to its cash movement (the basis for Avanza cash in
  `daily_values` later). `category` tells them apart.
- Ids: `h<date>-<fnv1a of Datum, Konto, type, Värdepapper/beskrivning, Antal, Kurs, Belopp, Transaktionsvaluta, ISIN>-<n>`.
  Fee, exchange rate, instrument currency and result are left out, in case Avanza renames or recomputes them. The
  reinvested dividend pair (same date and ISIN) differs in type, quantity, price and sign, so it never collides; a
  "Sedan start" export and a later month give the same ids for the days they share (0 added, 0 updated).
- `CsvParse.number` now requires each part to be digits, so a `.` in a decimal-comma file is an error instead of a
  decimal point (BIL's accepted inputs are unchanged).
- EUR conversion, Show status (`avanza: N files, M transactions, <from> to <to>`) and the import summary (added and
  updated per CSV connection) are unchanged code.
- No new OAuth scope; the web app is unchanged (no `clasp deploy`). Deployed 2026-10-30 (see Live state) with a one-year
  export. When the owner adds a "Sedan start" export: *Import CSV files* adds only the older rows (the shared days give
  the same ids), all on `investment_transactions`, nothing moved; *Show status* shows the wider date range; a second run
  adds and updates nothing.
- Not yet seen live: a `Köp`, `Insättning`, `Uttag`, `Ränta`, `Utländsk källskatt` or `Övrigt` row (the sample had
  sales, dividends, a reinvested dividend, yield tax and risk premium only), a foreign-currency trade, and a row
  without `Belopp`.

## Avanza

Compared on 2026-10-26 (code and docs only; no login, no Avanza call). Avanza is not covered by Enable Banking.

- **CSV export (chosen; built in 6b, see above):** website only (not the app), *Min ekonomi > Transaktioner*, any period including
  "Sedan start", *Exportera transaktioner*. Checked on the owner's export (structure only): UTF-8 with BOM, `;`, LF, header
  `Datum;Konto;Typ av transaktion;Värdepapper/beskrivning;Antal;Kurs;Belopp;Transaktionsvaluta;Courtage;Valutakurs;Instrumentvaluta;ISIN;Resultat`
  (an older published sample has `Courtage (SEK)`, so only the stable columns should be required); dates `YYYY-MM-DD`;
  decimal comma, no thousands separator; `Belopp` signed from the cash side (sale +, fees and taxes -, reinvested dividend
  -), `Antal` negative on a sale. `Konto` is the account **name** (no account number since Feb 2025), so accounts are told
  apart by name, and renaming one in Avanza changes its `accountId` (accepted, see 6b). No id column: ids as for BIL
  (date + hash + counter). Types seen: Sälj, Utdelning, Återinv. utdelning, Avkastningsskatt, Uttag av riskkostnad (also expected: Köp,
  Insättning, Uttag, Ränta, Utländsk källskatt, Övrigt). No positions or value export was found, and exports cannot be
  scheduled or emailed, so the CSV gives transactions, not `daily_values`.
- **Unofficial API (Phase 9, roadmap item, not planned):** the maintained client is `Qluxzz/avanza` (Python, commits to
  2026-09); `fhqvst/avanza` (Node) has not changed since 2021. Endpoints: `/_api/account-overview/overview/categorizedAccounts`
  (accounts, total value, buying power), `/_api/position-data/positions` (positions), `/_api/transactions/list`
  (transactions with ids). Login: username and password, then a TOTP code (`/_api/authentication/sessions/usercredentials`,
  `.../totp`), giving an `X-SecurityToken` header and cookies; sessions last 30-1440 minutes, so a daily unattended login is
  possible; no BankID. Feasible in Apps Script (cookies by hand, TOTP with `Utilities.computeHmacSignature`), but Avanza's
  terms forbid "robotar, spindlar, scrapers och andra automatiska verktyg" without written consent, the paths change several
  times a year, Cloudflare may block Google's servers (untested), and the three secrets (username, password, TOTP secret)
  are a full login with trading rights; there is no read-only key. It would give `Holding` and `Balance` rows with
  `valueKind`, so `daily_values` works unchanged. Only if the owner accepts the terms and the credential risk; then first a
  one-off login test from Apps Script.
- **Aggregators:** GoCardless (Nordigen) closed to new sign-ups in 2025; Tink and Insurely list Avanza but sell to
  businesses only; PSD2 covers payment accounts, not ISK or depot positions.

## Trading 212 report export (live since 2026-10-30)

Owner request (2026-10-30): request the Trading 212 report through the API every month instead of exporting it by
hand, with a manual option, and save it to the owner's own Drive folder ("Trading212 exports"). As built (see
`bank-sync/README.md`, "Trading 212 report export"):

- `t212-export`: `request` asks for a report per key with its secret whose exports folder `CSV_FOLDER_T212_<HOLDER>` is
  set (dividends, interest and transactions; not orders), `fetch` checks the reports list and saves a finished report as
  `trading212-report-<holder>-<from>-to-<to>.csv`. State per key in `T212_EXPORT_<HOLDER>` (report id, period, checks,
  last saved file, last error; never the download link). The client gained `requestReport`, `reports` and `download`;
  `lib/http` a `text` option for the CSV body and `followRedirects: false` (a redirect comes back as its target). The
  key goes only to trading212.com hosts (`T212Client.isOwnHost`), checked at every redirect hop (followed by hand,
  at most 3); a storage link is fetched without it (review #86).
- The folder property starts with `CSV_FOLDER_`, so the exports folder is also a CSV connection (`csv-t212-<holder>`):
  Sync all, the daily sync and *Import CSV files* apply the saved reports with no change to the CSV import, and report
  rows whose API rows arrive later are applied later, as for a report uploaded by hand.
- Period: at most `TRADING212.EXPORT_MAX_DAYS` (365) per report. Seen live on 2026-10-30: a request from 2016-01-01 to
  the request day got HTTP 400 (the error said only "request failed", because the body has no `message`; `lib/http`
  now shows the body instead), matching the app export's one-year cap. So the first report starts at the connection's
  oldest stored row (`T212Export.oldestDate`; `EXPORT_FIRST_FROM` only when there is none; it reaches back as far as the
  history backfill has, which is complete live) and covers a year; when a saved report ends before today, the same
  check requests the next year at once, until the history reaches today. A failed follow-up request is retried with
  each check (`catchUp` in the state counts the failures, at most `REPORT_MAX_CHECKS`), then *Show status* says the
  catch-up stopped and a manual request continues it (review #89). Every report starts 7 days
  before the last saved end, so a missed month catches up. `timeTo` is the end of the period's last day, or the
  request time to the second when it ends today.
- Triggers: *Enable daily sync* also creates a monthly trigger (`triggerT212ReportRequest`, day 1, 05:00-06:00);
  the request schedules a one-off `triggerT212ReportFetch` 5 minutes later; an unfinished report is checked every 5
  minutes, at most 6 checks, then stays pending until the next request (a new request replaces it). *Disable daily
  sync* removes both. Menu *Request Trading 212 report* requests now. Requests and checks run under the script lock
  Sync all takes, so they never overwrite each other's state (review #85); a busy lock fails a request (the menu says
  so, a trigger's failure email too) and reschedules a check without counting it. After every request a check stays
  scheduled while any report is pending with checks left, also when that request failed (review #84).
- *Show status*: "Monthly Trading 212 report: on/off" and per key "report export - CSV_FOLDER_T212_<HOLDER>:
  set/MISSING; last saved <date> (<from> to <to>); requested ...; last error: ...".
- OAuth: `drive.readonly` replaced by `drive` (manifest and allow-list): `Folder.createFile` needs `drive`, and
  `drive.file` cannot add a file to a folder the script did not create. The owner chose that over a script-created
  folder (B4). The code still opens only the configured folders and only adds files.
- Deploy: check that the Trading 212 key has the general **History** permission (the docs name none for exports);
  `clasp push --force` from `bank-sync/` after diffing the manifest (the scope change is the expected difference);
  reload the sheet and re-authorize (the daily trigger fails until then); set `CSV_FOLDER_T212_NIKLAS` to the exports
  folder id; *Enable daily sync* again (adds the monthly trigger; *Show status* shows it on); *Request Trading 212
  report*; after about 5 minutes *Show status* shows "last saved"; then *Import CSV files*. Expected: "Details from
  reports: N transactions updated", rows moved from `unclassified_transactions` to `transactions` (card spending,
  cashback, bank transfers) and `investment_transactions`, and `unclassified_transactions` close to 0; a second import
  updates and moves nothing. If the report is skipped ("its header matches no known export"), the API's report has
  other columns than the app's: send the header line (no values) for a format update. No `clasp deploy` (the web app
  is unchanged).
- Live (2026-10-30, after #88): *Request Trading 212 report* asked for 2024-08-12 (the oldest stored row) to
  2025-08-11; the chain saved three reports (2024-08-12 to 2025-08-11, 2025-08-04 to 2026-08-03, 2026-07-27 to
  2026-10-30) without help, and Trading 212's app notified "ready to download" for each. The API's report has the
  app export's columns plus `ISIN`, `Ticker`, `Name`, `No. of shares`, `Price / share`, `Exchange rate`, `Withholding
  tax`, `ATM Withdrawal Fee` and their currencies, so the format matched; but *Import CSV files* skipped all three
  files ("A Trading 212 report row has no ID"): **dividend rows (`Dividend (Dividend)`, `Dividend (Dividend
  manufactured payment)`) have no `ID`**. The format now skips those rows (dividends come from the API's dividend
  feed and would match nothing); any other row without an ID still skips the file. Actions seen: Deposit, Withdrawal,
  Interest on cash, Lending interest, Dividend (two kinds, no ID), Dividend adjustment, Card debit, Card credit, Card
  ATM Withdrawal, Spending cashback, New card cost (now `card`: the card's fee). Structure only; no values recorded.
- Not yet seen live: the download link's host, how long a report takes to build, and the 400's body (whether a year
  is the exact limit).

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

- **B1. Where Trading 212 and Avanza rows live: decided (owner, 2026-10-26), built (#76; Avanza in 6b).** See "B1" below:
  option (c), with a third tab for the rows the API cannot tell apart until a report names them; bank transfers to and
  from Trading 212 count as `transactions` (owner, 2026-10-26: its cash is used like a current account).
- **B2. Internal transfers** (open decision; owner, 2026-10-30: part of the later classification engine, not built on its
  own). The owner treats Trading 212's cash like a current account (2026-10-26), so a bank-to-Trading 212 transfer has
  both sides in `transactions`: the bank's outflow and the Trading 212 deposit (once a report names it `transfer`; until
  then in `unclassified_transactions`). The same holds for transfers between the owner's own bank accounts (Revolut,
  Danske, BIL). Avanza `Insättning`/`Uttag` stay on `investment_transactions` (owner, 2026-10-30), so a bank-to-Avanza
  transfer is in `transactions` once (the bank's outflow, a saving). Until B2, transfers inflate income and spending
  totals. Design notes from 2026-10-30, input for the engine:
  - pairs: bank and Trading 212 `transfer`; bank and Avanza `INSATTNING`/`UTTAG` (across tabs: the routed store loads all
    three); bank and bank only with a self-transfer hint (the other account's last four in `counterpartyAccount`, which
    Enable Banking rows would need mapped from `creditor_account`/`debtor_account`, or the holder's own name as
    counterparty), since amount and date alone would pair a salary with a rent payment. Same holder; the connection
    holder of a second Trading 212 key (`niklas-isa`) needs a mapping to the person;
  - matching: same currency on exact opposite `amount`; across currencies on `amountEur` within a tolerance (e.g. the
    larger of 2 EUR and 2.5%); dates a few days apart; a pair only when each side has exactly one free candidate, else
    both stay unpaired and are counted as ambiguous; existing pairs are kept and only undone when a side disappears or no
    longer matches, so a re-sync never re-pairs differently;
  - the mark cannot be `classification` (routing input, replaced by the Trading 212 report on every sync): a new optional
    column (e.g. `transferPair`, a stable id from the two row keys, kept by `TxnMerge.combine`, never set by `enrich`,
    ignored by routing), with a manual `none` to exclude a row; analysis skips internal pairs and shows bank-to-Avanza
    pairs as saving. Runs after the FX step; counts only in the summary and *Show status*.
- **B3. Card spending** (done in Phase 6). The Trading 212 API returns card payments only as plain `WITHDRAW` rows
  and cashback as `DEPOSIT` rows, without merchant or type. The Trading 212 report (exported from the app, or `POST
  /equity/history/exports`) names them, and its `ID` equals the API `reference`, so it **adds details to the API
  rows instead of adding rows** (adding would count every card payment twice):
  - format `trading212-report` (`kind: 'details'`): every report row gives its API row a `category` (the new
    column): the merchant category when there is one (e.g. `SERVICE_PROVIDERS`), else the action as a code
    (`CARD_DEBIT`, `SPENDING_CASHBACK`, `DEPOSIT`, ...); and a `counterparty`: the merchant name, else the notes
    (e.g. `Bank Transfer`). So a card payment is told apart from a withdrawal and cashback from a deposit. Report
    rows the API has not delivered yet are applied on a later sync (the folder is re-read every run).
  - `transactions` gained an optional last column `category` (older tabs load; the table is one column wider).
  - `SyncAll` collects `getEnrichments()` from providers that have it and, after every connection and before the
    FX step, applies them with the pure `TxnMerge.enrich` (provider + transactionId; it only fills empty fields, so a
    value the API gave, e.g. a dividend's instrument, is never replaced; saves only if something changed; a failure is reported and marks the run failed, never stops it). The summary says
    "Details from reports: N transactions updated (M read)."
  - `TxnMerge.combine` keeps a stored `counterparty`/`category` when a re-fetched row has none, so the Trading 212
    sync's 7-day re-read never wipes the merchant (and a merchant set by the report never counts as an update).
  - The Trading 212 provider itself is unchanged. Not used: the conversion fee columns (no column for them). Report
    rows for orders match nothing (orders are not imported); dividend rows match nothing either while our dividend
    ids are `div:<reference>` (no dividend row was in the sample, so whether the report's dividend `ID` is that
    reference is not checked).
- **B4. Narrow the Drive scope** (open decision, raised in review #69). The CSV import reads a Drive folder with
  `drive.readonly`, which can read all of Drive (no per-folder scope exists). Kept for now (owner, 2026-10-25) for
  the hands-off daily import. Widened to `drive` (owner, 2026-10-30) so the Trading 212 report export can save into
  the owner's own folder; `drive.file` would have meant a folder created by the script instead. Options to reconsider:
  - **Upload dialog, no Drive scope**: *Import CSV files...* opens a file chooser; the browser sends the file text;
    each upload runs a normal *Sync all* so report details find that day's API rows. Same manual effort as
    dropping a file in the folder; files are not kept, so a report row whose API row is missing must be uploaded
    again, and the daily sync no longer imports CSVs. Change: a dialog plus one `google.script.run` global, the
    provider fed from uploads instead of `Gas.driveFolder`, the scope removed from both manifests.
  - **`drive.file` with a Google Picker**: only picked files are readable. Needs a Google Cloud project, the Picker
    API and an API key; whether picking a folder exposes its files is not documented (to verify), so likely one
    pick per file.
  - **Keep a Drive scope** (current: `drive`; back to `drive.readonly` if the export saves into a folder the script
    created, with `drive.file`).
- **Daily values for banks:** tag the booked balance `cash` in the Enable Banking normalizer. BIL has no balance in
  its export, so it cannot appear in `daily_values` (total wealth) without another source: check whether BILnet
  offers a balance or statement export, or record an opening balance and sum the movements. Avanza: cash could be the
  sum of an account's rows from a "Sedan start" export (trades are imported for this); positions and values would
  still need another source (Phase 9).
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
  The first Phase 2 push set the time zone to Europe/Luxembourg, which Apps Script stores as Europe/Brussels (an alias);
  the repo says Europe/Brussels since 2026-10-26, so the time zone no longer triggers the prompt.
- After a push the user reloads the sheet (menus are built in `onOpen`) and may need to
  re-authorize.

## Secrets and boundaries

- Never read, print or copy `.pem` files, Script Property values, `~/.clasprc.json` or
  tokens. The owner pastes secrets. Never put them in the repo, logs, sheets or chat.
- Bank logins and consent renewals belong to the owner: give the link and wait.
- No personal financial data in logs, commits, PR text or chat; IBANs never appear in logs or error text (`Redact.text`;
  on the sheet they follow `MASK_ACCOUNT_NUMBERS`, owner decision 2026-10-30); tests
  and docs use fake data.
- Ignore "[github assistant] ..." errors in the Apps Script editor (a Chrome extension).
- A `client_secret_*.json` sits (gitignored) in the repo root. The owner chose to leave it
  there for now; do not open it.

## Housekeeping done

Removed the merged remote branches `add-handoff-prompt` and `feat/enable-banking-auth-v2`,
turned on automatic deletion of merged branches, and dropped the untracked SaltEdge
account-filtering note.
