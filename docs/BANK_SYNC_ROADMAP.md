# Bank Sync: roadmap and working state

Single source of truth for the `bank-sync/` project, so work can resume in a fresh
session. Read `CLAUDE.md` first (binding), then `docs/HANDOFF.md` (architecture rules:
SRP/DRY, JSDoc, provider abstraction, normalized models, FX design, idempotent
imports, Node-testable pure logic). This file records what has been built, what was
decided, and what comes next. Update it in the same PR that changes the state.

Last updated: 2026-10-24.

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
| 2 | Transactions: paginated fetch, normalization, idempotent upsert, `transactions` tab | PR #44 open, CI green, live-tested on Revolut (sync twice gives 0 added/updated/removed) |
| 2.5 | Connections, generic menu, bank/account labels (details below) | Next |
| 3 | FX (ECB via Frankfurter) | Planned |
| 4 | Automation and consent-expiry warnings | Planned |
| 5 | Trading212 provider | Planned |
| 6 | CSV import provider (Avanza, BIL) | Planned |
| 7 | Danske Bank (Sweden) via config only | Planned, any time after 2.5 |
| 8 | Wife's accounts (second holder at the same bank) | Last; already accounted for in the 2.5 design |

Rules for every phase: branch per task, one small PR per phase, squash merge, never
commit to `main`, conventional-commit PR title with a lowercase subject (e.g.
`feat(bank-sync): add transactions writer (phase 2)`), summary and test plan in the PR
body, open as draft and mark ready when CI is green (that starts the review), ask before merging, and stop after each phase for review.

## Current architecture (bank-sync)

Load order is `filePushOrder` in `bank-sync/.clasp.json` (entries include `src/`).

- `core/`: `config` (static, non-secret), `errors`, `redact`, `raw-sync` (accounts and
  balances), `date-range` (pure windows), `txn-merge` (pure upsert), `txn-sync`.
- `lib/`: `jwt`, `http` (retry/backoff), `hash` (FNV-1a), `gas-adapters` (the only place
  touching Apps Script services directly).
- `providers/bank-provider.js`: the contract (`authorize`, `listAccounts`, `getBalances`,
  `getTransactions`) and the `Account`, `Balance`, `Transaction` typedefs.
- `providers/enable-banking/`: `eb-client` (one method per endpoint), `eb-auth`,
  `eb-token`, `eb-session-store`, `eb-normalize`, `eb-provider`.
- `sheets/`: `raw-writer`, `transactions-store`.
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

## Phase 2.5: connections, generic menu, labels (next)

Motivation, from live use of Phase 2:

1. The menu is hardwired to Revolut (`BANK = BANKS.REVOLUT`). It must be config-driven.
2. Sessions are stored by bank key, so a second login at the same bank would overwrite
   the first. Introduce a **connection**: key = bank + holder (e.g. `revolut-niklas`), each
   with its own session, consent expiry and accounts. Adding a holder or bank becomes a
   config entry.
3. Tabs do not show which bank, connection or account a row belongs to.

Scope:

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

## Phase 7: Danske Bank (Sweden)

Already linked in the Enable Banking Control Panel. After 2.5 it should be added through
config only (a connection plus a bank entry), with no change to existing providers. Verify
the flow generalizes with SEK data. BGL, Spuerkeess, Raiffeisen and SEB will be linked
later and follow the same route.

## Phase 8: second holder at the same bank (last)

A partner's own accounts at a bank already connected. With the 2.5 connection model this
is a new connection and a separate consent (their own login). Blocked on confirming that
Enable Banking's restricted mode allows it; if not, it needs its own application or
another route. Decide only when this phase starts.

## Open decisions (raise, do not act)

- **One sheet:** move toward a single sheet. Options once Phase 2 is in: migrate the legacy
  views into Bank Sync, or split the repo.
- **Retire Plaid and SaltEdge** from `src/`: propose a removal PR.
- **`deploy.yml`** pushes only the legacy `src/` project. Do not extend it to bank-sync
  without asking.
- **Claude review workflow** (`claude-code-review.yml`): fixed. It runs once, when a PR is
  marked ready for review (PRs start as drafts), runs the checks on the PR head, and files
  findings as `code-review` issues. A workflow PR cannot be reviewed by the workflow it
  changes.

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
