# Bank Sync

Modular Apps Script project that pulls bank data via
[Enable Banking](https://enablebanking.com) into a dedicated Google Sheet
("Bank Sync"). Independent of the legacy `src/` project in this repo.

## Layout

```
bank-sync/
  .clasp.json     this project's script (run clasp from this folder)
  src/            pushed to Apps Script (rootDir)
  test/           Node unit tests, never pushed
```

Pure logic (normalisation, FX, dedupe) has no Apps Script dependencies and is
tested locally: `npm test` (Node 20+).

## One-time setup

1. `clasp login` (opens a browser). Enable the Apps Script API at
   https://script.google.com/home/usersettings.
2. From this folder: `clasp create --parentId <SHEET_ID> --title "Bank Sync" --rootDir src`
   to bind a script to the Sheet, then `clasp push`.
3. In the Apps Script editor, open Project Settings, then Script Properties,
   and add:
   - `EB_APP_ID`: your Enable Banking application ID
   - `EB_PRIVATE_KEY`: the full text of the downloaded `.pem` (PKCS#8,
     including the BEGIN/END lines)
   - `EB_REDIRECT_URL`: the redirect URL registered for the application

## Secrets

The application ID and private key live only in Script Properties. Never put
them in the repo, logs, sheets or chat. `.gitignore` excludes `*.pem`, `*.key`,
`.env`, `client_secret*.json` and `.clasprc.json`. Tests and docs use fake data.

## Workflow

Feature branch per module, small PRs, conventional commits, never push to main.

## Connecting Revolut (Phase 1)

1. `clasp push`, then `clasp deploy` once. The web app runs as you and is
   accessible only by you. Its redirect URL is
   `https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec`.
   After later code changes, redeploy to the *same* deployment with
   `clasp deploy -i <DEPLOYMENT_ID>`, otherwise `/exec` keeps serving the old version.
2. In the Enable Banking Control Panel, set the application's redirect URL to
   that `/exec` URL.
3. Add the three Script Properties (see above). `EB_REDIRECT_URL` must equal the URL from step 2.
4. Reload the Sheet. In the **Bank Sync** menu: *Show status*, then *Connect a bank...*:
   type a country code, then part of the bank name (the list comes live from Enable
   Banking), then a short name for whose login it is (e.g. `niklas`). Open the link and
   log in; the redirect lands on the web app, which stores the connection and shows the
   consent expiry. If Enable Banking has not enabled that bank for your application, you
   get a message to link it in the Enable Banking Control Panel first; nothing is saved
   until the bank grants at least one account.
5. *Sync all* refreshes the `raw_accounts` and `raw_balances` tabs and the
   `transactions` tab for every connected login (IBANs are masked to the last four
   characters). A login that is not connected or whose consent lapsed is reported and
   skipped; the others still run.

First run from the editor (e.g. `onOpen`) once so Google can ask you to grant the scopes.

## Connections, labels and raw tabs (Phase 2.5)

A **connection** is one login at one bank, keyed by bank plus holder (for example
`revolut-niklas`), with its own session, consent expiry and accounts. Connections are not in
the code: *Connect a bank...* creates one (they live in Script Properties), so another bank,
or another person at a bank, needs no config change. The only banks you can use are the ones
you have enabled in the Enable Banking Control Panel; Enable Banking has no API that lists
them, so a bank that is not enabled is detected when it is connected (an
`NO_ACCOUNTS_ADDED` / `ACCESS_DENIED` error, or a login that grants no account) and you are
told to link it first.

Every row on `raw_accounts`, `raw_balances` and `transactions` shows `bank`, `connection`,
`holderName` (the holder's name as the bank reports it), `accountName` (the name set by the
user or bank, Enable Banking `details`, else the product name, else the account type and
currency such as "Current account EUR", because banks like Revolut send no name) and `ibanMasked` (last four
characters), so you can group and total by any of them.

`raw_accounts`, `raw_balances` and `raw_transactions` also show **every field the API returns**, one `api.*`
column per field (nested fields use dotted names, e.g. `api.balance_amount.amount`,
`api.postal_address.country`). Values that look like an IBAN are masked. `raw_transactions` accumulates: every sync adds the rows it fetched (the last week or so after
the first run, which reads 90 days) and replaces older copies of the same transaction, newest first, so the raw
API fields are kept; rows older than 365 days (`RAW_RETENTION_DAYS`) are dropped. Rows fetched before this was introduced are not recoverable. Account details are
read automatically the first time and then every 7 days (one bank call per account), so names
and fields stay current without a manual step.

A session created before 2.5 moves to its connection automatically (no re-login). The first
*Sync all* after upgrading replaces the old `raw_accounts` and `raw_balances` contents (rows
from before labels existed are dropped as stale; anything you typed on those tabs is lost, so
keep notes elsewhere). Later runs only replace the rows of connections that synced. The
`transactions` tab is read by header name, so a tab from before 2.5 loads and is rewritten with
the label columns on the next sync.

## Transactions (Phase 2)

*Bank Sync > Sync all* writes the `transactions` tab (one row per
provider + account + transaction id; amounts signed, expenses negative).

- First run fetches the last 90 days; later runs restart each account at its
  latest booked date minus 7 days, so late-booked items and pending to booked
  changes are picked up. Re-running never duplicates rows.
- Fetches go oldest to newest in 31-day windows and the tab is saved after each,
  so a run that reaches the time budget (4.5 min) loses nothing. Run it again to
  continue.
- Pending rows inside a re-fetched window that the bank no longer returns are
  removed (they booked under a new id or were cancelled).
- `amountEur`, `fxRate` and `fxRateDate` are filled by the EUR conversion (below).
- Ids: `entry_reference`, else `transaction_id`, else a content hash.

## EUR conversion (Phase 3)

At the end of every *Sync all*, amounts are converted to EUR with ECB reference rates
from [Frankfurter](https://frankfurter.dev) (`/v2/rates?...&providers=ecb`, no API key).

- `transactions`: each row without an EUR amount gets `amountEur` (rounded to cents, same sign as `amount`),
  `fxRate` (units of the currency per 1 EUR, so `amountEur = amount / fxRate`) and
  `fxRateDate` (the date of the rate used). The rate is that of the booking date (else
  the value date; undated pending rows use the latest rate). Weekends and ECB holidays
  use the previous business day, never more than 7 days back. EUR rows get rate 1.
  Only new rows are converted: a row that already has `amountEur` is never recalculated,
  unless a re-fetch changes its amount, currency or date (then its EUR columns are cleared
  and it is converted again).
- `raw_balances`: the same three columns, at the latest rate.
- Rates are kept on the `rates` tab (date, base, currency, rate, source) and in the script
  cache for 6 hours. A run fetches only what the new rows need and is not stored yet, once
  per currency (one call for the whole range, never per row). A row dated today stays empty
  until today's rate is published (around 16:00 CET) and is converted on a later sync.
- A currency the ECB does not publish (or a failed fetch) leaves those rows' EUR columns
  as they were; *Sync all* names the currency in its summary.

## Daily sync and consent warnings (Phase 4)

*Bank Sync > Enable daily sync* creates one time-driven trigger that runs *Sync all* every day
between 06:00 and 07:00 (Europe/Luxembourg); *Disable daily sync* removes it. Enabling twice never
creates a second trigger (duplicates are removed). The trigger runs as the account that enabled it.

- The triggered run is the same *Sync all* as the menu item, under the same script lock. If it ends at
  the time budget (or the lock is busy because you are syncing by hand), it schedules one follow-up
  run 5 minutes later, which continues from where it stopped; at most 3 follow-ups per day.
- Every run (daily, follow-up or from the menu) records its outcome, counts only, in the
  `BS_LAST_RUN` Script Property. *Show status* shows whether the daily sync is on and the last run
  (OK or FAILED, per-connection counts, EUR conversion, warnings sent). A run that fails as a whole
  is also rethrown, so Apps Script sends its own failure email.
- **Consent warnings**: after each daily run, a connection whose consent lapses in 30, 14, 7 or 1 days
  (`CONSENT.WARN_DAYS`), or has lapsed, gets one email per threshold to the account the script runs as
  (the address is read at run time and never stored). Reaching several thresholds at once sends one
  email. What was sent is kept in `BS_CONSENT_WARNED_<connection>` together with the expiry date, so
  renewing the consent starts afresh. The email names the connection, the date and how to renew
  (*Connect a bank...* with the same country, bank and holder); no balances, amounts, counterparties
  or account numbers. The holder in it is the short label typed when connecting (needed to renew the
  same connection), not the bank's account-holder name, and the email goes only to the script's own
  account. Warnings are only sent by the daily sync, also when that run's sync fails; one connection
  failing (e.g. a mail error) never stops the others' warnings.
- *Show status* marks consents within 30 days with "renew soon" and lapsed ones with "LAPSED".

New OAuth scopes (re-authorize after the push; the first menu click asks):
`script.scriptapp` (create and delete triggers), `script.send_mail` (MailApp) and
`userinfo.email` (`Session.getEffectiveUser().getEmail()`, the warning recipient).
