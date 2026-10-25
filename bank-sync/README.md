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
- `amountEur`, `fxRate` and `fxRateDate` stay empty until the FX phase and are
  preserved on re-sync.
- Ids: `entry_reference`, else `transaction_id`, else a content hash.
