# Bank Sync

Modular Apps Script project that pulls bank data via
[Enable Banking](https://enablebanking.com), and cash, holdings and history from
[Trading 212](https://docs.trading212.com/api), and CSV exports from a Drive folder (BIL, and
the Trading 212 report for card details), into a dedicated Google Sheet ("Bank Sync").
Independent of the legacy `src/` project in this repo.

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

The application ID and private key, and the Trading 212 API keys and secrets, live only in Script Properties. Never put
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
between 06:00 and 07:00 (Europe/Brussels, the same time as Luxembourg); *Disable daily sync* removes it. Enabling twice never
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

## Trading 212 (Phase 5)

Each Trading 212 account (yours, your partner's, an ISA) is one connection, read with its own API key
over the [public API](https://docs.trading212.com/api) (live, v0). There is no login flow: the key is
pasted in Script Properties and *Sync all* picks it up.

1. In the Trading 212 app (Invest or Stocks ISA account): **Settings > API (Beta)**, create a key.
   Give it only these permissions: `account`, `portfolio`, `history:transactions`,
   `history:dividends`. Nothing that places or cancels orders. Leave the IP restriction off (Apps
   Script's addresses change). Each person creates their own key while logged in to their own account.
2. In the Apps Script editor, **Project Settings > Script Properties**, add one pair per account,
   with the holder in upper case:
   - `T212_API_KEY_<HOLDER>`: the API key, e.g. `T212_API_KEY_NIKLAS`
   - `T212_API_SECRET_<HOLDER>`: the API secret, e.g. `T212_API_SECRET_NIKLAS`

   A second account of the same person gets its own suffix, e.g. `T212_API_KEY_NIKLAS_ISA`
   (connection `trading212-niklas-isa`).
3. *Show status* lists each pair as `set`/`MISSING` (never the values), and for each complete pair how
   far the history backfill is. *Sync all* (and the daily sync) includes every complete pair. Both the
   key and the secret are needed: with only one, the connection shows as "not connected". Don't
   keep placeholder values for someone who has no key yet; add their pair when they have one.

What it writes:

- `daily_values` (new): one row per account and day with `total` (Trading 212's account value),
  `investments` (its investments value, which already includes the cash in pies) and `cash` (total
  minus investments: available cash plus anything the summary does not break down, e.g. amounts still
  being processed), and the same in EUR. So the total always matches the app. A second run on the same
  day replaces that day's row, so the tab is a daily history; the 06:00 daily sync records the
  previous market close. Rows are built from the balances tagged in `raw_balances.valueKind`
  (`total`, `investments`), not from a separate API call or rate.
- `holdings` (new): the open positions per account (ticker, ISIN, name, quantity, average and current
  price in the instrument currency, value, cost and unrealized result in the account currency,
  `amountEur`), replaced on each sync, with every API field as `api.*` columns.
- `raw_accounts` / `raw_balances`: one account per key (the account number from the API is its
  `accountId`); balances `cash.availableToTrade`, `cash.inPies`, `cash.reservedForOrders` (detail),
  `investments.currentValue` (`valueKind: investments`) and `totalValue` (`valueKind: total`).
- `investment_transactions` (and `raw_transactions`): cash movements (fees, interest, transfers between Trading 212
  accounts; `description` is the type) and dividends (id `div:<reference>`, counterparty the
  instrument). Withdrawals and fees are negative, deposits, interest and dividends positive. Orders
  (buys and sells) are not imported: they move money between cash and investments, which
  `daily_values` already shows. Each run reads from the latest booked date minus 7 days (the first
  run: 90 days back), like the banks, and on top of that **backfills the full history once**: up to
  40 older pages per feed per run, continuing where the last run stopped (kept in the Script Property
  `T212_HISTORY_<HOLDER>`), so an account with a long history completes over a few syncs. Delete
  that property to backfill again. EUR at the booking-date rate. `WITHDRAW` and `DEPOSIT` rows wait in
  `unclassified_transactions` until a Trading 212 report says what they are (card spending or a bank transfer); then
  they move to `transactions` (see "Transaction tabs").

Notes:

- Card payments are in the API only as plain `WITHDRAW` rows (cashback as small `DEPOSIT` rows), with no
  card type, merchant or category, so they cannot be told apart from other withdrawals and deposits. Dropping
  the Trading 212 report in your CSV folder adds the merchant and category to them (see "CSV import") and moves
  the card payments and cashback to `transactions` (see "Transaction tabs"). Bank transfers to and from Trading 212
  move to `transactions` too, because its cash is used like a current account: both sides of such a transfer (the
  bank's outflow and the Trading 212 deposit) are then in `transactions`, until B2 pairs them as internal.
- Rate limits are per Trading 212 account, whatever the key: summary 1 per 5 s, positions 1 per s,
  history 20 per minute. A sync makes about four calls per account. On HTTP 429 the call waits until
  the `x-ratelimit-reset` time (at most 65 s and never past the run's time budget), then retries.
  Other tools using the same account (e.g. the old portfolio sheet) share the limit.
- All amounts come in the account's primary currency; multi-currency accounts are not supported by
  the API. Only Invest and Stocks ISA accounts have API access.
- No new OAuth scope: `script.external_request` already covers the calls.

## CSV import (Phase 6)

Banks that Enable Banking does not cover are imported from their CSV exports. You drop the files in a
Google Drive folder; *Sync all* (and the daily sync) reads every `.csv` file in it, and
*Bank Sync > Import CSV files* does the same on demand (only the folders, no bank or Trading 212 calls).

1. Create a Drive folder per person, e.g. "Bank Sync imports (niklas)". Copy its id from the URL
   (`https://drive.google.com/drive/folders/<id>`).
2. In **Project Settings > Script Properties** add `CSV_FOLDER_<HOLDER>` = the folder id, e.g.
   `CSV_FOLDER_NIKLAS` (connection `csv-niklas`). The id is not a secret, but keep the folder private.
3. Export from the bank and upload the file as it is. If Drive is set to convert uploads to Google
   formats, turn that off (**Drive settings > Convert uploads**): a converted file is a Sheet, not a CSV,
   and is ignored. Don't edit exports in a spreadsheet app and save them again; that can change
   separators and number formats.
4. *Import CSV files*, or wait for the next sync. *Show status* shows per folder when it was last read
   and, per source, how many files and transactions it found and their date range, plus any file it
   skipped and why.

Supported exports:

- **BIL** (BILnet, "List of account movements", one account per file): `;`-separated, a few title lines
  and the account line (IBAN and product) above the header `Date;Value date;Name;Message;Amount`; dates
  `DD/MM/YYYY`; amounts signed (money out negative) with `.` as decimal mark and `,` between thousands
  (`8,828.03`, `-1,500`). Spreadsheet apps that open the file as comma-separated split those amounts
  across two columns; the import reads the raw text, so they are read correctly. UTF-8, else Windows-1252.
  The file has no currency column; BIL accounts are EUR. The `Message` is split on BIL's own labels
  (`Remittance Info`, `Beneficiary`, `Beneficiary account`, `At`, `By order of`, `Paying bank`, `End-to-end
  Identification`, `Payment initiation reference`), never on any colon, and in the order BIL writes them, so
  text such as "DATE CISS: 19.08.2026", or even "At:" typed by the payer, inside a remittance stays in it:
  - `description`: the remittance text (e.g. "Rent for August"); card payments keep the whole message;
  - `counterparty`: the beneficiary, the payer (`By order of`) or the card merchant;
  - `category`: the kind of movement as a code: `STANDING_ORDER`, `INSTANT_PAYMENT`, `CREDIT`, `CARD_PAYMENT`, ...;
  - `counterpartyAccount` (masked), `counterpartyBank` (`At` / `Paying bank`), `reference` (end-to-end id, else
    the payment reference);
  - every labelled part also as `api.parsed.<label>` on `raw_transactions`.
- **Trading 212 report** (exported from the Trading 212 app's history, any period, CSV): adds details to rows
  that *Sync all* already reads from the Trading 212 API, matched on the report's `ID` (the API's reference). It
  never adds rows. Each report row gives its API row a `category`: the merchant category for card payments
  (e.g. `SERVICE_PROVIDERS`), else the action as a code (`CARD_DEBIT`, `SPENDING_CASHBACK`, `DEPOSIT`, ...), so
  a card payment is told apart from a withdrawal and cashback from a deposit; and a `counterparty`: the
  merchant name, else the notes (e.g. `Bank Transfer`). Only empty fields are filled: what the API gave (e.g. a
  dividend's instrument) is kept. Orders are not imported, so their report rows change nothing. A report row
  whose API row is not there yet is applied on a later sync. Each row also gets a `classification`: `card` (a merchant, a
  card action such as `CARD_DEBIT`, or `SPENDING_CASHBACK`), `transfer` (`Deposit`, `Withdrawal`: money moved to or
  from a bank account) or `investment` (everything else); it decides the row's tab (see "Transaction tabs"). Unlike
  the other details, the report's classification replaces a stored one, so a changed rule applies to rows already
  stored. Comma-separated, UTF-8, times in UTC.

What it writes:

- `transactions`: one booked row per movement, `provider` `csv`, `bank` the source (`BIL`),
  `connection` the folder (`csv-niklas`), `accountName` the product from the account line, `ibanMasked`
  its last four characters, and the columns above per source. Account and card numbers inside the texts are
  masked to their last four digits. EUR at the booking-date rate, like every other row.
- `raw_accounts` and `raw_transactions`: the account and every CSV column as `api.*` (masked the same way).
- `transactions` ends with five columns that only some sources fill: `category` (BIL, the Trading 212 report),
  `counterpartyAccount`, `counterpartyBank` and `reference` (BIL), and `classification` (the Trading 212 report). A later
  sync that re-reads a row without a counterparty, category or classification keeps the ones already stored. The summary says how many rows got details from a
  report: "Details from reports: N transactions updated (M read)."
- Each source's parsing lives in its own file (`src/providers/csv/formats/bil.js`, `trading212-report.js`), but
  all of them write to the same tabs, so totals across banks stay one sum.
- The daily sync shows no window (a time-driven run has no screen). *Sync all*, *Import CSV files* and *Show
  status* show their summary in a window that does not keep the run waiting: the run ends as soon as the work
  is done, and the window stays until you close it.
- The `accountId` is the source plus a hash of the account number (`bil-1a2b3c4d`).

Re-importing is safe. Without a bank id, a row's id is its date, a hash of its identifying columns and a
counter for identical rows on the same day, so the same file, or a file overlapping an earlier one, gives
0 added and 0 updated; files can stay in the folder. Rows read from two overlapping files count once. A
row the bank later changes (a different text for the same movement) would come in as a new row; delete
the old one by hand if that ever happens. A file whose header matches no known export, or with a row that
cannot be read (e.g. an amount that is not a number), is skipped as a whole and named in *Show status*;
the other files still import. Nothing is guessed: a new bank needs its export layout added to
`src/providers/csv/csv-formats.js`.

New OAuth scope: `drive.readonly`, to list the folder and read files you uploaded (`drive.file` only covers
files the script created or you picked through a Google file picker, so it cannot see a folder you fill by
hand). Google offers no scope for one folder: `drive.readonly` lets the script read **all** of your Drive. The
code only opens the folders named in `CSV_FOLDER_<HOLDER>` (`lib/gas-adapters.js` `driveFolder` is the one place
that calls DriveApp) and never changes, moves or deletes files. Chosen on purpose (2026-10-25) over pasting
into a tab (no scope, but Sheets re-parses numbers and dates) and an upload dialog (no scope, one manual step
per file), so new files are picked up by the daily sync. Re-authorize after the push (the daily sync fails
until you do).

## Transaction tabs (B1)

Transactions are kept on three tabs with the same columns, so spending and income totals are not mixed with
movements inside investment accounts:

| Tab | What is on it |
|---|---|
| `transactions` | Bank and BIL rows (current-account cash flow), and Trading 212 card payments, cashback and bank transfers (deposits, withdrawals) once a report names them |
| `investment_transactions` | Trading 212 interest, lending interest, dividends, fees and transfers between Trading 212 accounts |
| `unclassified_transactions` | Trading 212 `WITHDRAW` and `DEPOSIT` rows no report has named yet: the API cannot tell a card payment or cashback from a bank transfer |

- The rule is `src/core/txn-routing.js` `tabFor`: a row's `classification` (set by the Trading 212 report: `card` and
  `transfer` go to `transactions`, `investment` to `investment_transactions`) decides; without one, Trading 212 rows go to `investment_transactions`, except `WITHDRAW` and
  `DEPOSIT`, which go to `unclassified_transactions`; every other row goes to `transactions`. The investment
  providers and their unclear types are in `Config.ROUTING`.
- Every sync reads all three tabs as one set and writes each row to its tab, so a row is found wherever it is:
  the API's 7-day re-read updates a moved row in place and never adds it twice (the key stays provider +
  account + transaction id), the report's details reach rows on any tab, and the EUR conversion runs once over
  all three tabs (rates fetched once). Only tabs whose content changed are written.
- A row that changes tab is written to its new tab first and removed from the old one after; if a save stops in
  between, the row is on both tabs until the next sync joins the two copies. A tab with no rows is not created.
- To label unclassified rows, export the Trading 212 report for that period and drop it in your CSV folder, then
  *Import CSV files* (or wait for the daily sync). The summary says what moved, e.g. "Moved between tabs: 3 to
  transactions."; *Show status* shows the rows per tab and the last run's moves.
- Upgrading: the first sync after the push moves the Trading 212 rows already on `transactions` to the other two
  tabs, with their EUR amounts and details. Card payments and bank transfers a report already named come back to
  `transactions` in the same run, because the report in the folder is read again. A second sync moves nothing. No new OAuth scope;
  the web app is unchanged, so no `clasp deploy`.
