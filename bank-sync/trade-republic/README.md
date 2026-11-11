# trade-republic

Python job (inside `bank-sync/`) that reads your Trade Republic timeline with the community library
[pytr](https://github.com/pytr-org/pytr) and keeps **one CSV file** in the Bank Sync CSV folder on Drive up to date.
Bank Sync imports it like any other export (format `trade-republic`, `src/providers/csv/formats/trade-republic.js`).
It exists because Trade Republic's official API needs a current account. The in-app *Transaction export* also works
but has to be done by hand.

> **This repo is public.** `.env`, the session cookies, `data/` (events and the CSV) and the service-account key live
> in `~/trade-republic/`, never in the repo. The pre-commit hook blocks CSVs, `.env` and key files.

> pytr uses Trade Republic's **private** API. It is unofficial, not affiliated with Trade Republic, and can break
> when they change it. `pytr` is pinned exactly; upgrading it is a deliberate change with its own PR.

## How it works

1. **Login (web login, v2).** `login.sh` asks for your PIN (never stored), then you confirm the login in the Trade
   Republic app. This is the app.traderepublic.com flow, so your phone stays logged in (no device reset). The session
   cookies are saved in `~/trade-republic/session/`.
2. **Scheduled run (daily 05:40, launchd).** It only resumes the saved session and never asks for anything. When the
   session has expired it exits with code 3 and shows a macOS notification once a day: run `~/trade-republic/login.sh`.
3. **Read.** The first run reads the whole timeline. Later runs read the last `TR_LOOKBACK_DAYS` (45) days, and pytr
   merges them into `data/all_events.json`.
4. **Write.** Every event that moved money becomes one row of `data/trade-republic.csv`, which always holds the whole
   history. The file is then copied over the Drive file `TR_DRIVE_FILE_ID`, and only when its content changed.
5. **Bank Sync** reads the folder on its next sync (daily between 06:00 and 07:00). Each row carries Trade Republic's
   own event id, so re-reading the file adds nothing.

## The CSV (`src/tr_export/csv_format.py`, read by `trade-republic.js`; a test keeps them in step)

`TR event id;Timestamp;Date;Account;Event type;Type;Title;Subtitle;Amount;Currency;ISIN;Shares;Fees;Taxes;Classification`

- `Date`: the Europe/Brussels day of `Timestamp` (UTC from Trade Republic).
- `Amount`: signed from the cash side, with a `.` decimal point and no exponent. **Saveback** is written as 0, because
  the shares are a gift and no cash leaves the account.
- `Type`: pytr's kind of movement (`BUY`, `SELL`, `DIVIDEND`, `DEPOSIT`, `REMOVAL`, `INTEREST`, ...).
- `Classification` sets the Bank Sync tab:

  | Classification | Events | Tab |
  |---|---|---|
  | `card` | every `CARD_*` event (payments, refunds, ATM) | `transactions` |
  | `transfer` | other deposits and removals (SEPA in/out, direct debits, top-ups) | `transactions` |
  | `investment` | trades, savings plans, round-ups, saveback, dividends, interest, taxes | `investment_transactions` |
  | *(empty)* | a type pytr does not know yet (named in the run log) | `transactions`, so you see it |

- Left out: events without an amount, cancelled events, and types pytr knows carry no money (orders created,
  documents, reports).

## Setup (once)

1. In Terminal, from the repo: `bank-sync/trade-republic/scripts/install-schedule.sh`. It creates `~/trade-republic`,
   builds the `tr-export:local` image, installs the launchd job `lu.niklas.trade-republic`, and copies the
   categoriser's service-account key if it finds one.
2. Set `TR_PHONE` in `~/trade-republic/.env`, then run `~/trade-republic/login.sh` and confirm in the app.
3. Run the job once by hand: `cd ~/trade-republic && docker compose run --rm tr-export`. It writes
   `data/trade-republic.csv` and logs `TR_DRIVE_FILE_ID not set`.
4. Upload that file to the Bank Sync CSV folder (the `CSV_FOLDER_<NAME>` folder) and **share the file with the service
   account's e-mail as Editor**. Put its id (from the URL `/file/d/<id>/view`) in `TR_DRIVE_FILE_ID` in `.env`.
5. Run it again: the log says `Drive file unchanged` (or `updated`). In the Bank Sync sheet, *Import CSV files*, then
   check *Show status*: `trade-republic: 1 file, N transactions, ...`. Add the account's owner on `account_owners`.

After a change to `bank-sync/trade-republic/` is merged, re-run `install-schedule.sh` (there is no auto-update yet).

## Logs and privacy

`~/trade-republic/logs/launchd.out.log` holds counts, date ranges and event type names only: never amounts,
merchants or account numbers. pytr's own log is set to warnings, because at INFO it prints event titles.

## Settings (`.env`)

See `.env.example`: `TR_PHONE`, `TR_ACCOUNT_NAME` (the account name on the tabs and in its id; use another name for a
second person), `TR_DRIVE_FILE_ID`, `TR_LOOKBACK_DAYS`, `TR_TIME_ZONE`.

## Tests

`pytest` in `bank-sync/trade-republic/` (no network: pytr's client and Drive are faked). CI runs them in the
`Tests (trade-republic)` job.
