# Handoff: Enable Banking bank-sync project

Prompt for the VS Code instance. It assumes the repo context is already known and only adds what is new: Enable Banking, SEK/EUR handling, modularity and the PR workflow. Scope is setup plus the first read-only connection only.

```
TASK: Set up a new, modular bank-sync project in this repo, using Enable Banking as the data source. Scope of THIS session is setup plus the first read-only connection only (Phases 0-1 below). Stop after Phase 1 and report. Do not port legacy sheets or formulas yet.

NEW CONTEXT YOU DON'T HAVE
- Enable Banking (https://enablebanking.com, API base https://api.enablebanking.com) is our new bank-data provider, replacing Salt Edge and Plaid. I registered a PRODUCTION application in "restricted" mode: free, personal use, and it can only read accounts I have linked in the Enable Banking Control Panel. Revolut is linked. Other banks (BGL, Spuerkeess, Raiffeisen, Danske Bank, SEB) will be linked later. BIL is NOT supported and is parked. Avanza is NOT supported, so it will use CSV import later.
- Auth model: every API call carries a short-lived JWT signed RS256 with my application's private key (the application ID is the JWT "kid"). VERIFY the exact header/claims, the auth/session flow, and the transactions endpoint and pagination against the official docs (https://enablebanking.com/docs/api/reference/) before coding. Do not rely on my summary. Also verify whether linked accounts in restricted mode can be queried directly or still need the app's own authorization/session flow, and whether the application's redirect URL can be edited later (it is currently the placeholder https://localhost:8080/callback; the plan is to swap in the Apps Script web-app URL).
- Consent expires every 90-180 days per bank and must be renewed by re-login. Store consent expiry and warn me by email before it lapses.
- Apps Script can sign with Utilities.computeRsaSha256Signature. Confirm the key format my downloaded .pem uses (PKCS#8 PEM expected).

SECRETS (non-negotiable)
- The private key and application ID live ONLY in Script Properties, pasted by me. Never in the repo, logs, chat or sheet. Add *.pem, *.key, .env and any clasp credential files to .gitignore. Never log tokens or IBANs. Use fake data in tests and docs.
- Any web app deployment is "execute as me, accessible only by me".

DECISIONS ALREADY MADE
- A completely NEW Google Sheet, created in my "Financial Planning" Drive folder. The legacy sheet and code are read-only inspiration for later; do not modify them. Propose how to separate the new project from the legacy one in this repo (own folder/rootDir and own .clasp.json so clasp never overwrites the old script project), and ask me before restructuring anything.
- Development flow: VS Code + clasp for pushing to the new Sheet's bound script. All changes reach GitHub via pull requests (feature branch per module, small PRs, conventional commit messages, never push to main).
- Reporting currency is EUR. I hold SEK and EUR accounts.

ARCHITECTURE REQUIREMENTS
- Follow SRP and DRY. Small functions, no duplicated logic, no globals leaking. Reuse the repo's existing module-loader convention if it fits.
- JSDoc on every function, class and typedef (@param, @returns, @throws, @typedef for data models).
- Provider abstraction: one common BankProvider interface (authorize, listAccounts, getBalances, getTransactions) with separate implementations: EnableBankingProvider now; Trading212Provider next (it has its own API key and returns holdings/cash, not bank transactions, so model it explicitly); CsvImportProvider later (Avanza, BIL). Adding a provider must not require touching existing ones.
- Normalized models: Transaction {provider, accountId, transactionId, bookingDate, valueDate, status (pending|booked), amount, currency, amountEur, fxRate, fxRateDate, counterparty, description}, Account {...}, Balance {...}. Keep the original amount+currency; add EUR as derived columns.
- Idempotent imports: dedupe/upsert on provider+accountId+transactionId; handle pending to booked transitions.
- FX: a separate, injectable FxService. Use ECB reference rates (e.g. the free Frankfurter API, which republishes ECB rates; verify its current endpoint), convert each SEK transaction at the rate of its booking date, fall back to the previous business day for weekends/holidays, and convert balances at the latest rate. Cache rates in a Rates tab (and CacheService), fetching each date range once, never per row.
- Performance: batch sheet reads and writes (getValues/setValues, never per-cell calls), minimize UrlFetchApp calls (fetchAll where useful), paginate, respect the 6-minute limit with resumable chunks, guard time-driven triggers with LockService.
- Testability: keep pure logic (normalization, FX conversion, dedupe) free of Apps Script services so it can be unit-tested locally with Node. Pass services in as dependencies.

PHASES
0. Setup: propose the folder layout, create the new Sheet and clasp project, .gitignore secrets, README with setup steps, and a Config module. No secrets in code.
1. EnableBankingProvider auth: JWT signing, then read-only listAccounts and getBalances for Revolut, written to a raw tab. Confirm it works and report the doc-verified findings listed above.
(Later, for reference only: 2 transactions + sheet writer + dedupe; 3 FX; 4 triggers + consent-expiry alerts; 5 Trading212; 6 CSV import.)

FIRST ACTIONS
Read the repo, then post your proposed folder layout and the doc-verified Enable Banking flow for my approval before writing code. Open one PR per phase. Tell me exactly which steps I must do manually (e.g. pasting the key into Script Properties).
```
