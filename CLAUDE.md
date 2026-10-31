# Financial Planning Tools — Agent Instructions

Google Apps Script add-on for a Google Sheets budget: imports bank transactions
(Plaid, SaltEdge) and builds an overview, analysis and reports. Code in `src/`
is pushed to Apps Script with clasp; everything else in the repo is dev tooling
that runs in Node and is never pushed.

The `Claude Code Review` workflow reviews every PR against this file, so keep it
accurate when conventions change.

---

## How Apps Script runs this code (read first)

- **One shared global scope.** Every `.js` file is concatenated into the same
  global scope. A top-level `function`/`var`/`const` in one file is visible in
  all others, and **two files declaring the same name silently overwrite each
  other** (the one loaded last wins).
- **Load order = `filePushOrder` in `.clasp.json`.** Entries must include the
  `rootDir` (`src/...`). Modules that receive dependencies as IIFE arguments,
  e.g. `})(FinancialPlanner.Utils, FinancialPlanner.Config)`, read them **at
  load time**, so the dependency's file must come earlier in the list.
  **Every new file under `src/` must be added to `filePushOrder`.**
- **Namespace:** all modules hang off `FinancialPlanner.*`. Each file starts with
  `var FinancialPlanner = FinancialPlanner || {};` (intentional; ESLint allows it).
- **Menu items and `google.script.run`** call global functions by name. Menu
  actions are exposed as `<name>_Global` via `createGlobalControllerAction`
  in `src/core/controllers.js`, backed by `<name>_Wrapped` on the controller.
- **Triggers:** there is exactly one global `onOpen`, `onEdit` and `doGet`, all
  in `src/core/controllers.js`. Route new edit handling through
  `Controllers.onEdit` — never declare a second `onEdit`.

`npm run check:manifest` enforces all of the above.

### `bank-sync/` (separate project)

`bank-sync/` is an independent Apps Script project (its own `.clasp.json`,
namespace `BankSync`, Node tests via `node --test`). Run clasp from inside that
folder. The rules above about load order and `src/`-prefixed `filePushOrder`
entries apply to it too. CI runs its tests in the `Tests (bank-sync)` job.
Current state, phases and deploy rules: `docs/BANK_SYNC_ROADMAP.md` (read it first).

### `bank-sync/categoriser/` (Python, separate from Apps Script)

Python job that reads the Bank Sync sheet's `transactions` tab and writes only to a
`categorised` tab and its `categoriser_issues` log (never to `transactions`; rows with
`reviewed` = TRUE are never overwritten). Problems such as duplicate `transactionId`
rows are skipped and logged once in `categoriser_issues`, never allowed to stop the
daily run. Not pushed with clasp. Details: `bank-sync/categoriser/README.md`.

- **Secrets and data live outside the repo** in `~/categoriser/` (`.env`, service-account
  key, `data/`, `models/`, `logs/`). Never add them to the repo; the pre-commit hook
  blocks CSVs, `.joblib`, `.env*` and credential files.
- **Runs daily at 06:15** via launchd (`lu.niklas.categoriser`) from `~/categoriser/`,
  using the Docker image `categoriser:local`.
- **Deploying = re-running `bank-sync/categoriser/scripts/install-schedule.sh`.** The image
  is a snapshot: after changing `bank-sync/categoriser/src/**`, `pyproject.toml`,
  `Dockerfile`, `scripts/run-scheduled.sh`, `deploy/*` or the plist template, nothing
  changes in the scheduled job until that script is re-run (from the branch you want
  live). Apps Script changes never need it.
- Tests: `pytest` in `bank-sync/categoriser/` (no network); CI runs them in the `Tests (categoriser)` job.

## Conventions

- Principles: see `docs/DEVELOPMENT_PRINCIPLES.md` (KISS, YAGNI, SRP, DRY).
- Sheet names, columns, transaction types, colours and formats live in
  `src/core/config.js`. Don't hard-code them elsewhere — e.g. expense types come
  from `config.getSection('EXPENSE_TYPES')` (`Essentials`, `Wants/Pleasure`, `Extra`).
- Amounts: expenses are **negative** numbers in sheets. Use `Math.abs` when you
  need magnitudes, and sort by magnitude when ranking spending.
- Spreadsheet I/O: read with one `getValues()` and write with one `setValues()`;
  avoid `getValue`/`setValue` inside loops (each call is a round trip).
- Charts: `Sheet.clear()`/`clearContent()` do **not** remove charts. Remove old
  charts before inserting new ones when rebuilding a sheet.
- Only call Apps Script APIs that exist — `npm run typecheck` checks against
  `@types/google-apps-script`.
- Errors: user-facing failures go through `FinancialPlanner.ErrorService.handle`.

## Security & privacy

- **Secrets** (Plaid keys, SaltEdge app ID/secret/RSA private key, access tokens)
  live only in `PropertiesService`. Never commit them, log them, or put them in
  sheets or URLs. Gitleaks scans every PR and the full history.
- **No personal financial data in logs.** Don't `Logger.log`/`console.log`
  transactions, balances or account numbers.
- **Escape before rendering.** Anything from `e.parameter`, an API response or a
  sheet that goes into HtmlService output must be escaped; in client HTML use
  `textContent`, not `innerHTML`.
- **OAuth scopes:** adding a scope to `src/appsscript.json` requires adding it to
  `scripts/allowed-scopes.json` too, deliberately, with a reason in the PR.
- Keep `appsscript.json` (root) identical to `src/appsscript.json`; only the
  `src/` copy is pushed.

## Tests

- Suites live in `src/tests/*.test.js` and use the small framework in
  `src/tests/testing-framework.js`. `tests/**` is in `.claspignore`, so tests
  run only in Node.
- `npm test` loads the real app files in `filePushOrder` into a fresh VM context
  per suite, then runs the suite. Test the **real** `FinancialPlanner.*` modules
  — never paste a copy of the implementation into a test.
- Use `FakeGas.install(global)` from `src/tests/gas-fakes.js` for an in-memory
  Spreadsheet/Ui/HtmlService. Values written through ranges can be read back;
  unimplemented methods throw, so extend the fake when the code needs more.
- New logic needs a test. Fix the code, not the assertion, unless the
  assertion is wrong.

## Commands

```bash
npm install            # dev tools + git pre-commit hook
npm run lint           # ESLint (Apps Script globals derived from src/)
npm run check:manifest # push order, scopes, globals, load order, menu handlers
npm test               # Node test runner (add --verbose for console output)
npm run typecheck      # TypeScript JSDoc check; fails only on errors not in the baseline
npm run check          # all of the above
```

`npm run typecheck -- --update` rewrites `scripts/typecheck-baseline.json`. Only
use it after *removing* errors; never baseline a new error.

## Git & PRs

- Never commit to `main`; branch (`feat/…`, `fix/…`, `chore/…`) and open a PR.
- PRs are squash-merged; the **PR title must be a Conventional Commit**
  (`feat(scope): …`, `fix(scope): …`, `chore: …`) — CI enforces it.
- Include a short summary and a test plan in the PR body.
- **Open every PR as a draft.** Mark it *Ready for review* only when CI is green and
  the work is done: that is what starts the Claude review, once. Don't push more
  work after marking ready unless it fixes review findings.
- **Review findings are issues.** The review files each bug and improvement as its
  own issue and all nits as one, labelled `code-review` (plus `blocking`,
  `enhancement` or `nits`), and links them from its PR comment. The session that
  opened the PR fixes them on the same branch; otherwise the owner comments
  `@claude` on the issue. Close the issue when the fix lands.
- **Report review fixes on the PR.** After pushing fixes for review findings (issues
  or review comments), post one comment on the PR that lists each finding with its
  issue number and the commit that fixed it, and each finding you skipped with the
  reason. Do this even when nothing was skipped.
- **Don't re-run the review yourself** (no `workflow_dispatch`, no draft/ready
  toggling) unless the owner asks. To get a fresh review after big changes, the
  owner converts the PR to draft and marks it ready again.

## CI (`.github/workflows/`)

| Workflow | What it does |
|---|---|
| `ci.yml` | PR-title commitlint, ESLint, tests, type check, Apps Script check (sticky PR comment with deltas vs `main`) |
| `security.yml` | Gitleaks (full history) + Semgrep (JS, OWASP, GitHub Actions rules), weekly too |
| `claude-code-review.yml` | When a PR is marked ready: runs the checks on the PR head, then Claude reviews it against this file, posts one comment and files findings as `code-review` issues |
| `claude.yml` | `@claude` in issues/PRs (owner/collaborators only), e.g. to fix a review issue |
| `auto-fix-ci.yml` | Claude fixes a red CI run on a PR branch and pushes the fix |
| `deploy.yml` | Manual `clasp push` from `main` (see below) |

## Deploying

Deploys are manual. Locally use the globally installed `clasp` (`clasp push`), never
`npx clasp`. Rules, learned the hard way:

- Run clasp from the folder of the project you mean: the repo root for the legacy
  `src/` project, `bank-sync/` for bank-sync. Never from the other one.
- Before every push, check the checked-out branch contains the code you intend to
  deploy; pushing from a branch without it overwrites the live script.
- After a bank-sync change that affects the web app, also run
  `clasp deploy -i <deployment id>` (see `clasp deployments`), otherwise `/exec` keeps
  serving the old version.
- Run `npm run check` (repo root) and `npm test` (`bank-sync/`) before every push.
- If `clasp push` asks to overwrite a changed manifest, diff the remote one first.

From GitHub: run **Deploy to Apps Script** on `main`; it runs the checks, pushes
the legacy project, and can create a version and point a deployment (e.g. the SaltEdge
callback web app) at it. It needs the `CLASPRC_JSON` secret — the contents of
`~/.clasprc.json` after `clasp login` — in the `apps-script` environment. It does not
deploy `bank-sync/`.
