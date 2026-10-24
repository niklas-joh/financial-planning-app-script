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

## CI (`.github/workflows/`)

| Workflow | What it does |
|---|---|
| `ci.yml` | PR-title commitlint, ESLint, tests, type check, Apps Script check (sticky PR comment with deltas vs `main`) |
| `security.yml` | Gitleaks (full history) + Semgrep (JS, OWASP, GitHub Actions rules), weekly too |
| `claude-code-review.yml` | Claude reviews each PR against this file |
| `claude.yml` | `@claude` in issues/PRs (owner/collaborators only) |
| `auto-fix-ci.yml` | Claude fixes a red CI run on a PR branch and pushes the fix |
| `deploy.yml` | Manual `clasp push` from `main` (see below) |

## Deploying

Deploys are manual. Locally: `npx clasp push` (uses the clasp version pinned in
`package.json`). From GitHub: run **Deploy to Apps Script** on `main`; it runs
the checks, pushes, and can create a version and point a deployment (e.g. the
SaltEdge callback web app) at it. It needs the `CLASPRC_JSON` secret — the
contents of `~/.clasprc.json` after `npx clasp login` — in the `apps-script`
environment.
