#!/usr/bin/env node
// Runs the src/tests/*.test.js suites in Node.
//
// Apps Script runs every file in one shared global scope, in filePushOrder.
// We reproduce that with a vm context per test file: the application files
// are evaluated as classic scripts into the same global, in load order, so
// top-level functions, `var`s and `const`s are shared across files exactly as
// they are inside Google.
//
// Each test file gets a fresh context because the suites install their mocks
// on `global` when they load — sharing one global would let the last suite's
// mocks leak into every other suite.
//
// Usage: node scripts/run-tests.mjs [--verbose] [file-substring ...]
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import vm from 'node:vm';
import { REPO_ROOT, scriptsInLoadOrder, listTestScripts } from './lib/source-files.mjs';
import { installGasStubs } from './lib/gas-stubs.mjs';

const args = process.argv.slice(2);
const verbose = args.includes('--verbose');
const filters = args.filter((a) => !a.startsWith('--'));

const appFiles = scriptsInLoadOrder().map((file) => ({
  file,
  script: new vm.Script(readFileSync(join(REPO_ROOT, file), 'utf8'), { filename: file }),
}));
const testFiles = listTestScripts();
const frameworkFile = testFiles.find((f) => f.endsWith('/testing-framework.js'));
if (!frameworkFile) {
  console.error('✖ src/tests/testing-framework.js not found.');
  process.exit(1);
}
// Shared helpers such as gas-fakes.js load after the framework, before suites.
const helpers = testFiles.filter((f) => f !== frameworkFile && !f.endsWith('.test.js'));
const suites = testFiles
  .filter((f) => f.endsWith('.test.js'))
  .filter((f) => filters.length === 0 || filters.some((x) => f.includes(x)));

/** Console that buffers output so passing suites stay quiet. */
function bufferedConsole(lines) {
  const record = (level) => (...a) => lines.push(`[${level}] ${a.map((x) => (x instanceof Error ? x.message : typeof x === 'object' ? JSON.stringify(x) : String(x))).join(' ')}`);
  return { log: record('log'), info: record('info'), warn: record('warn'), error: record('error'), debug: record('debug') };
}

function freshContext(consoleLines = []) {
  const sandbox = { console: bufferedConsole(consoleLines) };
  sandbox.global = sandbox; // the suites mock services via `global.X = …`
  vm.createContext(sandbox);
  installGasStubs(sandbox);
  return sandbox;
}

function run(context, file, script) {
  try {
    (script ?? new vm.Script(readFileSync(join(REPO_ROOT, file), 'utf8'), { filename: file })).runInContext(context);
    return null;
  } catch (error) {
    return error;
  }
}

// Load the app once on its own first: a load-time crash here would break the
// real spreadsheet too, so report it clearly rather than as 80 test failures.
{
  const ctx = freshContext();
  for (const { file, script } of appFiles) {
    const error = run(ctx, file, script);
    if (error) {
      console.error(`✖ ${file} threw while loading — Apps Script would fail the same way.\n${error.stack || error}`);
      process.exit(1);
    }
  }
}

const details = [];
const suiteOutput = new Map();
for (const suite of suites) {
  const lines = [];
  suiteOutput.set(suite, lines);
  const ctx = freshContext(lines);
  for (const { file, script } of appFiles) run(ctx, file, script);
  for (const helper of [frameworkFile, ...helpers]) {
    const error = run(ctx, helper);
    if (error) {
      console.error(`✖ ${helper} threw while loading:\n${error.stack || error}`);
      process.exit(1);
    }
  }
  const loadError = run(ctx, suite);
  if (loadError) {
    details.push({ suite, module: suite, test: '(suite failed to load)', status: 'failed', error: loadError });
    continue;
  }
  const summary = ctx.FinancialPlanner.Testing.runAll();
  for (const d of summary.details) details.push({ suite, ...d });
}

let currentModule = null;
for (const d of details) {
  const heading = `${d.module}  (${d.suite.replace(/^.*\//, '')})`;
  if (heading !== currentModule) {
    currentModule = heading;
    console.log(`\n${heading}`);
  }
  if (d.status === 'passed') console.log(`  ✔ ${d.test}`);
  else console.log(`  ✖ ${d.test}\n      ${d.error?.message ?? d.message}`);
}

const failed = details.filter((d) => d.status === 'failed');
console.log(`\n${details.length - failed.length} passed, ${failed.length} failed, ${details.length} total`);

if (failed.length) {
  if (process.env.GITHUB_ACTIONS) {
    for (const d of failed) {
      console.log(`::error file=${d.suite},title=Test failed: ${d.module}::${d.test} — ${String(d.error?.message ?? '').replace(/\n/g, ' ')}`);
    }
  }
  if (verbose) {
    for (const suite of new Set(failed.map((d) => d.suite))) {
      console.log(`\n── console output of ${suite} ──\n${suiteOutput.get(suite).join('\n')}`);
    }
  } else {
    console.log('Re-run with --verbose to see each failing suite\'s console output.');
  }
  process.exitCode = 1;
}
