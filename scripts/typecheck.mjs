#!/usr/bin/env node
// Type-checks src/ with TypeScript (JSDoc + @types/google-apps-script), the
// way Plume runs PHPStan with a baseline.
//
// The existing code has ~100 JSDoc-only errors (undeclared typedef names and
// similar). Those are recorded in scripts/typecheck-baseline.json; CI fails
// only on errors that are NOT in the baseline, so new mistakes — like calling
// an Apps Script method that does not exist — are caught without first
// fixing every historical annotation.
//
// Entries are keyed by file + error code + message (no line numbers), so
// unrelated edits that shift lines do not break the baseline.
//
// Usage:
//   node scripts/typecheck.mjs              # fail on new errors
//   node scripts/typecheck.mjs --update     # rewrite the baseline (after fixing errors)
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { REPO_ROOT } from './lib/source-files.mjs';

const BASELINE = join(REPO_ROOT, 'scripts', 'typecheck-baseline.json');
const update = process.argv.includes('--update');

function runTsc() {
  try {
    execFileSync(process.execPath, [join(REPO_ROOT, 'node_modules', 'typescript', 'bin', 'tsc'), '-p', 'tsconfig.json', '--pretty', 'false'], {
      cwd: REPO_ROOT,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    return '';
  } catch (e) {
    if (e.stdout === undefined) throw e;
    return e.stdout;
  }
}

/** Parses tsc output into [{ file, line, col, code, message, key }]. */
function parse(output) {
  const errors = [];
  for (const line of output.split('\n')) {
    const m = /^(.+?)\((\d+),(\d+)\): error (TS\d+): (.*)$/.exec(line);
    if (!m) continue;
    const [, file, ln, col, code, message] = m;
    errors.push({ file, line: +ln, col: +col, code, message, key: `${file} ${code} ${message}` });
  }
  return errors;
}

const errors = parse(runTsc());
const counts = {};
for (const e of errors) counts[e.key] = (counts[e.key] || 0) + 1;

if (update) {
  const sorted = Object.fromEntries(Object.entries(counts).sort(([a], [b]) => a.localeCompare(b)));
  writeFileSync(BASELINE, JSON.stringify(sorted, null, 2) + '\n');
  console.log(`Baseline updated: ${errors.length} known error(s) in ${Object.keys(sorted).length} entries.`);
  process.exit(0);
}

const baseline = existsSync(BASELINE) ? JSON.parse(readFileSync(BASELINE, 'utf8')) : {};
const remaining = { ...baseline };
const fresh = [];
for (const e of errors) {
  if (remaining[e.key] > 0) remaining[e.key]--;
  else fresh.push(e);
}
const fixed = Object.values(remaining).reduce((a, b) => a + b, 0);

if (fresh.length) {
  for (const e of fresh) {
    console.log(`::error file=${e.file},line=${e.line},col=${e.col}::${e.code} ${e.message}`);
  }
  console.log(`\n✖ ${fresh.length} new type error(s). Fix them, or — only if the code is right and the annotation is wrong — fix the JSDoc.`);
  process.exitCode = 1;
} else {
  console.log(`✔ No new type errors (${errors.length} known in baseline).`);
}
if (fixed > 0) {
  console.log(`ℹ ${fixed} baseline error(s) no longer occur — run \`npm run typecheck -- --update\` to shrink the baseline.`);
}
