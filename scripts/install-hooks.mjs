#!/usr/bin/env node
// Points git at scripts/hooks/ so the pre-commit checks run locally.
// Runs from `npm install` (the "prepare" script); a no-op in CI or outside git.
import { execFileSync } from 'node:child_process';

if (process.env.CI) process.exit(0);
try {
  execFileSync('git', ['rev-parse', '--git-dir'], { stdio: 'ignore' });
  execFileSync('git', ['config', 'core.hooksPath', 'scripts/hooks']);
  console.log('Git hooks enabled (core.hooksPath = scripts/hooks).');
} catch {
  // Not a git checkout (e.g. a downloaded zip) — nothing to install.
}
