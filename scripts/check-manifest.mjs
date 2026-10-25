#!/usr/bin/env node
// Apps Script project consistency check — the Apps Script equivalent of
// WordPress's Plugin Check. Catches problems that only show up after a
// `clasp push`, when the script runs inside Google:
//
//   push-order      .clasp.json filePushOrder lists missing files or misses pushed ones
//   manifest-sync   root appsscript.json and the pushed copy have drifted
//   scopes          an OAuth scope was added that is not in scripts/allowed-scopes.json
//   timezone        appsscript.json timeZone is not a valid IANA zone
//   tests-pushed    files under tests/ would be pushed to Apps Script
//   duplicate-global  two files declare the same global (the last one loaded silently wins)
//   undefined-namespace  code reads FinancialPlanner.X but nothing defines X
//   load-order      a file reads FinancialPlanner.X at load time before X's file has loaded
//   menu-handler    a menu item or google.script.run call names a function that doesn't exist
//   unknown-member  code uses FinancialPlanner.X.y but module X has no member y
//
// Usage:
//   node scripts/check-manifest.mjs                  # human-readable, exit 1 on errors
//   node scripts/check-manifest.mjs --json
//   node scripts/check-manifest.mjs --markdown out.md [--base <dir>] [--base-ref main]
import { readFileSync, existsSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import {
  REPO_ROOT,
  readClaspConfig,
  contentDir,
  listContentFiles,
  listPushedScripts,
  isTestFile,
  scriptsInLoadOrder,
  parse,
  visit,
  topLevelNames,
} from './lib/source-files.mjs';

const NAMESPACE = 'FinancialPlanner';

function readJson(path) {
  return JSON.parse(readFileSync(path, 'utf8'));
}

/** `FinancialPlanner.X` → 'X' when node is that member expression. */
function namespaceMember(node) {
  if (
    node.type === 'MemberExpression' &&
    !node.computed &&
    node.object.type === 'Identifier' &&
    node.object.name === NAMESPACE &&
    node.property.type === 'Identifier'
  ) {
    return node.property.name;
  }
  return null;
}

const isFunctionNode = (n) =>
  n.type === 'FunctionDeclaration' ||
  n.type === 'FunctionExpression' ||
  n.type === 'ArrowFunctionExpression';

/** Key names of an object literal (identifier and string keys only). */
function objectKeys(node) {
  const keys = new Set();
  for (const p of node.properties) {
    if (p.type !== 'Property' || p.computed) return null; // spread etc.: shape unknown
    if (p.key.type === 'Identifier') keys.add(p.key.name);
    else if (p.key.type === 'Literal') keys.add(String(p.key.value));
  }
  return keys;
}

/**
 * The members a module assignment exposes, when they can be read statically:
 * `X = { ... }` or `X = (function () { ...; return { ... }; })(...)`.
 * Returns null when the shape cannot be determined (e.g. `new Foo()`).
 */
function moduleShape(value) {
  if (value.type === 'ObjectExpression') return objectKeys(value);
  if (value.type === 'CallExpression' && isFunctionNode(value.callee) && value.callee.body.type === 'BlockStatement') {
    const returns = value.callee.body.body.filter((n) => n.type === 'ReturnStatement');
    const last = returns[returns.length - 1];
    if (last && last.argument && last.argument.type === 'ObjectExpression') return objectKeys(last.argument);
  }
  return null;
}

/**
 * Visits only the code that runs when the file is loaded: top-level
 * statements plus the bodies of immediately-invoked functions. Bodies of
 * functions that are merely defined are skipped.
 */
function visitLoadTime(node, fn) {
  if (!node || typeof node.type !== 'string') return;
  if (node.type === 'CallExpression' && isFunctionNode(node.callee)) {
    fn(node);
    node.arguments.forEach((a) => visitLoadTime(a, fn));
    visitLoadTime(node.callee.body, fn);
    return;
  }
  if (isFunctionNode(node) || node.type === 'ClassDeclaration' || node.type === 'ClassExpression') {
    return;
  }
  fn(node);
  for (const key of Object.keys(node)) {
    if (key === 'loc' || key === 'range') continue;
    const child = node[key];
    if (Array.isArray(child)) child.forEach((c) => visitLoadTime(c, fn));
    else if (child && typeof child.type === 'string') visitLoadTime(child, fn);
  }
}

/** Collects everything the checks and the PR summary need from one project root. */
export function analyse(root = REPO_ROOT) {
  const findings = [];
  const add = (rule, message, file) => findings.push({ rule, message, file });

  const clasp = readClaspConfig(root);
  const dir = contentDir(root);
  const allFiles = listContentFiles(root);
  const pushedScripts = listPushedScripts(root);
  const pushedHtml = allFiles.filter((f) => f.endsWith('.html') && !isTestFile(f, root));
  const manifestPath = join(root, dir, 'appsscript.json');
  const manifest = existsSync(manifestPath) ? readJson(manifestPath) : {};

  // ── push-order ──────────────────────────────────────────────────────────
  const pushOrder = clasp.filePushOrder || [];
  for (const entry of pushOrder) {
    if (!allFiles.includes(entry)) {
      add(
        'push-order',
        `filePushOrder lists "${entry}", which does not exist. Paths must include the rootDir ("${dir}/…").`,
        '.clasp.json',
      );
    }
  }
  if (pushOrder.length > 0) {
    for (const f of pushedScripts) {
      if (!pushOrder.includes(f)) {
        add('push-order', `"${f}" is pushed but missing from filePushOrder, so its load position is undefined.`, '.clasp.json');
      }
    }
  }

  // ── manifest-sync ───────────────────────────────────────────────────────
  const rootManifestPath = join(root, 'appsscript.json');
  if (dir !== '.' && existsSync(rootManifestPath)) {
    if (!isDeepStrictEqual(readJson(rootManifestPath), manifest)) {
      add('manifest-sync', `appsscript.json differs from ${dir}/appsscript.json. Only ${dir}/appsscript.json is pushed; keep them identical.`, 'appsscript.json');
    }
  }

  // ── scopes ──────────────────────────────────────────────────────────────
  const scopes = manifest.oauthScopes || [];
  const allowedPath = join(root, 'scripts', 'allowed-scopes.json');
  if (existsSync(allowedPath)) {
    const allowed = readJson(allowedPath).scopes || [];
    for (const s of scopes) {
      if (!allowed.includes(s)) {
        add('scopes', `OAuth scope "${s}" is not in scripts/allowed-scopes.json. Add it there deliberately if it is needed.`, `${dir}/appsscript.json`);
      }
    }
    // bank-sync/ is a separate Apps Script project; its scopes go through the same allow-list.
    const bankSyncManifest = join(root, 'bank-sync', 'src', 'appsscript.json');
    if (existsSync(bankSyncManifest)) {
      for (const s of readJson(bankSyncManifest).oauthScopes || []) {
        if (!allowed.includes(s)) {
          add('scopes', `OAuth scope "${s}" is not in scripts/allowed-scopes.json. Add it there deliberately if it is needed.`, 'bank-sync/src/appsscript.json');
        }
      }
    }
  }

  // ── timezone ────────────────────────────────────────────────────────────
  try {
    new Intl.DateTimeFormat('en', { timeZone: manifest.timeZone });
    if (!manifest.timeZone) throw new Error('missing');
  } catch {
    add('timezone', `timeZone "${manifest.timeZone}" is not a valid IANA time zone.`, `${dir}/appsscript.json`);
  }

  // ── tests-pushed ────────────────────────────────────────────────────────
  const ignorePath = join(root, '.claspignore');
  const ignore = existsSync(ignorePath) ? readFileSync(ignorePath, 'utf8') : '';
  if (allFiles.some((f) => isTestFile(f, root)) && !/^tests\/\*\*\s*$/m.test(ignore)) {
    add('tests-pushed', '.claspignore does not exclude tests/**, so test files would be pushed to Apps Script.', '.claspignore');
  }

  // ── source analysis ─────────────────────────────────────────────────────
  const loadOrder = scriptsInLoadOrder(root);
  const globals = new Map(); // name -> [file]
  const defined = new Map(); // namespace member -> first file (load order index)
  const reads = []; // { name, file }
  const loadTimeReads = []; // { name, file, index }
  const shapes = new Map(); // module name -> Set of member names (or null if unknown)
  const extraMembers = new Map(); // module name -> members added later via X.y = ...
  const memberReads = []; // { module, member, file, line }
  const menuHandlers = []; // { name, file }
  const controllerActions = new Set();
  const wrappedKeys = new Set();
  let parseErrors = 0;

  loadOrder.forEach((file, index) => {
    let ast;
    try {
      ast = parse(readFileSync(join(root, file), 'utf8'));
    } catch (e) {
      parseErrors++;
      add('syntax', `Could not parse: ${e.message}`, file);
      return;
    }
    for (const { name, line } of topLevelNames(ast)) {
      if (name === NAMESPACE) continue;
      if (!globals.has(name)) globals.set(name, []);
      globals.get(name).push(`${file}:${line}`);
    }
    visit(ast, (node) => {
      if (node.type === 'AssignmentExpression') {
        const name = namespaceMember(node.left);
        if (name && !defined.has(name)) defined.set(name, { file, index });
        if (name) shapes.set(name, shapes.has(name) ? null : moduleShape(node.right));
        // FinancialPlanner.X.y = ... adds a member after the fact.
        if (node.left.type === 'MemberExpression' && !node.left.computed && node.left.property.type === 'Identifier') {
          const owner = namespaceMember(node.left.object);
          if (owner) {
            if (!extraMembers.has(owner)) extraMembers.set(owner, new Set());
            extraMembers.get(owner).add(node.left.property.name);
          }
        }
      }
      if (node.type === 'MemberExpression' && !node.computed && node.property.type === 'Identifier') {
        const owner = namespaceMember(node.object);
        if (owner) memberReads.push({ module: owner, member: node.property.name, file: `${file}:${node.loc.start.line}` });
      }
      const member = namespaceMember(node);
      if (member) reads.push({ name: member, file });
      if (node.type === 'Property' && node.key.type === 'Identifier' && node.key.name.endsWith('_Wrapped')) {
        wrappedKeys.add(node.key.name.replace(/_Wrapped$/, ''));
      }
      if (node.type === 'CallExpression') {
        const callee = node.callee;
        const arg = node.arguments;
        if (callee.type === 'Identifier' && callee.name === 'createGlobalControllerAction' && arg[0]?.type === 'Literal') {
          controllerActions.add(arg[0].value);
        }
        if (callee.type === 'MemberExpression' && callee.property.name === 'addItem' && arg[1]?.type === 'Literal') {
          menuHandlers.push({ name: arg[1].value, file: `${file}:${node.loc.start.line}` });
        }
      }
    });
    // Reads that happen while the file loads (IIFE arguments, top-level code).
    const assignedHere = new Set();
    visitLoadTime(ast, (node) => {
      if (node.type === 'AssignmentExpression') {
        const name = namespaceMember(node.left);
        if (name) assignedHere.add(name);
        return;
      }
      if (node.type === 'UnaryExpression' && node.operator === 'typeof') return;
      const name = namespaceMember(node);
      if (name && !assignedHere.has(name)) loadTimeReads.push({ name, file, index });
    });
  });

  for (const [name, places] of globals) {
    if (places.length > 1) {
      add('duplicate-global', `Global "${name}" is declared ${places.length} times (${places.join(', ')}). Apps Script keeps only the last one loaded.`, places[0].split(':')[0]);
    }
  }

  const reported = new Set();
  for (const { name, file } of reads) {
    // Fields like VERSION/META are plain values; still need a definition.
    if (!defined.has(name) && !reported.has(name)) {
      reported.add(name);
      add('undefined-namespace', `${NAMESPACE}.${name} is used but never assigned anywhere in the pushed code.`, file);
    }
  }

  for (const { name, file, index } of loadTimeReads) {
    const def = defined.get(name);
    if (def && def.index > index) {
      add('load-order', `${file} reads ${NAMESPACE}.${name} while loading, but it is defined later in ${def.file}. Move ${def.file} earlier in filePushOrder.`, file);
    }
  }

  const reportedMembers = new Set();
  for (const { module, member, file } of memberReads) {
    const shape = shapes.get(module);
    if (!shape) continue; // unknown shape (constructor, reassigned, etc.)
    if (shape.has(member) || extraMembers.get(module)?.has(member)) continue;
    const key = `${module}.${member}`;
    if (reportedMembers.has(key)) continue;
    reportedMembers.add(key);
    add('unknown-member', `${NAMESPACE}.${module}.${member} is used, but ${module} has no member "${member}".`, file);
  }

  const callable = new Set([
    ...[...globals.keys()],
    ...[...controllerActions].map((a) => `${a}_Global`),
  ]);
  for (const { name, file } of menuHandlers) {
    if (!callable.has(name)) {
      add('menu-handler', `Menu item calls "${name}", which is not a global function.`, file);
    }
  }
  for (const action of controllerActions) {
    if (!wrappedKeys.has(action)) {
      add('menu-handler', `createGlobalControllerAction('${action}') has no matching ${action}_Wrapped controller method.`, 'src/core/controllers.js');
    }
  }
  for (const html of pushedHtml) {
    const source = readFileSync(join(root, html), 'utf8');
    for (const m of source.matchAll(/google\.script\.run[\s\S]*?\.([A-Za-z_$][\w$]*)\s*\(/g)) {
      // The chain is `.withSuccessHandler(...)` etc. followed by the target call.
      const name = m[1];
      if (/^with(Success|Failure)Handler$|^withUserObject$/.test(name)) continue;
      if (!callable.has(name)) {
        add('menu-handler', `google.script.run calls "${name}", which is not a global function.`, html);
      }
    }
  }

  const sizeBytes = [...pushedScripts, ...pushedHtml, `${dir}/appsscript.json`]
    .filter((f) => existsSync(join(root, f)))
    .reduce((sum, f) => sum + readFileSync(join(root, f)).length, 0);

  return {
    findings,
    metrics: {
      scripts: pushedScripts.length,
      html: pushedHtml.length,
      sizeBytes,
      globals: globals.size,
      namespaces: defined.size,
      menuItems: menuHandlers.length,
      scopes,
      timeZone: manifest.timeZone || '',
      runtime: manifest.runtimeVersion || '',
    },
    parseErrors,
  };
}

// ── Output ────────────────────────────────────────────────────────────────

function fmtDelta(n) {
  if (!n) return '';
  return n > 0 ? `(+${n})` : `(${n})`;
}

function fmtBytes(n) {
  return n < 1024 ? `${n} B` : `${(n / 1024).toFixed(1)} KB`;
}

function markdown(result, base, baseRef) {
  const m = result.metrics;
  const b = base?.metrics;
  const row = (label, value, delta) => `| ${label} | ${value} | ${b ? delta || '—' : '—'} |`;
  const lines = [
    '## 🔍 Apps Script Check — Summary',
    '',
    result.findings.length === 0
      ? '✅ No problems found.'
      : `❌ **${result.findings.length} problem(s) found** — see the list below.`,
    '',
    `| Metric | Value | Δ vs \`${baseRef}\` |`,
    '|---|---|---|',
    row('Script files pushed', m.scripts, b && fmtDelta(m.scripts - b.scripts)),
    row('HTML files pushed', m.html, b && fmtDelta(m.html - b.html)),
    row('Pushed size', fmtBytes(m.sizeBytes), b && (m.sizeBytes !== b.sizeBytes ? `(${m.sizeBytes > b.sizeBytes ? '+' : '-'}${fmtBytes(Math.abs(m.sizeBytes - b.sizeBytes))})` : '')),
    row('Global functions/vars', m.globals, b && fmtDelta(m.globals - b.globals)),
    row('`FinancialPlanner.*` modules', m.namespaces, b && fmtDelta(m.namespaces - b.namespaces)),
    row('Menu items', m.menuItems, b && fmtDelta(m.menuItems - b.menuItems)),
    row('Time zone', `\`${m.timeZone}\``, b && b.timeZone !== m.timeZone ? `was \`${b.timeZone}\`` : ''),
  ];

  const added = b ? m.scopes.filter((s) => !b.scopes.includes(s)) : [];
  const removed = b ? b.scopes.filter((s) => !m.scopes.includes(s)) : [];
  lines.push('', `**OAuth scopes (${m.scopes.length}):**`);
  for (const s of m.scopes) lines.push(`- \`${s}\`${added.includes(s) ? ' 🆕 **added in this PR**' : ''}`);
  for (const s of removed) lines.push(`- ~~\`${s}\`~~ removed in this PR`);

  if (result.findings.length) {
    lines.push('', '### Problems', '', '| Rule | File | Problem |', '|---|---|---|');
    for (const f of result.findings) {
      lines.push(`| \`${f.rule}\` | \`${f.file}\` | ${f.message.replace(/\|/g, '\\|')} |`);
    }
  }
  if (!b) {
    lines.push('', `_Baseline for \`${baseRef}\` was unavailable, so deltas are not shown._`);
  }
  lines.push('', '<sub>Posted by `scripts/check-manifest.mjs` in the CI workflow. This comment updates in place on each push.</sub>');
  return lines.join('\n') + '\n';
}

const isMain = import.meta.url === `file://${process.argv[1]}`;
if (isMain) {
  const args = process.argv.slice(2);
  const argValue = (flag) => {
    const i = args.indexOf(flag);
    return i === -1 ? undefined : args[i + 1];
  };
  const result = analyse();

  if (args.includes('--json')) {
    console.log(JSON.stringify(result, null, 2));
  } else {
    if (result.findings.length === 0) {
      console.log(`✔ Apps Script check passed (${result.metrics.scripts} scripts, ${result.metrics.html} HTML files).`);
    } else {
      for (const f of result.findings) {
        // GitHub Actions annotation format, readable locally too.
        console.log(`::error file=${f.file.split(':')[0]}::[${f.rule}] ${f.message}`);
      }
      console.log(`\n✖ ${result.findings.length} problem(s) found.`);
    }
  }

  const mdOut = argValue('--markdown');
  if (mdOut) {
    const baseDir = argValue('--base');
    let base = null;
    if (baseDir && existsSync(join(baseDir, '.clasp.json'))) {
      try {
        base = analyse(baseDir);
      } catch (e) {
        console.warn(`Baseline analysis failed: ${e.message}`);
      }
    }
    writeFileSync(mdOut, markdown(result, base, argValue('--base-ref') || 'main'));
  }

  process.exitCode = result.findings.length > 0 ? 1 : 0;
}
