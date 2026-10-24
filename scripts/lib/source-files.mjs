// Shared helpers for reading the Apps Script project the way clasp sees it.
import { readFileSync, readdirSync, statSync, existsSync } from 'node:fs';
import { join, relative, sep, posix } from 'node:path';
import { fileURLToPath } from 'node:url';
import * as espree from 'espree';

export const REPO_ROOT = fileURLToPath(new URL('../..', import.meta.url));

/** Reads .clasp.json from a project root (defaults to this repo). */
export function readClaspConfig(root = REPO_ROOT) {
  return JSON.parse(readFileSync(join(root, '.clasp.json'), 'utf8'));
}

/** Content directory clasp pushes from, relative to the project root. */
export function contentDir(root = REPO_ROOT) {
  const cfg = readClaspConfig(root);
  return (cfg.srcDir || cfg.rootDir || '.').replace(/\/+$/, '');
}

function walk(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const full = join(dir, name);
    if (statSync(full).isDirectory()) out.push(...walk(full));
    else out.push(full);
  }
  return out;
}

const toPosix = (p) => p.split(sep).join(posix.sep);

/**
 * Lists every file under the content directory, as POSIX paths relative to
 * the project root (the same form clasp compares `filePushOrder` against).
 */
export function listContentFiles(root = REPO_ROOT) {
  const dir = join(root, contentDir(root));
  if (!existsSync(dir)) return [];
  return walk(dir)
    .map((f) => toPosix(relative(root, f)))
    .sort();
}

/** True for files under the content dir's tests/ folder (ignored by .claspignore). */
export function isTestFile(path, root = REPO_ROOT) {
  return path.startsWith(`${contentDir(root)}/tests/`);
}

/** Script files that clasp pushes (.js/.gs), excluding tests. */
export function listPushedScripts(root = REPO_ROOT) {
  return listContentFiles(root).filter(
    (f) => /\.(js|gs)$/.test(f) && !isTestFile(f, root),
  );
}

export function listTestScripts(root = REPO_ROOT) {
  return listContentFiles(root).filter((f) => f.endsWith('.js') && isTestFile(f, root));
}

/**
 * Returns pushed scripts in the order Apps Script will load them: entries in
 * `filePushOrder` first (in that order), then the rest alphabetically. This
 * mirrors clasp 3's sort in core/files.js.
 */
export function scriptsInLoadOrder(root = REPO_ROOT) {
  const order = (readClaspConfig(root).filePushOrder || []).map((p) => posix.normalize(p));
  return listPushedScripts(root).sort((a, b) => {
    const ia = order.indexOf(a);
    const ib = order.indexOf(b);
    if (ia === -1 && ib === -1) return a.localeCompare(b);
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });
}

export function parse(source) {
  return espree.parse(source, { ecmaVersion: 2022, sourceType: 'script', loc: true });
}

/** Minimal recursive AST walker (no dependency on estraverse). */
export function visit(node, fn) {
  if (!node || typeof node.type !== 'string') return;
  fn(node);
  for (const key of Object.keys(node)) {
    if (key === 'loc' || key === 'range') continue;
    const child = node[key];
    if (Array.isArray(child)) child.forEach((c) => visit(c, fn));
    else if (child && typeof child.type === 'string') visit(child, fn);
  }
}

/** Top-level function/var/let/const/class names declared in one file. */
export function topLevelNames(ast) {
  const names = [];
  for (const node of ast.body) {
    if (node.type === 'FunctionDeclaration' || node.type === 'ClassDeclaration') {
      names.push({ name: node.id.name, kind: node.type, line: node.loc.start.line });
    } else if (node.type === 'VariableDeclaration') {
      for (const d of node.declarations) {
        if (d.id.type === 'Identifier') {
          names.push({ name: d.id.name, kind: node.kind, line: node.loc.start.line });
        }
      }
    }
  }
  return names;
}

/** Every top-level name across all scripts under src/ (pushed + tests). */
export function collectTopLevelNames(root = REPO_ROOT) {
  const names = new Set();
  for (const file of [...listPushedScripts(root), ...listTestScripts(root)]) {
    try {
      for (const { name } of topLevelNames(parse(readFileSync(join(root, file), 'utf8')))) {
        names.add(name);
      }
    } catch {
      // Syntax errors are reported by ESLint itself.
    }
  }
  return [...names].sort();
}
