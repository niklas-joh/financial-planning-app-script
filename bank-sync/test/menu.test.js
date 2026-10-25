// Menu summaries: shown in a modeless window (the run ends without waiting for a click), escaped,
// and only after the script lock is released. Apps Script globals and services are faked.
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');

// Entry files are not loaded by load.js; this one only touches globals when a handler runs.
require('../src/entry/menu.js');

/** Installs fake SpreadsheetApp, HtmlService and LockService that record what happens, in order. */
function fakeApps() {
  const events = [];
  global.SpreadsheetApp = {
    getUi: () => ({
      showModelessDialog: (html, title) => events.push({ kind: 'dialog', html, title }),
      alert: (text) => events.push({ kind: 'alert', text })
    })
  };
  global.HtmlService = {
    createHtmlOutput: (content) => {
      const out = { content, width: 0, height: 0 };
      out.setWidth = (w) => { out.width = w; return out; };
      out.setHeight = (h) => { out.height = h; return out; };
      return out;
    }
  };
  global.LockService = {
    getScriptLock: () => ({ tryLock: () => true, releaseLock: () => events.push({ kind: 'release' }) })
  };
  return events;
}

/** Replaces namespace members for one test and returns a function that puts them back. */
function stub(parts) {
  const saved = [];
  Object.keys(parts).forEach((module) => {
    BS[module] = BS[module] || {};
    Object.keys(parts[module]).forEach((member) => {
      saved.push([module, member, BS[module][member]]);
      BS[module][member] = parts[module][member];
    });
  });
  return () => saved.reverse().forEach(([module, member, value]) => { BS[module][member] = value; });
}

const props = () => {
  const map = {};
  return { getProperty: (k) => (k in map ? map[k] : null), setProperty: (k, v) => { map[k] = v; }, getKeys: () => Object.keys(map) };
};

test('menu status: the summary opens in a modeless window, escaped, height capped for long text', () => {
  const events = fakeApps();
  const restore = stub({
    Container: { services: () => ({ props: props() }) },
    Gas: { triggers: () => ({ list: () => [] }) },
    ProviderRegistry: { statusLines: () => ['A <b> & "c"'].concat(Array.from({ length: 60 }, (_, i) => 'line ' + i)) }
  });
  try {
    BS.Menu.status();
  } finally {
    restore();
  }
  assert.deepStrictEqual(events.map((e) => e.kind), ['dialog']);
  const { html, title } = events[0];
  assert.strictEqual(title, 'Bank Sync: Status');
  assert.ok(html.content.includes('A &lt;b&gt; &amp; &quot;c&quot;'));
  assert.ok(!html.content.includes('<b>'));
  assert.deepStrictEqual([html.width, html.height], [560, 560]);
});

test('menu sync all and import: the window opens after the lock is released, never as an alert', () => {
  const events = fakeApps();
  const outcome = { results: [{ connection: 'csv-alex', status: 'ok', transactions: { added: 1, updated: 0, removed: 0, complete: true } }], raw: null, details: null, fx: null };
  const p = props();
  const restore = stub({
    Container: {
      services: () => ({ props: p }),
      syncAllDeps: () => ({ connections: [{ key: 'csv-alex', provider: 'csv' }] })
    },
    Gas: { dateTimeFormatter: () => () => '2026-10-25 10:00' },
    SyncAll: { run: () => outcome }
  });
  try {
    BS.Menu.syncAll();
    BS.Menu.importCsv();
  } finally {
    restore();
  }
  assert.deepStrictEqual(events.map((e) => e.kind), ['release', 'dialog', 'release', 'dialog']);
  assert.deepStrictEqual(events.filter((e) => e.kind === 'dialog').map((e) => e.title), ['Bank Sync: Sync all', 'Bank Sync: Import CSV files']);
  assert.match(events[1].html.content, /csv-alex: 1 added, 0 updated, 0 pending removed/);
  assert.ok(p.getProperty('BS_LAST_RUN'), 'a manual Sync all is recorded');
});
