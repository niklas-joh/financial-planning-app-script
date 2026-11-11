// The summary window: Sync all's outcome as a view (headline, connections table, steps, accounts to update)
// and its HTML. Synthetic names only.
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');

const ok = (connection, added, extra) => ({
  connection, status: 'ok', transactions: Object.assign({ added, updated: 0, removed: 0, complete: true }, extra)
});
const fx = (extra) => Object.assign({ changed: 1, missing: 0, waiting: 1, balancesMissing: 0, holdingsMissing: 0, failed: [] }, extra);
const due = (n) => Array.from({ length: n }, (_, i) => ({ name: 'Testbank Account ' + i, action: 'export 2026-08-31 to today' }));

test('sync view: a clean run is green, counts new transactions and lists each step as a line', () => {
  const view = BS.SyncAll.view({
    results: [ok('testbank-alex', 0), ok('csv-alex', 2)],
    raw: { accounts: 3, balances: 2, transactions: 40 }, details: { enriched: 0, read: 5 }, fx: fx(),
    coverage: { rows: 4, due: '', dueItems: [] }
  });
  assert.deepStrictEqual(view.headline, { tone: 'ok', text: 'Sync finished: 2 new transactions' });
  assert.deepStrictEqual(view.table.rows, [
    { name: 'testbank-alex', counts: [0, 0, 0], note: '' }, { name: 'csv-alex', counts: [2, 0, 0], note: '' }]);
  assert.deepStrictEqual(view.lines.map((l) => l.text), [
    'Raw tabs: 3 accounts, 2 balances, 40 transactions.',
    'Details from reports: 0 transactions updated (5 read).',
    "EUR conversion: 1 transactions converted, 1 waiting for today's rate.",
    'Data coverage: 4 rows on data_coverage; all up to date.']);
  assert.strictEqual(view.list, null);
});

test('sync view: accounts to update become a list (capped) and turn the headline yellow', () => {
  const max = BS.Config.COVERAGE.MAX_LISTED;
  const view = BS.SyncAll.view({
    results: [ok('testbank-alex', 1)], raw: null, details: null, fx: null,
    coverage: { rows: 20, due: 'x', dueItems: due(max + 2) }
  });
  assert.deepStrictEqual(view.headline, { tone: 'warn', text: 'Sync finished: 1 new transaction, ' + (max + 2) + ' accounts to update' });
  assert.strictEqual(view.list.title, (max + 2) + ' accounts need an update (see data_coverage)');
  assert.strictEqual(view.list.items.length, max);
  assert.strictEqual(view.list.more, 2);
  assert.ok(!view.lines.some((l) => /Data coverage/.test(l.text)), 'the list replaces the coverage sentence');
});

test('sync view: failures are red (connection, step); an unfinished connection asks to run again', () => {
  const failed = BS.SyncAll.view({
    results: [ok('testbank-alex', 1), { connection: 'otherbank-alex', status: 'error', error: 'HttpError: 500' }],
    raw: null, details: null, fx: fx({ failed: ['SEK (HttpError: 503)'] }), coverage: { rows: 0, due: '', error: 'Error: boom' }
  });
  assert.deepStrictEqual(failed.headline, { tone: 'error', text: 'Sync finished with problems: 1 new transaction, 1 connection not read' });
  assert.deepStrictEqual(failed.table.rows[1], { name: 'otherbank-alex', error: 'error (HttpError: 500)' });
  assert.deepStrictEqual(failed.lines.filter((l) => l.error).map((l) => l.text), [
    'Rates unavailable for SEK (HttpError: 503)', 'Data coverage: data_coverage could not be updated (Error: boom).']);

  const unfinished = BS.SyncAll.view({
    results: [ok('testbank-alex', 3, { complete: false, unavailable: 2 })], raw: null, details: null, fx: null
  });
  assert.strictEqual(unfinished.headline.tone, 'warn');
  assert.match(unfinished.headline.text, /^Sync not finished, run it again to continue: 3 new transactions$/);
  assert.strictEqual(unfinished.table.rows[0].note, '2 date windows not served by the source; time budget reached; run again to continue');
});

test('sync view: the plain-text summary names the same things', () => {
  const text = BS.SyncAll.summarize({
    results: [ok('testbank-alex', 1, { complete: false, unavailable: 2 })], raw: null, details: null, fx: fx(),
    coverage: { rows: 2, due: 'Testbank Account 0 (export 2026-08-31 to today)', dueItems: due(1) }
  });
  assert.strictEqual(text, [
    'testbank-alex: 1 added, 0 updated, 0 pending removed, 2 date windows not served by the source, time budget reached; run again to continue',
    "EUR conversion: 1 transactions converted, 1 waiting for today's rate.",
    'Data coverage: 2 rows on data_coverage; to update: Testbank Account 0 (export 2026-08-31 to today).'].join('\n'));
});

test('summary html: everything is escaped; labels and values are cells; zero rows are greyed; errors are red', () => {
  const page = BS.SummaryView.render({
    headline: { tone: 'error', text: 'Bad <run>' },
    table: { title: 'Connections', columns: ['Connection', 'Added'], rows: [
      { name: 'a<b', counts: [0], note: '' }, { name: 'c', counts: [4], note: 'note & more' }, { name: 'd', error: 'error (<x>)' }] },
    lines: [{ text: 'Raw tabs: 3 accounts.' }, { text: 'EUR conversion failed (<y>)', error: true }],
    list: { title: '1 account needs an update', items: [{ name: 'Bank "A"', action: 'no data yet' }], more: 3 },
    extra: ['', '- row | with: colon', 'Still masked: none. Every stored row holds full account numbers.']
  });
  const html = page.html;
  assert.ok(!/<run>|<b|<x>|<y>|"A"/.test(html.replace(/<\/?(div|table|tr|td)[^>]*>/g, '')), 'no raw text reaches the page');
  assert.match(html, /background:#fce8e6[^>]*>Bad &lt;run&gt;</);
  assert.match(html, /color:#80868b"><td[^>]*>a&lt;b</);
  assert.match(html, /font-weight:bold[^>]*>4</);
  assert.match(html, /note &amp; more/);
  assert.match(html, /color:#c5221f"><td[^>]*>d<\/td><td colspan="1"[^>]*>error \(&lt;x&gt;\)</);
  assert.match(html, />Raw tabs<\/td><td[^>]*>3 accounts</);
  assert.match(html, /color:#c5221f;">EUR conversion failed \(&lt;y&gt;\)</);
  assert.match(html, />Bank &quot;A&quot;<\/td><td[^>]*>no data yet</);
  assert.match(html, />and 3 more</);
  assert.match(html, /colspan="2"[^>]*>- row \| with: colon</);
  assert.match(html, />Still masked<\/td><td[^>]*>none\. Every stored row holds full account numbers\.</);
  assert.strictEqual(page.width, 560);
});

test('summary html: the window grows with the content up to a cap', () => {
  const small = BS.SummaryView.render(BS.SummaryView.fromText('one\ntwo'));
  const big = BS.SummaryView.render(BS.SummaryView.fromText(Array.from({ length: 80 }, (_, i) => 'line ' + i).join('\n')));
  assert.ok(small.height < 120 && small.height > 60);
  assert.strictEqual(big.height, 560);
});
