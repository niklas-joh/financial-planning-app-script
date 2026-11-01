// Data coverage: how each account is fetched (account_owners) and how far its data reaches (data_coverage).
// Every account, file and amount below is made up.
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');
const { fakeSheetBook } = require('./fake-sheet');

const DC = BS.DataCoverage;
const NOW = Date.parse('2026-11-01T08:00:00Z');

const fakeProps = (init) => {
  const map = Object.assign({}, init);
  return {
    map,
    getProperty: (k) => (k in map ? map[k] : null),
    setProperty: (k, v) => { map[k] = v; },
    getKeys: () => Object.keys(map)
  };
};

const bil = (accountLine, ...rows) => ['BILnet', 'LIST OF ACCOUNT MOVEMENTS', accountLine, '', 'Date;Value date;Name;Message;Amount']
  .concat(rows).join('\r\n') + '\r\n';
const file = (name, text, created) => ({ name, size: text.length, text: () => text, created });
const CURRENT = 'LU120000000000001234 Current Private';
const SAVINGS = 'LU450000000000009876 Savings Account';
const pay = (d) => d + ';' + d + ';Payment order;Standing order Remittance Info: Rent;-824.50';
const toIso = (d) => d.split('/').reverse().join('-');

/** Rows of a tab on the fake book, keyed by its header. */
const tabRows = (book, name) => {
  const grid = book.sheets[name].grid;
  return grid.slice(1).map((r) => Object.fromEntries(grid[0].map((h, i) => [h, r[i]])));
};

/** Sync all over one CSV folder with the owners and coverage tabs on a fake spreadsheet. */
function folderRun(files, opts = {}) {
  const book = opts.book || fakeSheetBook();
  const props = fakeProps({ CSV_FOLDER_ALEX: 'folder-1' });
  let rows = [];
  const today = opts.today || '2026-11-01';
  const env = { driveFolder: () => ({ files: () => files }) };
  const wired = { props, nowMs: () => NOW, connections: { all: () => [] }, store: { loadSession: () => null } };
  const store = { load: () => rows, save: (r) => { rows = r; return r.length; } };
  const deps = {
    connections: BS.ProviderRegistry.connections(wired),
    isConnected: (c) => BS.ProviderRegistry.isConnected(c, wired),
    providerFor: (c) => BS.ProviderRegistry.create(c, wired, env),
    store,
    rawWriter: { writeTable: (tab, cols, r) => r.length },
    nowMs: () => NOW,
    todayIso: () => today,
    owners: BS.AccountOwners.create(BS.AccountOwnersTab.create(book)),
    coverage: DC.create({
      tab: BS.DataCoverageTab.create(book, (d) => d.toISOString().slice(0, 10)),
      owners: BS.AccountOwnersTab.create(book),
      store,
      read: DC.readStore(props),
      csvCoverage: () => BS.CsvFolders.exportEnds(props),
      todayIso: () => today
    })
  };
  return { book, props, run: () => BS.SyncAll.run(deps), coverage: () => tabRows(book, 'data_coverage') };
}

test('coverage defaults: API accounts daily, CSV accounts a manual upload (Avanza quarterly, others monthly)', () => {
  assert.deepStrictEqual(
    [DC.defaults({ provider: 'enablebanking', accountId: 'x' }), DC.defaults({ provider: 'pluxee', accountId: 'y' })]
      .map((d) => [d.source, d.frequency]),
    [['API (automatic)', 'daily'], ['API (automatic)', 'daily']]);
  assert.deepStrictEqual([DC.defaults({ provider: 'csv', accountId: 'bil-1a2b' }).frequency,
    DC.defaults({ provider: 'csv', accountId: 'avanza-3c4d' }).frequency], ['monthly', 'quarterly']);
  assert.strictEqual(DC.defaults({ provider: 'csv', accountId: 'bil-1a2b' }).source, 'CSV (manual upload)');
  assert.ok(DC.defaults({ provider: 'trading212', accountId: '1' }).howTo.includes('report'));
});

test('coverage status: OK within one period plus a day, due up to two, overdue after; no schedule for ad hoc', () => {
  const today = '2026-11-01';
  assert.deepStrictEqual(['2026-10-31', '2026-10-30', '2026-10-29', '2026-10-28'].map((d) => DC.statusOf('daily', d, today)),
    ['OK', 'OK', 'due', 'overdue']);
  assert.deepStrictEqual(['2026-09-30', '2026-09-29', '2026-08-30', '2026-08-29'].map((d) => DC.statusOf('Monthly', d, today)),
    ['OK', 'due', 'due', 'overdue']);
  assert.deepStrictEqual([DC.statusOf('ad hoc', '2020-01-01', today), DC.statusOf('', '', today), DC.statusOf('monthly', '', today)],
    ['no schedule', 'no schedule', 'no data yet']);
  assert.deepStrictEqual([DC.frequencyDays(' Quarterly '), DC.frequencyDays('whenever')], [92, null]);
});

test('csv export ends: files of one source uploaded the same day are one export; a newer upload moves only its account', () => {
  const ends = BS.CsvImportProvider.exportEnds([
    { source: 'bil', day: '2026-10-25', last: '2026-08-31', accountIds: ['bil-current'] },
    { source: 'bil', day: '2026-10-25', last: '2026-03-20', accountIds: ['bil-savings'] },
    { source: 'avanza', day: '2026-10-25', last: '2026-10-22', accountIds: ['avanza-isk'] },
    { source: 'bil', day: '2026-11-15', last: '2026-11-14', accountIds: ['bil-current'] }
  ]);
  assert.deepStrictEqual(ends, { 'avanza-isk': '2026-10-22', 'bil-current': '2026-11-14', 'bil-savings': '2026-08-31' });
});

test('coverage: a sync lists every account with its dates, export start and status; a quiet account takes its export\'s end', () => {
  const r = folderRun([
    file('current.csv', bil(CURRENT, pay('28/08/2026'), pay('05/04/2025')), '2026-10-25T09:00:00Z'),
    file('savings.csv', bil(SAVINGS, pay('20/03/2026')), '2026-10-25T09:01:00Z')
  ]);
  const outcome = r.run();
  const rows = r.coverage();
  assert.deepStrictEqual(rows.map((x) => [x.accountName, x.source, x.frequency, x.rows, x.firstDate, x.lastDate, x.upTo, x.exportFrom, x.status]), [
    ['Current Private', 'CSV (manual upload)', 'monthly', 2, '2025-04-05', '2026-08-28', '2026-08-28', '2026-08-28', 'overdue'],
    ['Savings Account', 'CSV (manual upload)', 'monthly', 1, '2026-03-20', '2026-03-20', '2026-08-28', '2026-08-28', 'overdue']
  ]);
  assert.strictEqual(rows[0].nextStep, 'Export 2026-08-28 to today and upload it to the CSV folder');
  assert.match(BS.SyncAll.summarize(outcome),
    /Data coverage: 2 rows on data_coverage; to update: BIL Current Private \(export 2026-08-28 to today\), BIL Savings Account/);
  // account_owners got how each account is fetched.
  const owners = tabRows(r.book, 'account_owners');
  assert.deepStrictEqual(owners.map((o) => [o.source, o.frequency]), [['CSV (manual upload)', 'monthly'], ['CSV (manual upload)', 'monthly']]);
  assert.ok(owners[0].howTo.includes('CSV folder'));
});

test('coverage: a frequency typed on account_owners wins, and a fresh export makes the account OK', () => {
  const book = fakeSheetBook();
  folderRun([file('current.csv', bil(CURRENT, pay('28/08/2026')), '2026-10-25T09:00:00Z')], { book }).run();
  const grid = book.sheets.account_owners.grid;
  grid[1][grid[0].indexOf('frequency')] = 'ad hoc';
  const r = folderRun([
    file('current.csv', bil(CURRENT, pay('28/08/2026')), '2026-10-25T09:00:00Z'),
    file('current-nov.csv', bil(CURRENT, pay('30/10/2026')), '2026-11-01T07:00:00Z')
  ], { book });
  r.run();
  const row = r.coverage()[0];
  assert.deepStrictEqual([row.frequency, row.upTo, row.status, row.nextStep], ['ad hoc', '2026-10-30', 'no schedule', '']);
  grid[1][grid[0].indexOf('frequency')] = 'monthly';
  r.run();
  assert.strictEqual(r.coverage()[0].status, 'OK');
});

test('coverage: manual_balances items are listed by their latest value; one without an amount has no data yet', () => {
  const book = fakeSheetBook();
  const manual = book.insertSheet('manual_balances');
  manual.grid.push(['date', 'bank', 'accountName', 'owner', 'kind', 'amount', 'currency', 'note', 'frequency']);
  manual.grid.push([new Date('2026-09-30T00:00:00Z'), 'Bank', 'Loan', 'Common', 'Loan', -1000, 'EUR', '', '']);
  manual.grid.push([new Date('2026-10-31T00:00:00Z'), 'Bank', 'Loan', 'Common', 'Loan', -900, 'EUR', '', 'quarterly']);
  manual.grid.push(['2026-11-01', 'Property', 'Flat', 'common', 'Property', '', 'EUR', 'not valued yet', '']);
  const r = folderRun([], { book });
  const outcome = r.run();
  assert.deepStrictEqual(r.coverage().map((x) => [x.owner, x.accountName, x.source, x.frequency, x.rows, x.firstDate, x.upTo, x.status]), [
    ['common', 'Loan', 'Manual entry', 'quarterly', 2, '2026-09-30', '2026-10-31', 'OK'],
    ['common', 'Flat', 'Manual entry', 'monthly', 0, '', '', 'no data yet']
  ]);
  assert.match(r.coverage()[1].nextStep, /manual_balances/);
  assert.match(BS.DataCoverage.summaryLine(outcome.coverage), /to update: Property Flat \(no data yet\)/);
});

test('coverage: an API account is up to its last good read; when reads stop it becomes due, then overdue', () => {
  const book = fakeSheetBook();
  BS.AccountOwnersTab.create(book).append([{ owner: 'alex', bank: 'Demo Bank', accountName: 'Main', provider: 'enablebanking', accountId: 'eb-1' }]);
  const props = fakeProps({});
  let today = '2026-10-20';
  const coverage = DC.create({
    tab: BS.DataCoverageTab.create(book, (d) => d.toISOString().slice(0, 10)),
    owners: BS.AccountOwnersTab.create(book),
    store: { load: () => [{ accountId: 'eb-1', bookingDate: '2026-10-19' }] },
    read: DC.readStore(props),
    csvCoverage: () => ({}),
    todayIso: () => today
  });
  coverage.update([{ provider: 'enablebanking', accountId: 'eb-1' }]);
  const at = (d) => { today = d; coverage.update([]); return tabRows(book, 'data_coverage')[0]; };
  assert.deepStrictEqual([at('2026-10-22').status, at('2026-10-23').status, at('2026-10-24').status], ['OK', 'due', 'overdue']);
  const row = tabRows(book, 'data_coverage')[0];
  assert.deepStrictEqual([row.source, row.frequency, row.upTo, row.exportFrom], ['API (automatic)', 'daily', '2026-10-20', '']);
  assert.strictEqual(row.nextStep, 'Not read since 2026-10-20: check Show status (consent or login)');
  assert.strictEqual(JSON.parse(props.map.BS_ACCOUNTS_READ)['eb-1'], '2026-10-20');
});

test('coverage tab: rewritten from the top, leftover rows below cleared (never the whole tab); Show status lists what is due', () => {
  const book = fakeSheetBook();
  const tab = BS.DataCoverageTab.create(book, (d) => d.toISOString().slice(0, 10));
  const row = (name, status) => ({ bank: 'BIL', accountName: name, status, exportFrom: status === 'OK' ? '' : '2026-08-31' });
  tab.write([row('A', 'due'), row('B', 'OK'), row('C', 'overdue')]);
  assert.strictEqual(DC.statusLine(tab),
    'Data coverage: 3 rows on data_coverage; to update: BIL A (export 2026-08-31 to today), BIL C (export 2026-08-31 to today)');
  tab.write([row('B', 'OK')]);
  assert.strictEqual(book.sheets.data_coverage.grid.length, 2);
  assert.strictEqual(DC.statusLine(tab), 'Data coverage: 1 rows on data_coverage; all up to date');
  assert.strictEqual(DC.statusLine(BS.DataCoverageTab.create(fakeSheetBook(), String)),
    'Data coverage: none yet (the next sync fills data_coverage)');
});

test('coverage: a failure is reported in the summary and never stops the sync', () => {
  const coverage = DC.create({
    tab: { write: () => {}, loadManual: () => [] },
    owners: { load: () => { throw new Error('header renamed'); } },
    store: { load: () => [] },
    read: { load: () => ({}), save: () => {} },
    csvCoverage: () => ({}),
    todayIso: () => '2026-11-01'
  });
  const outcome = coverage.update([]);
  assert.deepStrictEqual([outcome.rows, outcome.due], [0, '']);
  assert.match(DC.summaryLine(outcome), /^Data coverage: data_coverage could not be updated \(Error: header renamed\)\.$/);
});

test('account_owners: a tab in the old layout gains the new columns on the next append; typed cells stay', () => {
  const book = fakeSheetBook();
  const sheet = book.insertSheet('account_owners');
  sheet.grid.push(['owner', 'bank', 'accountName', 'ibanMasked', 'connection', 'provider', 'accountId', 'notes']);
  sheet.grid.push(['alex', 'BIL', 'Current', '', 'csv-alex', 'csv', 'bil-1', 'my note']);
  BS.AccountOwnersTab.create(book).append([{ owner: '', bank: 'BIL', accountName: 'Savings', provider: 'csv', accountId: 'bil-2',
    source: 'CSV (manual upload)', frequency: 'monthly', howTo: 'upload' }]);
  assert.deepStrictEqual(sheet.grid[0].slice(7), ['notes', 'source', 'frequency', 'howTo']);
  assert.deepStrictEqual(sheet.grid[1], ['alex', 'BIL', 'Current', '', 'csv-alex', 'csv', 'bil-1', 'my note']);
  assert.deepStrictEqual(sheet.grid[2].slice(6), ['bil-2', '', 'CSV (manual upload)', 'monthly', 'upload']);
});

test('csv import record: keeps each account\'s export end, in the local date of the upload', () => {
  const saved = [];
  const provider = BS.CsvImportProvider.create({
    connection: BS.CsvFolders.connections(fakeProps({ CSV_FOLDER_ALEX: 'f' }))[0],
    folder: { files: () => [file('a.csv', bil(CURRENT, pay('28/08/2026')), '2026-10-24T23:30:00Z')] },
    record: { save: (r) => saved.push(r) },
    nowMs: () => NOW,
    toLocalDate: (iso) => (iso.startsWith('2026-10-24T23') ? '2026-10-25' : iso.slice(0, 10))
  });
  const accounts = provider.listAccounts();
  assert.deepStrictEqual(saved[0].accounts, { [accounts[0].accountId]: toIso('28/08/2026') });
});
