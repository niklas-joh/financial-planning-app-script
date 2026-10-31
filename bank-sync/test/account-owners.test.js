// Account owners: the account_owners tab, the owner label on every row and a full Sync all over one CSV folder
// that holds several people's exports. Every file below is made up (fake accounts and people).
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');
const { fakeSheetBook } = require('./fake-sheet');

const NOW = Date.parse('2026-09-01T08:00:00Z');
const TAB = 'account_owners';

const fakeProps = (init) => {
  const map = Object.assign({}, init);
  return {
    getProperty: (k) => (k in map ? map[k] : null),
    setProperty: (k, v) => { map[k] = v; },
    getKeys: () => Object.keys(map)
  };
};

const bil = (accountLine, ...rows) => ['BILnet', 'LIST OF ACCOUNT MOVEMENTS', accountLine, '', 'Date;Value date;Name;Message;Amount']
  .concat(rows).join('\r\n') + '\r\n';
const file = (name, text) => ({ name, size: text.length, text: () => text });
const FILES = [
  file('mine.csv', bil('LU120000000000001234 Current Private', '28/08/2026;28/08/2026;Payment order;Standing order Remittance Info: Rent;-824.50')),
  file('hers.csv', bil('LU450000000000009876 Savings Account', '27/08/2026;27/08/2026;Credit;Credit Remittance Info: Gift;50.00'))
];

/** Sync all over one CSV folder, with the owners tab on a fake spreadsheet. */
function folderRun(files) {
  const book = fakeSheetBook();
  const props = fakeProps({ CSV_FOLDER_NIKLAS: 'folder-1' });
  let rows = [];
  const raw = {};
  const env = { driveFolder: () => ({ files: () => files }) };
  const wired = { props, nowMs: () => NOW, connections: { all: () => [] }, store: { loadSession: () => null } };
  const deps = () => ({
    connections: BS.ProviderRegistry.connections(wired),
    isConnected: (c) => BS.ProviderRegistry.isConnected(c, wired),
    providerFor: (c) => BS.ProviderRegistry.create(c, wired, env),
    store: { load: () => rows, save: (r) => { rows = r; return r.length; } },
    rawWriter: { writeTable: (tab, cols, r) => { raw[tab] = r; return r.length; } },
    nowMs: () => NOW,
    todayIso: () => '2026-09-01',
    owners: BS.AccountOwners.create(BS.AccountOwnersTab.create(book))
  });
  const grid = () => book.sheets[TAB].grid;
  /** Types an owner on the tab, as the owner would. */
  const setOwner = (accountName, owner) => {
    const header = grid()[0];
    const row = grid().find((r) => r[header.indexOf('accountName')] === accountName);
    row[header.indexOf('owner')] = owner;
  };
  return { book, raw, grid, setOwner, rows: () => rows, run: () => BS.SyncAll.run(deps()) };
}

const ownerOf = (rows, accountName) => rows.filter((t) => t.accountName === accountName).map((t) => t.owner);

test('owners: a first sync lists each CSV account on account_owners with no owner, and says so', () => {
  const r = folderRun(FILES);
  const outcome = r.run();
  assert.deepStrictEqual(r.grid()[0], ['owner', 'bank', 'accountName', 'ibanMasked', 'connection', 'provider', 'accountId']);
  assert.deepStrictEqual(r.grid().slice(1).map((row) => row.slice(0, 6)), [
    ['', 'BIL', 'Current Private', 'LU120000000000001234', 'csv-niklas', 'csv'],
    ['', 'BIL', 'Savings Account', 'LU450000000000009876', 'csv-niklas', 'csv']
  ]);
  assert.strictEqual(r.book.sheets[TAB].frozen, 1);
  assert.deepStrictEqual(outcome.owners, { added: 2, missing: ['BIL Current Private', 'BIL Savings Account'] });
  assert.ok(r.rows().every((t) => t.owner === ''));
  const summary = BS.SyncAll.summarize(outcome);
  assert.match(summary, /Account owners: 2 accounts added to account_owners; no owner yet for BIL Current Private, BIL Savings Account/);
});

test('owners: owners typed on the tab reach every row at the next sync, as slugs; nothing is added twice', () => {
  const r = folderRun(FILES);
  r.run();
  r.setOwner('Current Private', 'Niklas');
  r.setOwner('Savings Account', ' Marie-Aude ');
  const second = r.run();
  assert.deepStrictEqual(ownerOf(r.rows(), 'Current Private'), ['niklas']);
  assert.deepStrictEqual(ownerOf(r.rows(), 'Savings Account'), ['marie-aude']);
  assert.deepStrictEqual(second.results.map((x) => [x.transactions.added, x.transactions.updated]), [[0, 0]]);
  assert.deepStrictEqual(second.owners, { added: 0, missing: [] });
  assert.doesNotMatch(BS.SyncAll.summarize(second), /Account owners/);
  assert.strictEqual(r.grid().length, 3);
  assert.deepStrictEqual(r.raw.raw_accounts.map((a) => a.owner).sort(), ['marie-aude', 'niklas']);
  assert.deepStrictEqual(r.raw.raw_transactions.map((t) => t.owner).sort(), ['marie-aude', 'niklas']);
  // Changing an owner later moves the account's rows with it.
  r.setOwner('Savings Account', 'Common');
  r.run();
  assert.deepStrictEqual(ownerOf(r.rows(), 'Savings Account'), ['common']);
});

test('owners: a new export of another account is appended below the rows already typed', () => {
  const r = folderRun(FILES.slice(0, 1));
  r.run();
  r.setOwner('Current Private', 'Niklas');
  r.grid()[0].push('notes');
  r.grid()[1].push('my main account');
  const files = FILES.slice();
  const run = folderRun(files);
  run.book.sheets[TAB] = r.book.sheets[TAB];
  const outcome = run.run();
  assert.deepStrictEqual(outcome.owners, { added: 1, missing: ['BIL Savings Account'] });
  assert.deepStrictEqual(run.grid()[1].slice(0, 3), ['Niklas', 'BIL', 'Current Private']);
  assert.strictEqual(run.grid()[1][7], 'my main account');
  assert.deepStrictEqual(run.grid()[2].slice(0, 3), ['', 'BIL', 'Savings Account']);
  assert.strictEqual(run.grid()[2][7], '');
  assert.deepStrictEqual(ownerOf(run.rows(), 'Current Private'), ['niklas']);
});

test('owners: a login or API key starts with its holder; a CSV folder with no owner', () => {
  assert.strictEqual(BS.AccountOwners.defaultOwner({ holder: 'niklas', bank: {} }), 'niklas');
  assert.strictEqual(BS.AccountOwners.defaultOwner({ holder: 'niklas-isa', provider: 'trading212' }), 'niklas-isa');
  assert.strictEqual(BS.AccountOwners.defaultOwner({ holder: 'niklas', provider: 'csv' }), '');
  const added = [];
  const owners = BS.AccountOwners.create({ load: () => [], append: (rows) => added.push(...rows) });
  const accounts = owners.assign([{ provider: 'enablebanking', accountId: 'u1', bank: 'Revolut', connection: 'revolut-niklas', accountName: '', currency: 'EUR', accountType: 'CACC' }],
    { holder: 'niklas', bank: { aspspName: 'Revolut' } });
  assert.strictEqual(accounts[0].owner, 'niklas');
  assert.deepStrictEqual(owners.save(), { added: 1, missing: [] });
  assert.deepStrictEqual([added[0].owner, added[0].accountName], ['niklas', 'Current account EUR']);
});

test('owners: a tab that cannot be read keeps the stored owners, adds nothing and is reported', () => {
  const book = fakeSheetBook();
  book.insertSheet(TAB).grid = [['who', 'account'], ['niklas', 'bil-1']];
  const tab = BS.AccountOwnersTab.create(book);
  const owners = BS.AccountOwners.create(tab);
  const account = { provider: 'csv', accountId: 'bil-1', bank: 'BIL', connection: 'csv-niklas', accountName: 'Current' };
  const assigned = owners.assign([account], { holder: 'niklas', provider: 'csv' });
  assert.ok(!('owner' in assigned[0]));
  const stored = { accountId: 'bil-1', owner: 'common' };
  assert.strictEqual(BS.AccountLabels.stamp(stored, BS.AccountLabels.lookup(assigned)).owner, 'common');
  const outcome = owners.save();
  assert.strictEqual(outcome.added, 0);
  assert.match(outcome.error, /needs an "owner" and an "accountId" column/);
  assert.match(BS.AccountOwners.summaryLine(outcome), /account_owners could not be used .*owners left as they were/);
  assert.deepStrictEqual(book.sheets[TAB].grid.length, 2);
  assert.match(BS.AccountOwners.statusLine(tab), /cannot be read/);
});

test('owners: Show status counts the accounts and names the ones without an owner', () => {
  const book = fakeSheetBook();
  const tab = BS.AccountOwnersTab.create(book);
  assert.match(BS.AccountOwners.statusLine(tab), /none listed yet/);
  tab.append([
    { owner: 'niklas', bank: 'BIL', accountName: 'Current', accountId: 'bil-1' },
    { owner: '', bank: 'Avanza', accountName: 'ISK', accountId: 'avanza-2' }
  ]);
  assert.strictEqual(BS.AccountOwners.statusLine(tab), 'Account owners: 2 accounts on account_owners; no owner yet: Avanza ISK');
  book.sheets[TAB].grid[2][0] = 'Common';
  assert.strictEqual(BS.AccountOwners.statusLine(tab), 'Account owners: 2 accounts on account_owners, all with an owner');
  assert.ok(book.sheets[TAB].formats.length > 0);
});
