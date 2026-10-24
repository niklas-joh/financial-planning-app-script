const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load');
const { fakeSheetBook } = require('./fake-sheet');

const NOW = Date.parse('2026-04-01T12:00:00Z');
const memProps = () => {
  const m = {};
  return { getProperty: (k) => (k in m ? m[k] : null), setProperty: (k, v) => { m[k] = v; } };
};
const acct = (o) => ({
  provider: 'enablebanking', accountId: 'a1', bank: 'Revolut', connection: 'revolut-niklas',
  accountName: 'Main', currency: 'EUR', ibanMasked: '****1111', accountType: 'CACC', product: '', ...o
});
const txn = (o) => ({
  provider: 'enablebanking', bank: '', connection: '', accountName: '', ibanMasked: '', accountId: 'a1',
  transactionId: 't1', bookingDate: '2026-03-30', valueDate: '', status: 'booked', amount: -1, currency: 'EUR',
  amountEur: null, fxRate: null, fxRateDate: '', counterparty: '', description: '', ...o
});

test('connections: config lists bank plus holder, labels read well', () => {
  const [first] = BS.Connections.all();
  assert.strictEqual(first.key, 'revolut-niklas');
  assert.strictEqual(BS.Connections.label(first), 'Revolut (niklas)');
  const keys = BS.Connections.all().map((c) => c.key);
  assert.strictEqual(new Set(keys).size, keys.length);
});

test('labels: account name falls back to the product, unknown accounts are left untouched', () => {
  const labelsFor = BS.AccountLabels.lookup([acct({}), acct({ accountId: 'a2', accountName: '', product: 'Savings' })]);
  assert.strictEqual(labelsFor('a2').accountName, 'Savings');
  assert.strictEqual(labelsFor('zzz'), null);
  const row = { accountId: 'zzz', bank: 'Other' };
  assert.strictEqual(BS.AccountLabels.stamp(row, labelsFor), row);
  assert.strictEqual(BS.AccountLabels.stamp({ accountId: 'a1' }, labelsFor).connection, 'revolut-niklas');
});

test('session store: migrateKey moves a legacy session once and never overwrites', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  store.saveSession('revolut', { sessionId: 'old' });
  assert.strictEqual(store.migrateKey('revolut', 'revolut-niklas'), true);
  assert.strictEqual(store.loadSession('revolut'), null);
  assert.strictEqual(store.loadSession('revolut-niklas').sessionId, 'old');
  assert.strictEqual(store.migrateKey('revolut', 'revolut-niklas'), false);
  store.saveSession('revolut', { sessionId: 'newer' });
  assert.strictEqual(store.migrateKey('revolut', 'revolut-niklas'), false);
  assert.strictEqual(store.loadSession('revolut-niklas').sessionId, 'old');
});

test('two people at one bank keep separate sessions and pending states', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  store.saveSession('revolut-niklas', { sessionId: 'n', accounts: [] });
  store.saveSession('revolut-partner', { sessionId: 'p', accounts: [] });
  assert.strictEqual(store.loadSession('revolut-niklas').sessionId, 'n');
  assert.strictEqual(store.loadSession('revolut-partner').sessionId, 'p');
});

test('auth.refreshAccounts re-reads details per account and keeps the session otherwise', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  const connection = BS.Config.CONNECTIONS.REVOLUT_NIKLAS;
  store.saveSession(connection.key, { sessionId: 's', accessValidUntil: '2027-01-01T00:00:00Z', accounts: [acct({ accountName: '' })] });
  const asked = [];
  const client = { getAccountDetails: (uid) => { asked.push(uid); return { uid, currency: 'EUR', details: 'Holiday pot' }; } };
  const auth = BS.EbAuth.create({ client, store, nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x' });
  assert.strictEqual(auth.refreshAccounts(connection), 1);
  assert.deepStrictEqual(asked, ['a1']);
  const saved = store.loadSession(connection.key);
  assert.strictEqual(saved.accounts[0].accountName, 'Holiday pot');
  assert.strictEqual(saved.sessionId, 's');
  assert.throws(() => auth.refreshAccounts({ key: 'revolut-partner' }), { name: 'AuthError' });
});

function runAll(over) {
  const connections = [
    { key: 'revolut-niklas', holder: 'niklas', bank: { aspspName: 'Revolut' } },
    { key: 'revolut-partner', holder: 'partner', bank: { aspspName: 'Revolut' } },
    { key: 'other-niklas', holder: 'niklas', bank: { aspspName: 'Other' } }
  ];
  let rows = [];
  const raw = {};
  const clock = { now: NOW };
  const deps = {
    connections,
    isConnected: (c) => c.key !== 'other-niklas',
    providerFor: (c) => ({
      listAccounts: () => {
        if (c.key === 'revolut-partner') throw BS.Errors.auth('Consent has expired; re-authorize');
        return [acct({ accountId: 'a-' + c.key, connection: c.key })];
      },
      getBalances: (id) => [{ accountId: id, amount: 1 }],
      getTransactions: (id) => [txn({ accountId: id, transactionId: 'x-' + id })]
    }),
    store: { load: () => rows, save: (r) => { rows = r; return r.length; } },
    rawWriter: { writeTable: (tab, cols, r) => { raw[tab] = r; return r.length; } },
    nowMs: () => clock.now,
    todayIso: () => '2026-04-01',
    ...over
  };
  return { deps, rows: () => rows, raw, clock };
}

test('sync all: one lapsed connection does not stop the others; statuses are reported', () => {
  const { deps, rows, raw } = runAll();
  const outcome = BS.SyncAll.run(deps);
  assert.deepStrictEqual(outcome.results.map((r) => [r.connection, r.status]),
    [['revolut-niklas', 'ok'], ['revolut-partner', 'error'], ['other-niklas', 'not connected']]);
  assert.match(outcome.results[1].error, /expired/);
  assert.deepStrictEqual(rows().map((t) => t.connection), ['revolut-niklas']);
  assert.strictEqual(rows()[0].accountName, 'Main');
  assert.strictEqual(raw.raw_accounts.length, 1);
  assert.strictEqual(raw.raw_balances[0].connection, 'revolut-niklas');
  assert.deepStrictEqual(outcome.raw, { accounts: 1, balances: 1 });
});

test('sync all: rows of a connection that is not synced keep their labels', () => {
  const { deps, rows } = runAll();
  const earlier = txn({ accountId: 'a-other', transactionId: 'old', bank: 'Other', connection: 'other-niklas', accountName: 'Pot', ibanMasked: '****9' });
  deps.store.save([earlier]);
  BS.SyncAll.run(deps);
  const kept = rows().find((t) => t.transactionId === 'old');
  assert.deepStrictEqual([kept.bank, kept.connection, kept.accountName, kept.ibanMasked], ['Other', 'other-niklas', 'Pot', '****9']);
});

test('sync all: skips connections once the shared time budget is nearly spent, writes no raw tabs if none ran', () => {
  const { deps, raw } = runAll();
  const real = deps.nowMs;
  // The first call (run start) sees the real time; every later call sees an almost spent budget.
  let first = true;
  deps.nowMs = () => { if (first) { first = false; return real(); } return real() + BS.Config.TRANSACTIONS.TIME_BUDGET_MS - 1000; };
  const outcome = BS.SyncAll.run(deps);
  assert.ok(outcome.results.filter((r) => r.status !== 'not connected').every((r) => r.status === 'skipped: time budget'));
  assert.strictEqual(outcome.raw, null);
  assert.deepStrictEqual(raw, {});
});

test('sync all: the summary has counts per connection and no row content', () => {
  const { deps } = runAll();
  const text = BS.SyncAll.summarize(BS.SyncAll.run(deps));
  assert.match(text, /revolut-niklas: 1 added, 0 updated, 0 pending removed/);
  assert.match(text, /revolut-partner: error \(AuthError: Consent has expired/);
  assert.match(text, /other-niklas: not connected/);
  assert.match(text, /Raw tabs: 1 accounts, 1 balances\./);
  assert.ok(!text.includes('x-a-'));
});

test('transactions tab gets the label columns next to each row', () => {
  const m = BS.TransactionsStore.toMatrix([txn({ bank: 'Revolut', connection: 'revolut-niklas', accountName: 'Main', ibanMasked: '****1' })]);
  const header = m[0];
  const row = m[1];
  ['bank', 'connection', 'accountName', 'ibanMasked'].forEach((k) => assert.ok(header.indexOf(k) !== -1 && header.indexOf(k) < header.indexOf('accountId')));
  assert.strictEqual(row[header.indexOf('connection')], 'revolut-niklas');
});

test('raw writer: refreshed connections are replaced, others kept, legacy rows dropped, never a whole-tab clear', () => {
  const book = fakeSheetBook();
  const writer = BS.RawWriter.create(book);
  const cols = BS.RawWriter.ACCOUNT_COLUMNS;
  writer.writeTable('raw_accounts', cols, [
    { provider: 'p', connection: 'revolut-niklas', accountId: 'n1', accountName: 'Old' },
    { provider: 'p', connection: 'revolut-partner', accountId: 'p1', accountName: 'Keep' },
    { provider: 'p', connection: 'revolut-partner', accountId: 'p2', accountName: 'Keep2' }]);
  book.sheets.raw_accounts.grid.push(['p', '', '', 'old-legacy', 'x', '', '', '', '']);
  const written = writer.writeTable('raw_accounts', cols,
    [{ provider: 'p', connection: 'revolut-niklas', accountId: 'n1', accountName: 'New' }], ['revolut-niklas']);
  assert.strictEqual(written, 3);
  const names = book.sheets.raw_accounts.grid.slice(1).map((r) => r[4]).filter(Boolean).sort();
  assert.deepStrictEqual(names, ['Keep', 'Keep2', 'New']);
  assert.ok(!book.sheets.raw_accounts.grid.some((r) => r.includes('old-legacy')));
});

test('raw writer: text columns are plain text (no formulas), the amount stays numeric', () => {
  const book = fakeSheetBook();
  BS.RawWriter.create(book).writeTable('raw_balances', BS.RawWriter.BALANCE_COLUMNS, [{ connection: 'c', amount: 5 }]);
  const cols = BS.RawWriter.BALANCE_COLUMNS.map((c) => c.key);
  const formatted = book.sheets.raw_balances.formats.map((f) => f[0]);
  assert.ok(formatted.includes(cols.indexOf('accountName') + 1));
  assert.ok(!formatted.includes(cols.indexOf('amount') + 1));
});

test('sync all: a connection that failed keeps its earlier raw rows on the next run', () => {
  const book = fakeSheetBook();
  const first = runAll({ rawWriter: BS.RawWriter.create(book) });
  BS.SyncAll.run(first.deps);
  const okProvider = (c) => ({
    listAccounts: () => [acct({ accountId: 'a-' + c.key, connection: c.key })],
    getBalances: (id) => [{ accountId: id, amount: 1 }], getTransactions: () => []
  });
  const failing = { listAccounts: () => { throw BS.Errors.auth('expired'); } };
  const second = runAll({
    rawWriter: BS.RawWriter.create(book),
    isConnected: () => true,
    providerFor: (c) => (c.key === 'revolut-niklas' ? failing : okProvider(c))
  });
  BS.SyncAll.run(second.deps);
  const kept = book.sheets.raw_accounts.grid.slice(1).map((r) => r[2]).filter(Boolean);
  assert.ok(kept.includes('revolut-niklas'), 'earlier row of the failed connection is still there');
  assert.ok(kept.includes('revolut-partner'));
});

test('auth.complete rejects a pending state saved before connections existed', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  store.savePending('old', { bankKey: 'revolut', validUntil: '2026-06-01T00:00:00Z' });
  const auth = BS.EbAuth.create({ client: { createSession: () => assert.fail('must not reach the bank') },
    store, nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x' });
  assert.throws(() => auth.complete('code', 'old'), /started before an update/);
  assert.strictEqual(store.loadSession('undefined'), null);
});

test('connections.choose accepts listed numbers only', () => {
  const list = [{ key: 'a' }, { key: 'b' }];
  assert.strictEqual(BS.Connections.choose(list, ' 2 ').key, 'b');
  ['0', '3', '', 'x', '1.5', '-1'].forEach((bad) => assert.throws(() => BS.Connections.choose(list, bad), { name: 'InputError' }));
});
