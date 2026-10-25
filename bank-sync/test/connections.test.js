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
  holderName: 'Holder X', accountName: 'Main', currency: 'EUR', ibanMasked: '****1111', accountType: 'CACC', product: '', ...o
});
const txn = (o) => ({
  provider: 'enablebanking', bank: '', connection: '', holderName: '', accountName: '', ibanMasked: '', accountId: 'a1',
  transactionId: 't1', bookingDate: '2026-03-30', valueDate: '', status: 'booked', amount: -1, currency: 'EUR',
  amountEur: null, fxRate: null, fxRateDate: '', counterparty: '', description: '', ...o
});

test('connections: key is bank plus holder, cleaned; empty holder is rejected', () => {
  const c = BS.Connections.build({ bankName: 'Danske Bank', country: 'se', holder: ' Niklas J. ' });
  assert.deepStrictEqual(c, { key: 'danske-bank-niklas-j', holder: 'niklas-j', bank: { aspspName: 'Danske Bank', country: 'SE' } });
  assert.strictEqual(BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' }).key, 'revolut-niklas');
  assert.strictEqual(BS.Connections.label(c), 'Danske Bank (niklas-j)');
  assert.throws(() => BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: '  ' }), { name: 'InputError' });
});

test('connections: store adds, replaces and lists in key order, per bank and holder', () => {
  const store = BS.Connections.create(memProps());
  store.save(BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'partner' }));
  store.save(BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' }));
  store.save(BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' }));
  assert.deepStrictEqual(store.all().map((c) => c.key), ['revolut-niklas', 'revolut-partner']);
  assert.strictEqual(store.find('revolut-partner').holder, 'partner');
  assert.strictEqual(store.find('nope'), null);
});

test('bank picker: exact name wins, one partial match resolves, several are listed, none is empty', () => {
  const names = ['Revolut', 'Revolut Business', 'Danske Bank', 'BGL BNP Paribas'];
  assert.strictEqual(BS.Connections.matchBanks(names, 'revolut').bank, 'Revolut');
  assert.strictEqual(BS.Connections.matchBanks(names, 'danske').bank, 'Danske Bank');
  const several = BS.Connections.matchBanks(names, 'rev');
  assert.strictEqual(several.bank, null);
  assert.deepStrictEqual(several.matches, ['Revolut', 'Revolut Business']);
  assert.deepStrictEqual(BS.Connections.matchBanks(names, 'zzz'), { bank: null, matches: [] });
  assert.deepStrictEqual(BS.Connections.matchBanks(names, '  '), { bank: null, matches: [] });
});

test('not-enabled errors are explained, other errors are not', () => {
  assert.match(BS.Connections.explainNotEnabled(BS.Errors.http(403, 'x', 'NO_ACCOUNTS_ADDED')), /Control Panel/);
  assert.match(BS.Connections.explainNotEnabled(BS.Errors.http(403, 'x', 'ACCESS_DENIED')), /Control Panel/);
  assert.strictEqual(BS.Connections.explainNotEnabled(BS.Errors.http(500, 'x', 'ASPSP_ERROR')), null);
  assert.strictEqual(BS.Connections.explainNotEnabled(new Error('x')), null);
});

test('http errors carry the API error code', () => {
  const http = BS.Http.create({
    fetch: () => ({ getResponseCode: () => 422, getContentText: () => JSON.stringify({ message: 'm', error: 'NO_ACCOUNTS_ADDED' }) }),
    sleep() {}, maxRetries: 0 });
  assert.throws(() => http.request({ method: 'GET', url: 'x' }), { name: 'HttpError', apiCode: 'NO_ACCOUNTS_ADDED' });
});

test('flatten: nested objects and arrays become dotted keys, IBANs are masked, null is empty', () => {
  const f = BS.Flatten.flat({ a: { b: 1 }, c: [{ d: 'x' }, 'GB82WEST12345698765432'], e: null, f: true, g: {}, h: [] });
  assert.deepStrictEqual(f, { 'a.b': 1, 'c[0].d': 'x', 'c[1]': '****5432', e: '', f: true });
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

function staleSetup(over) {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  const connection = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' });
  const saved = { sessionId: 's', accessValidUntil: '2027-01-01T00:00:00Z', accounts: [acct({ accountName: '', detailsAt: '' })] };
  store.saveSession(connection.key, saved);
  const asked = [];
  const client = { getAccountDetails: (uid) => { asked.push(uid); return { uid, currency: 'EUR', name: 'Holder X', details: 'Holiday pot', postal_address: { country: 'LU' } }; }, ...over };
  const auth = BS.EbAuth.create({ client, store, connections: BS.Connections.create(memProps()), nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x' });
  return { store, connection, asked, auth };
}

test('auth.refreshStale reads details for accounts never read, stores API fields apart from the session', () => {
  const { store, connection, asked, auth } = staleSetup();
  const accounts = auth.refreshStale(connection);
  assert.deepStrictEqual(asked, ['a1']);
  assert.strictEqual(accounts[0].accountName, 'Holiday pot');
  assert.ok(accounts[0].detailsAt);
  const saved = store.loadSession(connection.key);
  assert.strictEqual(saved.accounts[0].accountName, 'Holiday pot');
  assert.strictEqual(saved.accounts[0].apiFields, undefined);
  assert.strictEqual(store.loadApiFields('a1')['postal_address.country'], 'LU');
});

test('auth.refreshStale does nothing while details are fresh, and again after the max age', () => {
  const { store, connection, asked, auth } = staleSetup();
  auth.refreshStale(connection);
  asked.length = 0;
  auth.refreshStale(connection);
  assert.deepStrictEqual(asked, []);
  const old = store.loadSession(connection.key);
  old.accounts[0].detailsAt = new Date(NOW - (BS.Config.TRANSACTIONS.DETAILS_MAX_AGE_DAYS + 1) * 86400000).toISOString();
  store.saveSession(connection.key, old);
  auth.refreshStale(connection);
  assert.deepStrictEqual(asked, ['a1']);
});

test('auth.refreshStale keeps the old account when the bank call fails, and needs a session', () => {
  const { store, connection, auth } = staleSetup({ getAccountDetails: () => { throw BS.Errors.http(500, 'down'); } });
  const accounts = auth.refreshStale(connection);
  assert.strictEqual(accounts[0].accountName, '');
  assert.strictEqual(store.loadSession(connection.key).accounts[0].detailsAt, '');
  assert.throws(() => auth.refreshStale({ key: 'nobody' }), { name: 'AuthError' });
});

test('auth.complete stores the connection only when the bank granted accounts', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  const connections = BS.Connections.create(memProps());
  const connection = BS.Connections.build({ bankName: 'Danske Bank', country: 'SE', holder: 'niklas' });
  const grant = (accounts) => BS.EbAuth.create({ client: { createSession: () => ({ session_id: 's', accounts }) },
    store, connections, nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x' });
  store.savePending('empty', { connection, validUntil: '2027-01-01T00:00:00Z' });
  assert.throws(() => grant([]).complete('code', 'empty'), { name: 'AuthError', apiCode: 'NO_ACCOUNTS_ADDED' });
  assert.deepStrictEqual(connections.all(), []);
  assert.strictEqual(store.loadSession(connection.key), null);
  store.savePending('ok', { connection, validUntil: '2027-01-01T00:00:00Z' });
  grant([{ uid: 'd1', currency: 'SEK', name: 'Holder X', details: 'Everyday' }]).complete('code', 'ok');
  assert.deepStrictEqual(connections.all().map((c) => c.key), ['danske-bank-niklas']);
  assert.strictEqual(store.loadSession('danske-bank-niklas').accounts[0].accountName, 'Everyday');
  assert.strictEqual(store.loadApiFields('d1').details, 'Everyday');
});

test('session store: each connection has its own property, and the old single map still reads', () => {
  const props = memProps();
  props.setProperty('EB_SESSIONS', JSON.stringify({ revolut: { sessionId: 'legacy' } }));
  const store = BS.EbSessionStore.create(props, () => NOW);
  assert.strictEqual(store.loadSession('revolut').sessionId, 'legacy');
  assert.strictEqual(store.migrateKey('revolut', 'revolut-niklas'), true);
  assert.strictEqual(store.loadSession('revolut-niklas').sessionId, 'legacy');
  assert.strictEqual(store.loadSession('revolut'), null);
  store.saveSession('danske-bank-niklas', { sessionId: 'd' });
  assert.ok(props.getProperty('EB_SESSION_danske-bank-niklas'));
  assert.strictEqual(JSON.parse(props.getProperty('EB_SESSIONS')).revolut, undefined);
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
  const header = book.sheets.raw_accounts.grid[0];
  const row = (cells) => header.map((h) => cells[h] === undefined ? '' : cells[h]);
  book.sheets.raw_accounts.grid.push(row({ provider: 'p', accountId: 'old-legacy', accountName: 'x' }));
  const written = writer.writeTable('raw_accounts', cols,
    [{ provider: 'p', connection: 'revolut-niklas', accountId: 'n1', accountName: 'New' }], ['revolut-niklas']);
  assert.strictEqual(written, 3);
  const nameCol = header.indexOf('accountName');
  const names = book.sheets.raw_accounts.grid.slice(1).map((r) => r[nameCol]).filter(Boolean).sort();
  assert.deepStrictEqual(names, ['Keep', 'Keep2', 'New']);
  assert.ok(!book.sheets.raw_accounts.grid.some((r) => r.includes('old-legacy')));
});

test('raw writer: every API field becomes a column, kept rows keep theirs, formulas stay text', () => {
  const book = fakeSheetBook();
  const writer = BS.RawWriter.create(book);
  const cols = BS.RawWriter.ACCOUNT_COLUMNS;
  const mk = (conn, id, api) => BS.RawWriter.toRow({ provider: 'p', connection: conn, accountId: id, apiFields: api });
  writer.writeTable('raw_accounts', cols, [mk('c1', 'a1', { 'postal_address.country': 'LU', credit_limit: 5 })]);
  writer.writeTable('raw_accounts', cols, [mk('c2', 'a2', { usage: '=HYPERLINK("x")', legal_age: true })], ['c2']);
  const [header, ...rows] = book.sheets.raw_accounts.grid;
  ['api.credit_limit', 'api.legal_age', 'api.postal_address.country', 'api.usage'].forEach((h) => assert.ok(header.includes(h), h));
  const by = (id) => rows.find((r) => r[header.indexOf('accountId')] === id);
  assert.strictEqual(by('a1')[header.indexOf('api.postal_address.country')], 'LU');
  assert.strictEqual(by('a1')[header.indexOf('api.credit_limit')], 5);
  assert.strictEqual(by('a2')[header.indexOf('api.legal_age')], true);
  const formatted = book.sheets.raw_accounts.formats.map((f) => f[0]);
  assert.ok(formatted.includes(header.indexOf('api.usage') + 1), 'string column is plain text');
  assert.ok(!formatted.includes(header.indexOf('api.credit_limit') + 1), 'number column is not');
});

test('raw writer: the balance amount stays numeric, text columns are plain text', () => {
  const book = fakeSheetBook();
  BS.RawWriter.create(book).writeTable('raw_balances', BS.RawWriter.BALANCE_COLUMNS, [{ connection: 'c', accountName: 'n', amount: 5 }]);
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
