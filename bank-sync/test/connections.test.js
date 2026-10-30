const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load');
const { withMasking } = require('./masking');
const { fakeSheetBook } = require('./fake-sheet');

const NOW = Date.parse('2026-04-01T12:00:00Z');
const memProps = () => {
  const m = {};
  return { getProperty: (k) => (k in m ? m[k] : null), setProperty: (k, v) => { m[k] = v; }, deleteProperty: (k) => { delete m[k]; } };
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

test('flatten: nested objects and arrays become dotted keys, IBANs are masked, null is empty', () => withMasking(BS, true, () => {
  const f = BS.Flatten.flat({ a: { b: 1 }, c: [{ d: 'x' }, 'GB82WEST12345698765432'], e: null, f: true, g: {}, h: [] });
  assert.deepStrictEqual(f, { 'a.b': 1, 'c[0].d': 'x', 'c[1]': '****5432', e: '', f: true });
}));

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
  // The fake provider returns one row per fetch window and the 90-day first run has three windows.
  assert.deepStrictEqual(outcome.raw, { accounts: 1, balances: 1, transactions: 3 });
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
  assert.match(text, /Raw tabs: 1 accounts, 1 balances, 3 transactions\./);
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

test('flatten: account and card numbers are masked by key, the holder name is kept', () => withMasking(BS, true, () => {
  const f = BS.Flatten.flat({
    name: 'Jane Holder', account_id: { other: { identification: '12-34-56 98765432', scheme_name: 'BBAN' } },
    all_account_ids: [{ identification: '12345678' }], card_number: '5500000000000004', sort_code: '123456'
  });
  assert.strictEqual(f.name, 'Jane Holder');
  assert.strictEqual(f['account_id.other.identification'], '****5432');
  assert.strictEqual(f['all_account_ids[0].identification'], '****5678');
  assert.strictEqual(f.card_number, '****0004');
  assert.strictEqual(f.sort_code, '****3456');
  assert.strictEqual(f['account_id.other.scheme_name'], 'BBAN');
  assert.ok(!JSON.stringify(f).includes('98765432'));
}));

test('slug: accents are dropped so names like Asa still work', () => {
  assert.strictEqual(BS.Connections.slug('Åsa Öberg'), 'asa-oberg');
  assert.strictEqual(BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'Åsa' }).key, 'revolut-asa');
});

test('a freshly connected login has its details read on the first sync (detailsAt starts empty)', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  const connections = BS.Connections.create(memProps());
  const connection = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' });
  const asked = [];
  const client = {
    createSession: () => ({ session_id: 's', accounts: [{ uid: 'n1', currency: 'EUR' }] }),
    getAccountDetails: (uid) => { asked.push(uid); return { uid, currency: 'EUR', details: 'Main' }; }
  };
  const auth = BS.EbAuth.create({ client, store, connections, nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x' });
  store.savePending('p', { connection, validUntil: '2027-01-01T00:00:00Z' });
  auth.complete('code', 'p');
  assert.strictEqual(store.loadSession(connection.key).accounts[0].detailsAt, '');
  assert.strictEqual(auth.refreshStale(connection)[0].accountName, 'Main');
  assert.deepStrictEqual(asked, ['n1']);
});

test('refreshStale re-throws consent problems (401/403) but tolerates transient errors', () => {
  [401, 403].forEach((status) => {
    const { connection, auth } = staleSetup({ getAccountDetails: () => { throw BS.Errors.http(status, 'no'); } });
    assert.throws(() => auth.refreshStale(connection), { name: 'HttpError', status });
  });
  [429, 500].forEach((status) => {
    const { connection, auth } = staleSetup({ getAccountDetails: () => { throw BS.Errors.http(status, 'later'); } });
    assert.doesNotThrow(() => auth.refreshStale(connection));
  });
});

test('API fields too large for one property are skipped without failing', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  assert.strictEqual(store.saveApiFields('a1', { k: 'x'.repeat(9000) }), false);
  assert.deepStrictEqual(store.loadApiFields('a1'), {});
  assert.strictEqual(store.saveApiFields('a1', { k: 'small' }), true);
  assert.strictEqual(store.loadApiFields('a1').k, 'small');
});

test('sync all: a connection that fails while reading balances writes nothing for it and keeps its raw rows', () => {
  const book = fakeSheetBook();
  const ok = (c) => ({ listAccounts: () => [acct({ accountId: 'a-' + c.key, connection: c.key })], getBalances: (id) => [{ accountId: id, amount: 1 }], getTransactions: () => [] });
  const first = runAll({ rawWriter: BS.RawWriter.create(book), isConnected: () => true, providerFor: ok });
  first.deps.connections = first.deps.connections.slice(0, 2);
  BS.SyncAll.run(first.deps);
  const second = runAll({
    rawWriter: BS.RawWriter.create(book), isConnected: () => true,
    providerFor: (c) => (c.key === 'revolut-partner'
      ? { listAccounts: () => [acct({ accountId: 'a-x', connection: c.key })], getBalances: () => { throw BS.Errors.http(500, 'down'); } }
      : ok(c))
  });
  second.deps.connections = second.deps.connections.slice(0, 2);
  const outcome = BS.SyncAll.run(second.deps);
  assert.deepStrictEqual(outcome.results.map((r) => r.status), ['ok', 'error']);
  const grid = book.sheets.raw_accounts.grid;
  const col = grid[0].indexOf('connection');
  assert.ok(grid.slice(1).some((r) => r[col] === 'revolut-partner'), 'its earlier raw row is still there');
});

test('account name falls back to the readable type and currency when the bank sends no name', () => {
  const labelsFor = BS.AccountLabels.lookup([
    acct({ accountId: 'x', accountName: '', product: '', accountType: 'CACC', currency: 'SEK' }),
    acct({ accountId: 'y', accountName: '', product: '', accountType: 'ZZZZ', currency: 'EUR' }),
    acct({ accountId: 'z', accountName: 'Mine', product: 'P', accountType: 'CACC' })]);
  assert.strictEqual(labelsFor('x').accountName, 'Current account SEK');
  assert.strictEqual(labelsFor('y').accountName, 'ZZZZ EUR');
  assert.strictEqual(labelsFor('z').accountName, 'Mine');
});

test('transactions keep every API field (account numbers masked) for the raw tab', () => withMasking(BS, true, () => {
  const { transactions } = BS.EbNormalize.normalizeTransactions([{
    entry_reference: 'e1', status: 'BOOK', booking_date: '2026-03-01', credit_debit_indicator: 'DBIT',
    transaction_amount: { amount: '5.00', currency: 'EUR' }, creditor: { name: 'Shop' },
    creditor_account: { iban: 'GB82WEST12345698765432' }, bank_transaction_code: { code: 'PMNT' }, merchant_category_code: '5411'
  }], 'a1');
  const f = transactions[0].apiFields;
  assert.strictEqual(f['creditor.name'], 'Shop');
  assert.strictEqual(f['creditor_account.iban'], '****5432');
  assert.strictEqual(f['bank_transaction_code.code'], 'PMNT');
  assert.strictEqual(f.merchant_category_code, '5411');
  assert.ok(!JSON.stringify(f).includes('GB82'));
}));

test('sync: raw rows carry API fields and labels; the stored rows never get apiFields', () => {
  let stored = [];
  const rows = [{ ...txn({ transactionId: 'r1', bookingDate: '2026-03-30' }), apiFields: { 'creditor.name': 'Shop' } }];
  const result = BS.TxnSync.run({
    provider: { getTransactions: () => rows },
    accounts: [acct({})],
    store: { load: () => stored, save: (r) => { stored = r; return r.length; } },
    nowMs: () => NOW, todayIso: () => '2026-04-01'
  });
  assert.ok(result.rawRows.length >= 1);
  assert.deepStrictEqual([result.rawRows[0].apiFields['creditor.name'], result.rawRows[0].connection], ['Shop', 'revolut-niklas']);
  assert.ok(stored.length >= 1 && stored.every((t) => t.apiFields === undefined));
});

test('sync all: raw_transactions shows each connection\'s latest fetch with api.* columns and keeps the others', () => {
  const book = fakeSheetBook();
  const provider = (field) => (c) => ({
    listAccounts: () => [acct({ accountId: 'a-' + c.key, connection: c.key })],
    getBalances: () => [],
    getTransactions: (id) => [{ ...txn({ accountId: id, transactionId: 't-' + id }), apiFields: { 'creditor.name': field } }]
  });
  const first = runAll({ rawWriter: BS.RawWriter.create(book), isConnected: () => true, providerFor: provider('First') });
  first.deps.connections = first.deps.connections.slice(0, 2);
  BS.SyncAll.run(first.deps);
  const second = runAll({
    rawWriter: BS.RawWriter.create(book), isConnected: (c) => c.key === 'revolut-niklas', providerFor: provider('Second')
  });
  second.deps.connections = second.deps.connections.slice(0, 2);
  BS.SyncAll.run(second.deps);
  const [header, ...rows] = book.sheets.raw_transactions.grid;
  const by = (conn) => rows.find((r) => r[header.indexOf('connection')] === conn);
  assert.ok(header.includes('api.creditor.name'));
  assert.strictEqual(by('revolut-niklas')[header.indexOf('api.creditor.name')], 'Second');
  assert.strictEqual(by('revolut-partner')[header.indexOf('api.creditor.name')], 'First');
  assert.strictEqual(by('revolut-niklas')[header.indexOf('accountName')], 'Main');
});

test('sync all: raw_transactions accumulates across runs, newest first, and a refetched row replaces its older copy', () => {
  const book = fakeSheetBook();
  const fetchOf = (rows) => (c) => ({
    listAccounts: () => [acct({ accountId: 'a-' + c.key, connection: c.key })],
    getBalances: () => [],
    getTransactions: (id) => rows.map((r) => ({ ...txn({ accountId: id, transactionId: r.id, bookingDate: r.date }),
      apiFields: { note: r.note } }))
  });
  const run = (rows) => {
    const r = runAll({ rawWriter: BS.RawWriter.create(book), isConnected: () => true, providerFor: fetchOf(rows) });
    r.deps.connections = r.deps.connections.slice(0, 1);
    BS.SyncAll.run(r.deps);
  };
  run([{ id: 'old', date: '2026-01-05', note: 'first' }, { id: 'same', date: '2026-02-01', note: 'first' }]);
  run([{ id: 'same', date: '2026-02-01', note: 'second' }, { id: 'new', date: '2026-03-01', note: 'second' }]);
  const [header, ...rows] = book.sheets.raw_transactions.grid;
  const col = (name) => header.indexOf(name);
  assert.deepStrictEqual(rows.map((r) => r[col('transactionId')]), ['new', 'same', 'old']);
  assert.deepStrictEqual(rows.map((r) => r[col('api.note')]), ['second', 'second', 'first']);
});

test('raw writer: no column for an API field that is only a null parent of nested fields', () => {
  const book = fakeSheetBook();
  const mk = (api) => BS.RawWriter.toRow({ provider: 'p', connection: 'c', accountId: 'x' + Math.random(), apiFields: api });
  BS.RawWriter.create(book).writeTable('raw_transactions', BS.RawWriter.TRANSACTION_COLUMNS, [
    mk({ creditor: '', 'creditor_account.iban': '****1', note: '' }),
    mk({ 'creditor.name': 'Shop', 'creditor_account.iban': '', note: 'n' })]);
  const header = book.sheets.raw_transactions.grid[0];
  assert.ok(header.includes('api.creditor.name') && header.includes('api.note'));
  assert.ok(!header.includes('api.creditor'), 'null parent column dropped');
  assert.ok(header.includes('api.creditor_account.iban'));
});

test('balances name their description column clearly and keep the bank text', () => {
  const b = BS.EbNormalize.normalizeBalance({ name: 'Available balance', balance_type: 'ITAV',
    balance_amount: { amount: '1.00', currency: 'SEK' } }, 'a1', 't');
  assert.strictEqual(b.description, 'Available balance');
  assert.strictEqual(b.name, undefined);
  assert.ok(BS.RawWriter.BALANCE_COLUMNS.some((c) => c.key === 'description'));
  assert.ok(!BS.RawWriter.BALANCE_COLUMNS.some((c) => c.key === 'name'));
});

function migrationSetup(session) {
  const props = memProps();
  if (session) props.setProperty('EB_SESSION_revolut', JSON.stringify(session));
  const store = BS.EbSessionStore.create(props, () => NOW);
  const connections = BS.Connections.create(memProps());
  return { props, store, connections };
}

test('legacy migration moves the session, records the connection once, and is a no-op after the flag', () => {
  const { props, store, connections } = migrationSetup({ sessionId: 'old', accounts: [] });
  BS.LegacyMigration.run(props, store, connections);
  assert.strictEqual(store.loadSession('revolut-niklas').sessionId, 'old');
  assert.strictEqual(store.loadSession('revolut'), null);
  assert.deepStrictEqual(connections.all().map((c) => c.key), ['revolut-niklas']);
  assert.ok(props.getProperty('EB_MIGRATED'));
  props.setProperty('EB_SESSION_revolut', JSON.stringify({ sessionId: 'newer' }));
  BS.LegacyMigration.run(props, store, connections);
  assert.strictEqual(store.loadSession('revolut').sessionId, 'newer', 'flag set: nothing is touched again');
});

test('legacy migration repairs a run that stopped after moving the session but before saving the connection', () => {
  const { props, store, connections } = migrationSetup(null);
  store.saveSession('revolut-niklas', { sessionId: 'moved', accounts: [] });
  BS.LegacyMigration.run(props, store, connections);
  assert.deepStrictEqual(connections.all().map((c) => c.key), ['revolut-niklas']);
});

test('legacy migration converges when two callers run it one after the other', () => {
  const { props, store, connections } = migrationSetup({ sessionId: 'old', accounts: [] });
  const second = memProps();
  BS.LegacyMigration.run(props, store, connections);
  BS.LegacyMigration.run(second, store, connections);
  assert.deepStrictEqual(connections.all().map((c) => c.key), ['revolut-niklas']);
  assert.strictEqual(store.loadSession('revolut-niklas').sessionId, 'old');
});

test('legacy migration with no old session records no connection', () => {
  const { props, store, connections } = migrationSetup(null);
  BS.LegacyMigration.run(props, store, connections);
  assert.deepStrictEqual(connections.all(), []);
  assert.ok(props.getProperty('EB_MIGRATED'));
});

test('a full Script Properties quota gives a clear error for the session and is tolerated for API fields', () => {
  const full = { getProperty: () => null, setProperty: () => { throw new Error('Exceeded storage quota'); }, deleteProperty() {} };
  const store = BS.EbSessionStore.create(full, () => NOW);
  assert.throws(() => store.saveSession('revolut-niklas', { sessionId: 's' }), /Script Properties may be full/);
  assert.strictEqual(store.saveApiFields('a1', { k: 'v' }), false);
});

test('a connect that cannot store API fields still completes', () => {
  const props = memProps();
  const real = props.setProperty;
  props.setProperty = (k, v) => { if (k.startsWith('EB_API_')) throw new Error('quota'); real(k, v); };
  const store = BS.EbSessionStore.create(props, () => NOW);
  const connections = BS.Connections.create(memProps());
  const connection = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' });
  const auth = BS.EbAuth.create({ client: { createSession: () => ({ session_id: 's', accounts: [{ uid: 'a1', currency: 'EUR' }] }) },
    store, connections, nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x' });
  store.savePending('p', { connection, validUntil: '2027-01-01T00:00:00Z' });
  assert.strictEqual(auth.complete('code', 'p').accountCount, 1);
  assert.strictEqual(connections.all().length, 1);
});

test('refreshStale stops reading details once the time budget is spent and leaves the rest for next time', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  const connection = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' });
  store.saveSession(connection.key, { sessionId: 's', accessValidUntil: '2027-01-01T00:00:00Z',
    accounts: [acct({ accountId: 'a1', detailsAt: '' }), acct({ accountId: 'a2', detailsAt: '' }), acct({ accountId: 'a3', detailsAt: '' })] });
  const asked = [];
  const client = { getAccountDetails: (uid) => { asked.push(uid); return { uid, currency: 'EUR', details: 'N-' + uid }; } };
  const auth = BS.EbAuth.create({ client, store, connections: BS.Connections.create(memProps()), nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x' });
  const accounts = auth.refreshStale(connection, () => asked.length >= 1);
  assert.deepStrictEqual(asked, ['a1']);
  assert.deepStrictEqual(accounts.map((a) => a.detailsAt !== ''), [true, false, false]);
  asked.length = 0;
  auth.refreshStale(connection);
  assert.deepStrictEqual(asked, ['a2', 'a3'], 'the next sync picks up the rest');
});

test('sync all hands the shared time check to listAccounts', () => {
  const { deps } = runAll({ isConnected: () => true });
  let received;
  deps.connections = deps.connections.slice(0, 1);
  deps.providerFor = () => ({
    listAccounts: (options) => { received = options; return [acct({})]; },
    getBalances: () => [], getTransactions: () => []
  });
  BS.SyncAll.run(deps);
  assert.strictEqual(typeof received.outOfTime, 'function');
  assert.strictEqual(received.outOfTime(), false);
});

test('a failing details call does not block reading accounts, balances or transactions', () => {
  const store = BS.EbSessionStore.create(memProps(), () => NOW);
  const connection = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' });
  store.saveSession(connection.key, { sessionId: 's', accessValidUntil: '2027-01-01T00:00:00Z', accounts: [acct({ detailsAt: '' })] });
  const client = {
    getSession: () => ({ status: 'AUTHORIZED' }),
    getAccountDetails: () => { throw BS.Errors.http(500, 'down'); },
    getBalances: () => ({ balances: [{ name: 'b', balance_type: 'ITAV', balance_amount: { amount: '5', currency: 'EUR' } }] }),
    getTransactions: () => ({ transactions: [], continuation_key: null })
  };
  const auth = BS.EbAuth.create({ client, store, connections: BS.Connections.create(memProps()), nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x' });
  const provider = BS.EnableBankingProvider.create({ connection, client, store, auth, nowMs: () => NOW });
  const accounts = provider.listAccounts();
  assert.strictEqual(accounts.length, 1);
  assert.strictEqual(provider.getBalances('a1')[0].amount, 5);
  assert.deepStrictEqual(provider.getTransactions('a1', { dateFrom: '2026-01-01', dateTo: '2026-01-31' }), []);
});

test('holders that differ only in punctuation map to the same connection, so they replace rather than duplicate', () => {
  const a = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'A B' });
  const b = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'A-B' });
  assert.strictEqual(a.key, b.key);
  const store = BS.Connections.create(memProps());
  store.save(a);
  store.save(b);
  assert.strictEqual(store.all().length, 1, 'one connection and one session key, never two that fight');
});
