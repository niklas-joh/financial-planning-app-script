const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load');
const { fakeSheetBook } = require('./fake-sheet');

const raw = (o) => ({
  entry_reference: 'e1', status: 'BOOK', booking_date: '2026-01-10', value_date: '2026-01-10',
  transaction_amount: { amount: '12.50', currency: 'SEK' }, credit_debit_indicator: 'DBIT',
  creditor: { name: 'Shop' }, remittance_information: ['Coffee', 'beans'], ...o
});
const norm = (list) => BS.EbNormalize.normalizeTransactions(list, 'a1');
const tx = (o) => ({
  provider: 'enablebanking', bank: '', connection: '', holderName: '', accountName: '', ibanMasked: '', accountId: 'a1', transactionId: 't1', bookingDate: '2026-01-10', valueDate: '',
  status: 'booked', amount: -5, currency: 'EUR', amountEur: null, fxRate: null, fxRateDate: '',
  counterparty: '', description: '', ...o
});

test('normalize: debit is negative, credit positive, counterparty is the other party', () => {
  const { transactions } = norm([
    raw({}),
    raw({ entry_reference: 'e2', credit_debit_indicator: 'CRDT', debtor: { name: 'Boss' }, creditor: { name: 'Me' } })
  ]);
  assert.strictEqual(transactions[0].amount, -12.5);
  assert.strictEqual(transactions[0].counterparty, 'Shop');
  assert.strictEqual(transactions[0].description, 'Coffee beans');
  assert.strictEqual(transactions[1].amount, 12.5);
  assert.strictEqual(transactions[1].counterparty, 'Boss');
  assert.strictEqual(transactions[0].amountEur, null);
});

test('normalize: status mapping, skipped statuses and bad data', () => {
  const r = norm([raw({ entry_reference: 'a', status: 'PDNG' }), raw({ entry_reference: 'b', status: 'CNCL' }),
    raw({ entry_reference: 'c', status: 'SCHD' }), raw({ entry_reference: 'd', status: 'HOLD' })]);
  assert.deepStrictEqual(r.transactions.map((t) => t.status), ['pending', 'pending']);
  assert.strictEqual(r.skipped, 2);
  assert.throws(() => norm([raw({ credit_debit_indicator: undefined })]), { name: 'DataError' });
  assert.throws(() => norm([raw({ transaction_amount: { amount: 'x', currency: 'EUR' } })]), { name: 'DataError' });
  [ '', null, undefined ].forEach((amount) =>
    assert.throws(() => norm([raw({ transaction_amount: { amount, currency: 'EUR' } })]), { name: 'DataError' }));
});

test('normalize: id prefers entry_reference, then transaction_id, then a stable content hash', () => {
  const ids = (list) => norm(list).transactions.map((t) => t.transactionId);
  assert.deepStrictEqual(ids([raw({ transaction_id: 'x' })]), ['e1']);
  assert.deepStrictEqual(ids([raw({ entry_reference: null, transaction_id: 'x' })]), ['x']);
  const bare = raw({ entry_reference: null });
  const [first, second] = ids([bare, bare]);
  assert.notStrictEqual(first, second);
  assert.deepStrictEqual(ids([bare, bare]), [first, second]);
  // pending -> booked with unchanged content keeps its id
  assert.strictEqual(ids([{ ...bare, status: 'PDNG' }])[0], first);
});

test('merge: upsert is idempotent and reports counts', () => {
  const a = tx({ transactionId: 'a' });
  const first = BS.TxnMerge.merge([], [a], []);
  assert.strictEqual(first.added, 1);
  const again = BS.TxnMerge.merge(first.rows, [a], []);
  assert.deepStrictEqual([again.rows.length, again.added, again.updated], [1, 0, 0]);
  const edited = BS.TxnMerge.merge(first.rows, [tx({ transactionId: 'a', description: 'new' })], []);
  assert.deepStrictEqual([edited.added, edited.updated], [0, 1]);
});

test('merge: pending to booked under the same id updates in place', () => {
  const pending = tx({ transactionId: 'a', status: 'pending' });
  const m = BS.TxnMerge.merge([pending], [tx({ transactionId: 'a', status: 'booked' })], []);
  assert.strictEqual(m.rows.length, 1);
  assert.strictEqual(m.rows[0].status, 'booked');
});

test('merge: pending that books under a new id is replaced, not duplicated', () => {
  const pending = tx({ transactionId: 'p', status: 'pending', bookingDate: '2026-01-10' });
  const range = [{ accountId: 'a1', from: '2026-01-01', to: '2026-01-31' }];
  const m = BS.TxnMerge.merge([pending], [tx({ transactionId: 'b' })], range);
  assert.deepStrictEqual(m.rows.map((t) => t.transactionId), ['b']);
  assert.strictEqual(m.removed, 1);
});

test('merge: stale pending outside the fetched range or account is kept; undated only for includesUndated', () => {
  const range = [{ accountId: 'a1', from: '2026-01-01', to: '2026-01-31', includesUndated: true }];
  const outside = tx({ transactionId: 'o', status: 'pending', bookingDate: '2025-12-01' });
  const otherAccount = tx({ transactionId: 'x', status: 'pending', accountId: 'a2' });
  const undated = tx({ transactionId: 'u', status: 'pending', bookingDate: '' });
  const booked = tx({ transactionId: 'k', bookingDate: '2026-01-05' });
  const m = BS.TxnMerge.merge([outside, otherAccount, undated, booked], [], range);
  assert.deepStrictEqual(m.rows.map((t) => t.transactionId).sort(), ['k', 'o', 'x']);
  assert.strictEqual(m.removed, 1);
  const older = [{ accountId: 'a1', from: '2026-01-01', to: '2026-01-31' }];
  assert.strictEqual(BS.TxnMerge.merge([undated], [], older).rows.length, 1);
});

test('merge: a booked row is never downgraded; FX columns survive an unchanged re-fetch', () => {
  const booked = tx({ transactionId: 'a', amountEur: 0.5, fxRate: 0.1, fxRateDate: '2026-01-10' });
  assert.strictEqual(BS.TxnMerge.merge([booked], [tx({ transactionId: 'a', status: 'pending' })], []).rows[0].status, 'booked');
  const same = BS.TxnMerge.merge([booked], [tx({ transactionId: 'a' })], []).rows[0];
  assert.deepStrictEqual([same.amountEur, same.fxRate, same.fxRateDate], [0.5, 0.1, '2026-01-10']);
  const moved = BS.TxnMerge.merge([booked], [tx({ transactionId: 'a', amount: -9 })], []).rows[0];
  assert.strictEqual(moved.amountEur, null);
});

test('merge: sorted newest first and inputs are not mutated', () => {
  const existing = [tx({ transactionId: 'a', bookingDate: '2026-01-01' })];
  const snapshot = JSON.stringify(existing);
  const m = BS.TxnMerge.merge(existing, [tx({ transactionId: 'b', bookingDate: '2026-02-01' })], []);
  assert.deepStrictEqual(m.rows.map((t) => t.transactionId), ['b', 'a']);
  assert.strictEqual(JSON.stringify(existing), snapshot);
});

test('date range: addDays, startFor and planWindows', () => {
  const D = BS.DateRange;
  assert.strictEqual(D.addDays('2026-03-01', -1), '2026-02-28');
  const cfg = { INITIAL_LOOKBACK_DAYS: 90, OVERLAP_DAYS: 7 };
  assert.strictEqual(D.startFor('', '2026-04-01', cfg), '2026-01-01');
  assert.strictEqual(D.startFor('2026-03-20', '2026-04-01', cfg), '2026-03-13');
  const w = D.planWindows('2026-01-01', '2026-01-10', 4);
  assert.deepStrictEqual(w, [
    { from: '2026-01-01', to: '2026-01-04' }, { from: '2026-01-05', to: '2026-01-08' },
    { from: '2026-01-09', to: '2026-01-10' }]);
  assert.deepStrictEqual(D.planWindows('2026-01-02', '2026-01-01', 4), []);
});

function providerWith(pages) {
  const calls = [];
  const client = { getTransactions: (id, params) => { calls.push(params); return pages[calls.length - 1]; } };
  const session = { sessionId: 's', accessValidUntil: '2999-01-01T00:00:00Z', accounts: [] };
  const fakeStore = { loadSession: () => session };
  const provider = BS.EnableBankingProvider.create({
    connection: BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' }), client, store: fakeStore, auth: {}, nowMs: () => 0 });
  return { provider, calls };
}

test('provider.getTransactions follows continuation keys across pages', () => {
  const { provider, calls } = providerWith([
    { transactions: [raw({ entry_reference: 'a' })], continuation_key: 'k1' },
    { transactions: [raw({ entry_reference: 'b' })], continuation_key: null }]);
  const result = provider.getTransactions('a1', { dateFrom: '2026-01-01', dateTo: '2026-01-31' });
  assert.deepStrictEqual(result.map((t) => t.transactionId), ['a', 'b']);
  assert.deepStrictEqual(calls[0], { date_from: '2026-01-01', date_to: '2026-01-31', continuation_key: undefined });
  assert.strictEqual(calls[1].continuation_key, 'k1');
});

test('provider.getTransactions fails when pagination never ends', () => {
  const page = { transactions: [], continuation_key: 'again' };
  const { provider } = providerWith(Array(BS.Config.TRANSACTIONS.MAX_PAGES + 1).fill(page));
  assert.throws(() => provider.getTransactions('a1', { dateFrom: 'x', dateTo: 'y' }), { name: 'DataError' });
});

test('client.getTransactions builds the query and omits undefined parameters', () => {
  let url;
  const client = BS.EbClient.create({ request: (r) => { url = r.url; return {}; } }, () => 't');
  client.getTransactions('u 1', { date_from: '2026-01-01', date_to: undefined, continuation_key: 'k/1' });
  assert.ok(url.endsWith('/accounts/u%201/transactions?date_from=2026-01-01&continuation_key=k%2F1'));
});

test('store round-trips rows, forces text format on text columns, tolerates Date cells', () => {
  const book = fakeSheetBook();
  const store = BS.TransactionsStore.create(book, (d) => d.toISOString().slice(0, 10));
  assert.deepStrictEqual(store.load(), []);
  const rows = [tx({ transactionId: '007', description: '=SUM(A1)', amountEur: -0.5, fxRate: 0.1 })];
  assert.strictEqual(store.save(rows), 1);
  // A row without a category reads back with an empty one.
  assert.deepStrictEqual(store.load(), rows.map((r) => ({ ...r, category: '' })));
  store.save([{ ...rows[0], category: 'GROCERIES' }]);
  assert.strictEqual(store.load()[0].category, 'GROCERIES');
  const sheet = book.sheets.transactions;
  assert.strictEqual(sheet.frozen, 1);
  const textCols = sheet.formats.map((f) => f[0]);
  const keys = BS.TransactionsStore.COLUMNS.map((c) => c.key);
  assert.ok(textCols.includes(keys.indexOf('description') + 1));
  assert.ok(textCols.includes(keys.indexOf('transactionId') + 1));
  assert.ok(!textCols.includes(keys.indexOf('amount') + 1));
  sheet.grid[1][keys.indexOf('bookingDate')] = new Date('2026-01-10T00:00:00Z');
  assert.strictEqual(store.load()[0].bookingDate, '2026-01-10');
});

test('store: save never clears the whole tab, trims leftovers, keeps columns to the right', () => {
  const book = fakeSheetBook();
  const store = BS.TransactionsStore.create(book, (d) => d.toISOString().slice(0, 10));
  store.save([tx({ transactionId: 'a' }), tx({ transactionId: 'b' }), tx({ transactionId: 'c' })]);
  const sheet = book.sheets.transactions;
  sheet.grid[1][BS.TransactionsStore.COLUMNS.length] = 'my note';
  store.save([tx({ transactionId: 'a' })]);
  assert.strictEqual(sheet.grid.length, 2);
  assert.strictEqual(sheet.grid[1][BS.TransactionsStore.COLUMNS.length], 'my note');
  assert.deepStrictEqual(sheet.cleared[0].slice(0, 3), [3, 1, 2]);
  assert.strictEqual(store.load().length, 1);
});

test('store: load fails loudly when the header was changed', () => {
  const book = fakeSheetBook();
  const store = BS.TransactionsStore.create(book, (d) => d.toISOString().slice(0, 10));
  store.save([tx({ transactionId: 'a' })]);
  const keys = BS.TransactionsStore.COLUMNS.map((c) => c.key);
  book.sheets.transactions.grid[0][keys.indexOf('accountId')] = 'renamed';
  assert.throws(() => store.load(), /accountId/);
});

test('store: an older tab without the label or category columns, or with reordered columns, still loads', () => {
  const keys = BS.TransactionsStore.COLUMNS.map((c) => c.key);
  const legacyKeys = keys.filter((k) => !['bank', 'connection', 'holderName', 'accountName', 'ibanMasked', 'category'].includes(k));
  const row = tx({ transactionId: 'a', amount: -2, category: '' });
  const cells = (order) => order.map((k) => (row[k] === null ? '' : row[k]));
  const book = fakeSheetBook();
  const sheet = book.insertSheet('transactions');
  sheet.grid = [legacyKeys, cells(legacyKeys)];
  const store = BS.TransactionsStore.create(book, (d) => d.toISOString());
  assert.deepStrictEqual(store.load(), [row]);
  const reversed = legacyKeys.slice().reverse();
  sheet.grid = [reversed, cells(reversed)];
  assert.deepStrictEqual(store.load(), [row]);
  store.save([row]);
  assert.deepStrictEqual(sheet.grid[0].slice(0, keys.length), keys);
});

function syncDeps(over) {
  const saves = [];
  let rows = [];
  const fetches = [];
  const deps = {
    provider: {
      getTransactions: (id, range) => { fetches.push(range); return [tx({ transactionId: 'n' + fetches.length, bookingDate: range.dateTo })]; }
    },
    store: { load: () => rows, save: (r) => { rows = r; saves.push(r.length); return r.length; } },
    accounts: [{ accountId: 'a1', bank: 'Revolut', connection: 'revolut-niklas', accountName: 'Main', ibanMasked: '****1' }],
    nowMs: () => Date.parse('2026-04-01T12:00:00Z'),
    todayIso: () => '2026-04-01',
    ...over
  };
  return { deps, saves, fetches, rows: () => rows };
}

test('sync: first run backfills in windows and saves once per account; a re-run adds nothing', () => {
  const { deps, saves, fetches } = syncDeps();
  const result = BS.TxnSync.run(deps);
  assert.strictEqual(result.complete, true);
  assert.strictEqual(fetches[0].dateFrom, '2026-01-01');
  assert.strictEqual(fetches[fetches.length - 1].dateTo, '2026-04-01');
  assert.strictEqual(saves.length, 1);
  assert.strictEqual(result.added, fetches.length);
  const again = BS.TxnSync.run({ ...deps, provider: { ...deps.provider,
    getTransactions: (id, range) => [tx({ transactionId: 'n' + (fetches.findIndex((f) => f.dateTo === range.dateTo) + 1), bookingDate: range.dateTo })] } });
  assert.strictEqual(again.added, 0);
});

test('sync: later runs start at the latest booked date minus the overlap', () => {
  const { deps, fetches } = syncDeps();
  BS.TxnSync.run(deps);
  fetches.length = 0;
  BS.TxnSync.run(deps);
  assert.strictEqual(fetches[0].dateFrom, '2026-03-25');
});

test('sync: stops at the time budget, saves what it has, and the next run continues', () => {
  let now = Date.parse('2026-04-01T12:00:00Z');
  const { deps, saves, rows } = syncDeps({ nowMs: () => { now += 100000; return now; } });
  const partial = BS.TxnSync.run(deps);
  assert.strictEqual(partial.complete, false);
  assert.strictEqual(saves.length, 1);
  assert.ok(rows().length >= 1 && rows().length < 4);
});

test('sync: with two accounts, an exhausted budget skips the second without saving it', () => {
  let now = Date.parse('2026-04-01T12:00:00Z');
  const base = syncDeps({ nowMs: () => { now += 100000; return now; } });
  base.deps.accounts = [{ accountId: 'a1' }, { accountId: 'a2' }];
  const result = BS.TxnSync.run(base.deps);
  assert.strictEqual(result.complete, false);
  assert.ok(base.rows().every((t) => t.accountId === 'a1'));
});

test('sync: a failing fetch still saves what was merged before it, then rethrows', () => {
  const { deps, saves, rows } = syncDeps();
  let calls = 0;
  const inner = deps.provider.getTransactions;
  deps.provider.getTransactions = (id, range) => { if (++calls === 3) throw BS.Errors.http(500, 'boom'); return inner(id, range); };
  assert.throws(() => BS.TxnSync.run(deps), { name: 'HttpError' });
  assert.strictEqual(saves.length, 1);
  assert.strictEqual(rows().length, 2);
});

test('sync: undated pending rows survive early windows and are dropped only in the window ending today', () => {
  const stale = tx({ transactionId: 'u', status: 'pending', bookingDate: '' });
  const { deps, rows } = syncDeps();
  deps.store.save(deps.store.load().concat([stale]));
  deps.provider.getTransactions = () => [];
  BS.TxnSync.run(deps);
  assert.strictEqual(rows().length, 0);
});

test('sync: provider errors propagate', () => {
  const { deps } = syncDeps({ provider: { getTransactions: () => { throw BS.Errors.auth('expired'); } } });
  assert.throws(() => BS.TxnSync.run(deps), { name: 'AuthError' });
});

test('merge: a row whose keys come in another order than the stored copy is not counted as updated', () => {
  // A provider builds its row with labels appended last; the copy read back from the sheet has the
  // columns' order. Same values, so a re-fetch must report 0 updated (seen live with Trading 212).
  const conn = BS.T212Credentials.connections({ getKeys: () => ['T212_API_KEY_ALEX'] })[0];
  const account = BS.T212Normalize.account({ id: 7, currency: 'EUR' }, conn);
  const labels = BS.AccountLabels.lookup([account]);
  const fetched = () => {
    const t = BS.AccountLabels.stamp(BS.T212Normalize.cashMovement(
      { type: 'DEPOSIT', amount: 5, currency: 'EUR', dateTime: '2026-10-20T10:00:00Z', reference: 'r1' }, '7', (iso) => iso.slice(0, 10)), labels);
    delete t.apiFields;
    return t;
  };
  const book = fakeSheetBook();
  const store = BS.TransactionsStore.create(book, (d) => d.toISOString().slice(0, 10));
  const range = [{ accountId: '7', from: '2026-10-01', to: '2026-10-25' }];
  store.save(BS.TxnMerge.merge(store.load(), [fetched()], range).rows);
  const again = BS.TxnMerge.merge(store.load(), [fetched()], range);
  assert.deepStrictEqual([again.added, again.updated, again.removed], [0, 0, 0]);
  const amountChanged = Object.assign(fetched(), { amount: 6 });
  assert.strictEqual(BS.TxnMerge.merge(store.load(), [amountChanged], range).updated, 1, 'a real change still counts');
});

test('merge: a re-fetch without counterparty or category keeps the stored ones (details from another source)', () => {
  const stored = tx({ transactionId: 'a', counterparty: 'Demo Cafe', category: 'RESTAURANTS' });
  const again = BS.TxnMerge.merge([stored], [tx({ transactionId: 'a' })], []);
  assert.deepStrictEqual([again.added, again.updated], [0, 0]);
  assert.deepStrictEqual([again.rows[0].counterparty, again.rows[0].category], ['Demo Cafe', 'RESTAURANTS']);
  // A row read back from the sheet has category '' where the provider's row has none: not a change.
  const fromSheet = BS.TxnMerge.merge([tx({ transactionId: 'b', category: '' })], [tx({ transactionId: 'b' })], []);
  assert.strictEqual(fromSheet.updated, 0);
  const renamed = BS.TxnMerge.merge([stored], [tx({ transactionId: 'a', counterparty: 'Other' })], []);
  assert.deepStrictEqual([renamed.updated, renamed.rows[0].counterparty, renamed.rows[0].category], [1, 'Other', 'RESTAURANTS']);
});

test('enrich: details fill the empty fields of rows with the same provider and id, never replace a value, idempotent', () => {
  const rows = [tx({ transactionId: 'a', provider: 'trading212' }), tx({ transactionId: 'a' }), tx({ transactionId: 'b', provider: 'trading212', counterparty: 'Kept' })];
  const details = [
    { provider: 'trading212', transactionId: 'a', counterparty: 'Demo Cafe', category: 'RESTAURANTS', date: '2026-01-10' },
    { provider: 'trading212', transactionId: 'b', counterparty: 'Report name', category: 'SPENDING_CASHBACK' },
    { provider: 'trading212', transactionId: 'missing', counterparty: 'X', category: 'Y' }
  ];
  const once = BS.TxnMerge.enrich(rows, details);
  assert.strictEqual(once.enriched, 2);
  assert.deepStrictEqual(once.rows.map((t) => [t.provider, t.counterparty, t.category]), [
    ['trading212', 'Demo Cafe', 'RESTAURANTS'], ['enablebanking', '', undefined], ['trading212', 'Kept', 'SPENDING_CASHBACK']
  ]);
  assert.ok(!('date' in once.rows[0]));
  assert.strictEqual(rows[0].counterparty, '');
  assert.strictEqual(BS.TxnMerge.enrich(once.rows, details).enriched, 0);
});
