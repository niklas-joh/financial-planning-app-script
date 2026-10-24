const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load');

const raw = (o) => ({
  entry_reference: 'e1', status: 'BOOK', booking_date: '2026-01-10', value_date: '2026-01-10',
  transaction_amount: { amount: '12.50', currency: 'SEK' }, credit_debit_indicator: 'DBIT',
  creditor: { name: 'Shop' }, remittance_information: ['Coffee', 'beans'], ...o
});
const norm = (list) => BS.EbNormalize.normalizeTransactions(list, 'a1');
const tx = (o) => ({
  provider: 'enablebanking', accountId: 'a1', transactionId: 't1', bookingDate: '2026-01-10', valueDate: '',
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

test('merge: stale pending outside the fetched range or account is kept; undated pending is covered', () => {
  const range = [{ accountId: 'a1', from: '2026-01-01', to: '2026-01-31' }];
  const outside = tx({ transactionId: 'o', status: 'pending', bookingDate: '2025-12-01' });
  const otherAccount = tx({ transactionId: 'x', status: 'pending', accountId: 'a2' });
  const undated = tx({ transactionId: 'u', status: 'pending', bookingDate: '' });
  const booked = tx({ transactionId: 'k', bookingDate: '2026-01-05' });
  const m = BS.TxnMerge.merge([outside, otherAccount, undated, booked], [], range);
  assert.deepStrictEqual(m.rows.map((t) => t.transactionId).sort(), ['k', 'o', 'x']);
  assert.strictEqual(m.removed, 1);
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
    bank: BS.Config.BANKS.REVOLUT, client, store: fakeStore, auth: {}, nowMs: () => 0 });
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

function fakeSheetBook() {
  const sheets = {};
  const makeSheet = () => {
    const s = { grid: [], formats: [], frozen: 0 };
    s.getLastRow = () => s.grid.length;
    s.clearContents = () => { s.grid = []; };
    s.setFrozenRows = (n) => { s.frozen = n; };
    s.getRange = (row, col, rows, cols) => ({
      getValues: () => s.grid.slice(row - 1, row - 1 + rows).map((r) => r.slice(col - 1, col - 1 + cols)),
      setValues: (m) => { m.forEach((r, i) => { s.grid[row - 1 + i] = r.slice(); }); },
      setNumberFormat: (f) => { s.formats.push([col, f]); }
    });
    return s;
  };
  return {
    sheets,
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet())
  };
}

test('store round-trips rows, forces text format on text columns, tolerates Date cells', () => {
  const book = fakeSheetBook();
  const store = BS.TransactionsStore.create(book, (d) => d.toISOString().slice(0, 10));
  assert.deepStrictEqual(store.load(), []);
  const rows = [tx({ transactionId: '007', description: '=SUM(A1)', amountEur: -0.5, fxRate: 0.1 })];
  assert.strictEqual(store.save(rows), 1);
  assert.deepStrictEqual(store.load(), rows);
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

function syncDeps(over) {
  const saves = [];
  let rows = [];
  const fetches = [];
  const deps = {
    provider: {
      listAccounts: () => [{ accountId: 'a1' }],
      getTransactions: (id, range) => { fetches.push(range); return [tx({ transactionId: 'n' + fetches.length, bookingDate: range.dateTo })]; }
    },
    store: { load: () => rows, save: (r) => { rows = r; saves.push(r.length); return r.length; } },
    nowMs: () => Date.parse('2026-04-01T12:00:00Z'),
    ...over
  };
  return { deps, saves, fetches, rows: () => rows };
}

test('sync: first run backfills in windows, saving after each, then is idempotent', () => {
  const { deps, saves, fetches } = syncDeps();
  const result = BS.TxnSync.run(deps);
  assert.strictEqual(result.complete, true);
  assert.strictEqual(fetches[0].dateFrom, '2026-01-01');
  assert.strictEqual(fetches[fetches.length - 1].dateTo, '2026-04-01');
  assert.strictEqual(saves.length, fetches.length);
  assert.strictEqual(result.added, fetches.length);
});

test('sync: later runs start at the latest booked date minus the overlap', () => {
  const { deps, fetches } = syncDeps();
  BS.TxnSync.run(deps);
  fetches.length = 0;
  BS.TxnSync.run(deps);
  assert.strictEqual(fetches[0].dateFrom, '2026-03-25');
});

test('sync: stops at the time budget without losing saved work, and the next run continues', () => {
  let now = Date.parse('2026-04-01T12:00:00Z');
  const { deps, saves, rows } = syncDeps({ nowMs: () => { now += 100000; return now; } });
  const partial = BS.TxnSync.run(deps);
  assert.strictEqual(partial.complete, false);
  assert.ok(saves.length >= 1 && saves.length < 4);
  assert.strictEqual(rows().length, saves[saves.length - 1]);
});

test('sync: provider errors propagate', () => {
  const { deps } = syncDeps({ provider: { listAccounts: () => { throw BS.Errors.auth('expired'); } } });
  assert.throws(() => BS.TxnSync.run(deps), { name: 'AuthError' });
});
