const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load');
const { fakeSheetBook } = require('./fake-sheet');

// Fake rates (units per 1 EUR). 2026-04-03 (Good Friday) and 2026-04-06 (Easter Monday) have no ECB rate.
const SEK = [
  { date: '2026-04-01', rate: 10.9 },
  { date: '2026-04-02', rate: 11 },
  { date: '2026-04-07', rate: 11.2 }
];
const opts = { reporting: 'EUR', maxAgeDays: 7 };

function fakeCache() {
  const store = {};
  return {
    store,
    get: (k) => (k in store ? store[k] : null),
    put: (k, v) => { store[k] = v; },
    remove: (k) => { delete store[k]; }
  };
}

function service(over = {}) {
  const calls = [];
  let saved = null;
  const deps = {
    fetchRates: (cur, from, to) => { calls.push([cur, from, to]); return SEK.filter((r) => r.date >= '2026-04-01' && r.date <= to); },
    rates: { load: () => ({}), save: (t) => { saved = t; return 0; } },
    cache: fakeCache(),
    todayIso: () => '2026-04-07',
    ...over
  };
  return { fx: BS.FxService.create(deps), calls, saved: () => saved, deps };
}

const tx = (o) => ({ provider: 'p', accountId: 'a1', transactionId: 't1', status: 'booked', bookingDate: '', valueDate: '',
  amount: -5, currency: 'EUR', amountEur: null, fxRate: null, fxRateDate: '', counterparty: '', description: '', ...o });
const txn = (o) => ({ accountId: 'a1', transactionId: 't', status: 'booked', bookingDate: '2026-04-02', valueDate: '',
  amount: -110, currency: 'SEK', amountEur: null, fxRate: null, fxRateDate: '', ...o });

test('fx: rate of the day, else the previous business day (weekend, Easter), never too old or before the series', () => {
  assert.deepStrictEqual(BS.Fx.rateOn(SEK, '2026-04-02', 7), SEK[1]);
  assert.deepStrictEqual(BS.Fx.rateOn(SEK, '2026-04-03', 7), SEK[1]);
  assert.deepStrictEqual(BS.Fx.rateOn(SEK, '2026-04-06', 7), SEK[1]);
  assert.deepStrictEqual(BS.Fx.rateOn(SEK, '2026-04-14', 7), SEK[2]);
  assert.strictEqual(BS.Fx.rateOn(SEK, '2026-04-15', 7), null);
  assert.strictEqual(BS.Fx.rateOn(SEK, '2026-03-31', 7), null);
  assert.strictEqual(BS.Fx.rateOn(undefined, '2026-04-02', 7), null);
});

test('fx: converts by dividing by the rate, rounds to cents, keeps the sign, EUR passes through', () => {
  assert.strictEqual(BS.Fx.toReporting(-110, 11), -10);
  assert.strictEqual(BS.Fx.toReporting(100, 10.9), 9.17);
  assert.ok(Object.is(BS.Fx.toReporting(-0.001, 11), 0));
  const table = { SEK };
  assert.deepStrictEqual(BS.Fx.convert(-110, 'SEK', '2026-04-04', table, opts), { amountEur: -10, fxRate: 11, fxRateDate: '2026-04-02' });
  assert.deepStrictEqual(BS.Fx.convert(-5, 'EUR', '2026-04-04', table, opts), { amountEur: -5, fxRate: 1, fxRateDate: '2026-04-04' });
  assert.strictEqual(BS.Fx.convert(5, 'USD', '2026-04-04', table, opts), null);
  assert.strictEqual(BS.Fx.convert(NaN, 'SEK', '2026-04-04', table, opts), null);
});

test('fx: mergeSeries replaces by date and sorts; missingRanges asks only for the edges', () => {
  const merged = BS.Fx.mergeSeries([SEK[2], SEK[0]], [{ date: '2026-04-01', rate: 9 }, SEK[1]]);
  assert.deepStrictEqual(merged.map((r) => [r.date, r.rate]), [['2026-04-01', 9], ['2026-04-02', 11], ['2026-04-07', 11.2]]);
  assert.deepStrictEqual(BS.Fx.missingRanges([], '2026-01-01', '2026-04-07'), [{ from: '2026-01-01', to: '2026-04-07' }]);
  assert.deepStrictEqual(BS.Fx.missingRanges(SEK, '2026-04-02', '2026-04-07'), []);
  assert.deepStrictEqual(BS.Fx.missingRanges(SEK, '2026-03-20', '2026-04-09'),
    [{ from: '2026-03-20', to: '2026-04-01' }, { from: '2026-04-07', to: '2026-04-09' }]);
});

test('fx service: one fetch per currency for the whole range, stored on the tab and in the cache', () => {
  const { fx, calls, saved, deps } = service();
  const rows = [txn({ transactionId: 'a' }), txn({ transactionId: 'b', bookingDate: '2026-04-06' }),
    txn({ transactionId: 'c', currency: 'EUR', amount: -3 }), txn({ transactionId: 'd', bookingDate: '2026-04-01', amount: 109 })];
  const out = fx.convertTransactions(rows);
  assert.deepStrictEqual(calls, [['SEK', '2026-04-01', '2026-04-07']]);
  assert.deepStrictEqual(out.rows.map((t) => [t.amountEur, t.fxRate, t.fxRateDate]),
    [[-10, 11, '2026-04-02'], [-10, 11, '2026-04-02'], [-3, 1, '2026-04-02'], [10, 10.9, '2026-04-01']]);
  assert.deepStrictEqual([out.changed, out.missing, out.failed], [4, 0, []]);
  assert.deepStrictEqual(saved().SEK, SEK);
  assert.deepStrictEqual(JSON.parse(deps.cache.get('fx:rates')).SEK, SEK);
  assert.strictEqual(rows[0].amountEur, null, 'inputs are not mutated');
  // Balances in the same run need no further fetch and use the latest rate.
  const balances = fx.convertBalances([{ accountId: 'a1', amount: 1120, currency: 'SEK' }]);
  assert.deepStrictEqual(calls.length, 1);
  assert.deepStrictEqual([balances.rows[0].amountEur, balances.rows[0].fxRateDate], [100, '2026-04-07']);
});

test('fx service: a second run reads the cache, fetches nothing, and leaves converted rows unchanged', () => {
  const first = service();
  const converted = first.fx.convertTransactions([txn({})]).rows;
  const second = service({ cache: first.deps.cache, rates: { load: () => { throw new Error('tab not needed'); }, save: () => 0 } });
  const out = second.fx.convertTransactions(converted);
  assert.deepStrictEqual([second.calls.length, out.changed, out.missing], [0, 0, 0]);
  assert.strictEqual(out.rows[0], converted[0]);
});

test('fx service: a range already checked recently is not fetched again', () => {
  const cache = fakeCache();
  const empty = { fetchRates: (cur, from, to) => { calls.push([cur, from, to]); return []; } };
  const calls = [];
  service({ ...empty, cache }).fx.convertTransactions([txn({})]);
  service({ ...empty, cache }).fx.convertTransactions([txn({})]);
  assert.strictEqual(calls.length, 1);
});

test('fx service: only rows without an EUR amount are converted; converted rows are never recalculated', () => {
  const done = txn({ transactionId: 'old', bookingDate: '2026-01-05', amountEur: -9.5, fxRate: 11.58, fxRateDate: '2026-01-05' });
  const { fx, calls } = service();
  const out = fx.convertTransactions([done, txn({ transactionId: 'new' })]);
  assert.strictEqual(out.rows[0], done);
  assert.strictEqual(out.rows[1].amountEur, -10);
  assert.deepStrictEqual([out.changed, out.missing], [1, 0]);
  // The old row's date does not widen the fetch: only the new row's range is asked for.
  assert.deepStrictEqual(calls, [['SEK', '2026-04-02', '2026-04-07']]);
  const none = service();
  const again = none.fx.convertTransactions(out.rows);
  assert.deepStrictEqual([none.calls.length, again.changed], [0, 0]);
});

test('fx service: a row dated today waits for today\'s rate instead of storing the previous day\'s', () => {
  // Run on 04-07 before the ECB publishes that day: only 04-02 is known.
  const early = service({ fetchRates: () => SEK.slice(0, 2) });
  const first = early.fx.convertTransactions([txn({ bookingDate: '2026-04-07' }), txn({ transactionId: 'h', bookingDate: '2026-04-06' })]);
  assert.strictEqual(first.rows[0].amountEur, null);
  assert.deepStrictEqual([first.rows[1].fxRate, first.rows[1].fxRateDate], [11, '2026-04-02'], 'a past holiday uses the previous business day');
  assert.deepStrictEqual([first.changed, first.waiting, first.missing], [1, 1, 0]);
  // Later run (cache expired): the 04-07 rate exists now.
  const later = service({ rates: { load: () => ({ SEK: SEK.slice(0, 2) }), save: () => 0 } });
  const out = later.fx.convertTransactions(first.rows);
  assert.deepStrictEqual(later.calls, [['SEK', '2026-04-02', '2026-04-07']]);
  assert.deepStrictEqual([out.changed, out.rows[0].fxRate, out.rows[0].fxRateDate, out.rows[0].amountEur], [1, 11.2, '2026-04-07', -9.82]);
  // Balances always use the latest rate, even before today's is published.
  const balance = service({ fetchRates: () => SEK.slice(0, 2) }).fx.convertBalances([{ amount: 110, currency: 'SEK' }]);
  assert.deepStrictEqual([balance.rows[0].amountEur, balance.rows[0].fxRateDate], [10, '2026-04-02']);
});

test('fx service: a failed fetch leaves rows empty, reports the currency, and other currencies still convert', () => {
  const { fx } = service({
    fetchRates: (cur) => { if (cur === 'USD') throw BS.Errors.http(422, 'invalid currency'); return SEK; }
  });
  const usd = txn({ transactionId: 'u', currency: 'USD' });
  const out = fx.convertTransactions([usd, txn({})]);
  assert.strictEqual(out.rows[0], usd);
  assert.strictEqual(out.rows[1].amountEur, -10);
  assert.deepStrictEqual([out.changed, out.missing], [1, 1]);
  assert.match(out.failed[0], /^USD: HttpError: HTTP 422/);
});

test('merge: a changed amount, currency or date clears the EUR columns so the row is converted again', () => {
  const stored = tx({ amountEur: -10, fxRate: 11, fxRateDate: '2026-04-02', bookingDate: '2026-04-02' });
  const keep = BS.TxnMerge.merge([stored], [tx({ bookingDate: '2026-04-02', description: 'edited' })], []).rows[0];
  assert.strictEqual(keep.amountEur, -10);
  [{ amount: -6 }, { currency: 'SEK' }, { bookingDate: '2026-04-03' }].forEach((change) => {
    const row = BS.TxnMerge.merge([stored], [tx({ bookingDate: '2026-04-02', ...change })], []).rows[0];
    assert.ok(BS.Fx.needsConversion(row), JSON.stringify(change));
  });
});

test('fx service: undated pending rows convert at the latest rate', () => {
  const out = service().fx.convertTransactions([txn({ status: 'pending', bookingDate: '', valueDate: '', amount: -112 })]);
  assert.deepStrictEqual([out.rows[0].amountEur, out.rows[0].fxRateDate], [-10, '2026-04-07']);
});

test('frankfurter: asks for ECB rates in one call and validates the answer', () => {
  const requests = [];
  const reply = [{ date: '2026-04-07', base: 'EUR', quote: 'SEK', rate: 11.2 }, { date: '2026-04-02', base: 'EUR', quote: 'SEK', rate: 11 }];
  const client = BS.Frankfurter.create({ request: (r) => { requests.push(r); return reply; } });
  assert.deepStrictEqual(client.fetchRates('SEK', '2026-04-01', '2026-04-07'), [SEK[1], SEK[2]]);
  assert.strictEqual(requests[0].url,
    'https://api.frankfurter.dev/v2/rates?from=2026-04-01&to=2026-04-07&base=EUR&quotes=SEK&providers=ecb');
  const bad = (body) => BS.Frankfurter.create({ request: () => body }).fetchRates('SEK', '2026-04-01', '2026-04-07');
  assert.deepStrictEqual(bad([]), []);
  assert.throws(() => bad({ rates: {} }), { name: 'DataError' });
  assert.throws(() => bad([{ date: '2026-04-07', base: 'EUR', quote: 'USD', rate: 1.1 }]), { name: 'DataError' });
  assert.throws(() => bad([{ date: '2026-04-07', base: 'EUR', quote: 'SEK', rate: 0 }]), { name: 'DataError' });
});

test('rates store: round-trips through the tab, newest first, dates kept as text', () => {
  const book = fakeSheetBook();
  const store = BS.RatesStore.create(book, (d) => d.toISOString().slice(0, 10));
  assert.deepStrictEqual(store.load(), {});
  store.save({ SEK, NOK: [{ date: '2026-04-02', rate: 11.5 }] });
  const [header, ...rows] = book.sheets.rates.grid;
  assert.deepStrictEqual(header, ['date', 'base', 'currency', 'rate', 'source']);
  assert.deepStrictEqual(rows[0], ['2026-04-02', 'EUR', 'NOK', 11.5, 'ECB']);
  assert.deepStrictEqual(rows[1], ['2026-04-07', 'EUR', 'SEK', 11.2, 'ECB']);
  assert.ok(book.sheets.rates.formats.some(([col, f]) => col === 1 && f === '@'));
  book.sheets.rates.grid[2][0] = new Date('2026-04-07T00:00:00Z');
  assert.deepStrictEqual(store.load(), { NOK: [{ date: '2026-04-02', rate: 11.5 }], SEK });
});

test('sync all with fx: balances get EUR columns before the raw write, stored transactions are converted and saved', () => {
  let rows = [txn({ transactionId: 'old', bookingDate: '2026-04-01', amount: 109 })];
  const raw = {};
  const { fx } = service();
  const outcome = BS.SyncAll.run({
    connections: [{ key: 'bank-x', holder: 'x', bank: { aspspName: 'Bank' } }],
    isConnected: () => true,
    providerFor: () => ({
      listAccounts: () => [{ provider: 'p', accountId: 'a1', connection: 'bank-x', currency: 'SEK' }],
      getBalances: (id) => [{ accountId: id, amount: 1120, currency: 'SEK' }],
      getTransactions: () => []
    }),
    store: { load: () => rows, save: (r) => { rows = r; return r.length; } },
    rawWriter: { writeTable: (tab, cols, r) => { raw[tab] = r; return r.length; } },
    nowMs: () => 0,
    todayIso: () => '2026-04-07',
    fx
  });
  assert.deepStrictEqual([raw.raw_balances[0].amountEur, raw.raw_balances[0].fxRate], [100, 11.2]);
  assert.deepStrictEqual([rows[0].amountEur, rows[0].fxRateDate], [10, '2026-04-01']);
  assert.deepStrictEqual(outcome.fx, { changed: 1, missing: 0, waiting: 0, balancesMissing: 0, failed: [] });
  assert.match(BS.SyncAll.summarize(outcome), /EUR conversion: 1 transactions converted\./);
  assert.ok(BS.RawWriter.BALANCE_COLUMNS.some((c) => c.key === 'amountEur' && c.numeric));
});

test('sync all with fx: an FX failure is reported and the rest of the sync still lands', () => {
  let rows = [txn({})];
  const broken = { convertBalances: () => { throw new Error('boom'); }, convertTransactions: () => { throw new Error('boom'); } };
  const raw = {};
  const outcome = BS.SyncAll.run({
    connections: [{ key: 'bank-x', holder: 'x', bank: { aspspName: 'Bank' } }],
    isConnected: () => true,
    providerFor: () => ({ listAccounts: () => [{ accountId: 'a1', connection: 'bank-x' }], getBalances: () => [], getTransactions: () => [] }),
    store: { load: () => rows, save: (r) => { rows = r; return r.length; } },
    rawWriter: { writeTable: (tab, cols, r) => { raw[tab] = r; return r.length; } },
    nowMs: () => 0,
    todayIso: () => '2026-04-07',
    fx: broken
  });
  assert.strictEqual(outcome.results[0].status, 'ok');
  assert.ok(raw.raw_accounts);
  assert.match(BS.SyncAll.summarize(outcome), /EUR conversion failed \(Error: boom\)/);
});

test('fx service: if saving the rates fails, the range is fetched again next run', () => {
  const cache = fakeCache();
  const failing = { load: () => ({}), save: () => { throw new Error('tab locked'); } };
  const first = service({ cache, rates: failing });
  assert.throws(() => first.fx.convertTransactions([txn({})]), /tab locked/);
  const second = service({ cache });
  second.fx.convertTransactions([txn({})]);
  assert.strictEqual(second.calls.length, 1);
});
