// Daily values (cash vs investments per account per day) and the end-to-end Sync all path with a
// Trading 212 provider next to a bank. All data below is made up.
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');
const { fakeSheetBook } = require('./fake-sheet.js');

const bal = (o) => Object.assign({ provider: 'trading212', connection: 'trading212-alex', accountId: 'acc', currency: 'SEK', fetchedAt: 't' }, o);

test('daily values: sums cash and investments per account; balances without a kind are ignored', () => {
  const rows = BS.DailyValues.fromBalances([
    bal({ valueKind: 'cash', amount: 100.1, amountEur: 10.01, fxRate: 10, fxRateDate: '2026-11-01' }),
    bal({ valueKind: 'cash', amount: 0.2, amountEur: 0.02, fxRate: 10, fxRateDate: '2026-11-01' }),
    bal({ valueKind: 'investments', amount: 900, amountEur: 90, fxRate: 10, fxRateDate: '2026-11-01' }),
    bal({ valueKind: '', amount: 5000, amountEur: 500 }),
    bal({ provider: 'enablebanking', connection: 'bank-alex', accountId: 'b1', amount: 50, amountEur: 5 })
  ], '2026-11-02');
  assert.strictEqual(rows.length, 1);
  const r = rows[0];
  assert.deepStrictEqual([r.date, r.connection, r.accountId, r.currency, r.cash, r.investments, r.total],
    ['2026-11-02', 'trading212-alex', 'acc', 'SEK', 100.3, 900, 1000.3]);
  assert.deepStrictEqual([r.cashEur, r.investmentsEur, r.totalEur, r.fxRate, r.fxRateDate], [10.03, 90, 100.03, 10, '2026-11-01']);
});

test('daily values: an account with a balance lacking a rate gets no partial EUR sum', () => {
  const [r] = BS.DailyValues.fromBalances([
    bal({ valueKind: 'cash', amount: 1, amountEur: 0.1, fxRate: 10 }),
    bal({ valueKind: 'investments', amount: 9 })
  ], '2026-11-02');
  assert.deepStrictEqual([r.total, r.cashEur, r.investmentsEur, r.totalEur, r.fxRate], [10, null, null, null, null]);
});

test('daily values tab: a second run on the same day replaces that day, the next day adds a row, newest first', () => {
  const book = fakeSheetBook();
  const writer = BS.RawWriter.create(book);
  const write = (date, cash) => writer.writeTable('daily_values', BS.RawWriter.DAILY_VALUE_COLUMNS,
    BS.DailyValues.fromBalances([bal({ valueKind: 'cash', amount: cash })], date), ['trading212-alex'], BS.DailyValues.accumulate());
  write('2026-11-01', 1);
  write('2026-11-01', 2);
  const n = write('2026-11-02', 3);
  assert.strictEqual(n, 2);
  const [header, ...rows] = book.sheets.daily_values.grid;
  const col = (k) => header.indexOf(k);
  assert.deepStrictEqual(rows.map((r) => [r[col('date')], r[col('cash')]]), [['2026-11-02', 3], ['2026-11-01', 2]]);
});

// --- Sync all with a Trading 212 provider and a bank ---

const SUMMARY = { id: 42, currency: 'EUR', cash: { availableToTrade: 60, inPies: 0, reservedForOrders: 0 }, investments: { currentValue: 140 }, totalValue: 200 };
const POSITION = {
  instrument: { ticker: 'DEMO_US_EQ', name: 'Demo Corp', currency: 'USD' }, quantity: 2, averagePricePaid: 60, currentPrice: 70,
  walletImpact: { currency: 'EUR', currentValue: 140, totalCost: 120, unrealizedProfitLoss: 20 }
};
const CASH = { items: [{ type: 'DEPOSIT', amount: 100, currency: 'EUR', dateTime: '2026-10-20T09:00:00Z', reference: 'r1' }], nextPagePath: null };
const DIVIDENDS = { items: [{ amount: 1.5, currency: 'EUR', paidOn: '2026-10-25T09:00:00Z', reference: 'd1', ticker: 'DEMO_US_EQ' }], nextPagePath: null };

const T212_CONN = BS.T212Credentials.connections({ getKeys: () => ['T212_API_KEY_ALEX'] })[0];
const BANK_CONN = { key: 'demo-bank-alex', holder: 'alex', bank: { aspspName: 'Demo Bank', country: 'LU' } };

function t212Provider(failing) {
  return BS.Trading212Provider.create({
    connection: T212_CONN,
    nowMs: () => Date.parse('2026-11-02T06:30:00Z'),
    toLocalDate: (iso) => new Date(iso).toISOString().slice(0, 10),
    client: {
      summary: () => { if (failing) throw BS.Errors.http(401, 'Bad API key'); return SUMMARY; },
      positions: () => [POSITION],
      transactionsPage: () => CASH,
      dividendsPage: () => DIVIDENDS
    }
  });
}

const bankProvider = (failing) => ({
  listAccounts: () => {
    if (failing) throw BS.Errors.auth('Consent has expired; re-authorize');
    return [{ provider: 'enablebanking', accountId: 'b1', bank: 'Demo Bank', connection: BANK_CONN.key, currency: 'EUR', accountName: 'Main' }];
  },
  getBalances: (id) => [{ provider: 'enablebanking', accountId: id, balanceType: 'CLBD', amount: 5, currency: 'EUR' }],
  getTransactions: () => []
});

function runSync(book, store, opts) {
  const o = opts || {};
  return BS.SyncAll.run({
    connections: [BANK_CONN, T212_CONN],
    isConnected: () => true,
    providerFor: (c) => (c.provider === 'trading212' ? t212Provider(o.t212Fails) : bankProvider(o.bankFails)),
    store,
    rawWriter: BS.RawWriter.create(book),
    nowMs: () => Date.parse('2026-11-02T06:30:00Z'),
    todayIso: () => o.today || '2026-11-02',
    fx: {
      convertBalances: (rows) => ({ rows: rows.map((r) => Object.assign({}, r, { amountEur: r.amount, fxRate: 1, fxRateDate: '2026-11-02' })), missing: 0, failed: [] }),
      convertTransactions: (rows) => ({ rows, changed: 0, missing: 0, waiting: 0, failed: [] })
    }
  });
}

const memoryStore = () => {
  let rows = [];
  return { load: () => rows, save: (r) => { rows = r; return r.length; }, rows: () => rows };
};
const tab = (book, name) => {
  const [header, ...rows] = book.sheets[name].grid;
  return rows.map((r) => Object.fromEntries(header.map((h, i) => [h, r[i]])));
};

test('sync all: Trading 212 fills holdings, daily values and transactions next to the bank, idempotently', () => {
  const book = fakeSheetBook();
  const store = memoryStore();
  const first = runSync(book, store);
  assert.deepStrictEqual(first.results.map((r) => [r.connection, r.status, r.transactions.added]),
    [['demo-bank-alex', 'ok', 0], ['trading212-alex', 'ok', 2]]);
  const holdings = tab(book, 'holdings');
  assert.deepStrictEqual(holdings.map((h) => [h.connection, h.accountName, h.ticker, h.amount, h.amountEur]),
    [['trading212-alex', 'Trading 212 EUR', 'DEMO_US_EQ', 140, 140]]);
  const daily = tab(book, 'daily_values');
  assert.deepStrictEqual(daily.map((d) => [d.connection, d.accountId, d.cash, d.investments, d.total, d.totalEur]),
    [['trading212-alex', '42', 60, 140, 200, 200]], 'only accounts with tagged balances get a daily value');
  assert.deepStrictEqual(store.rows().map((t) => [t.connection, t.transactionId, t.amount, t.accountName]),
    [['trading212-alex', 'div:d1', 1.5, 'Trading 212 EUR'], ['trading212-alex', 'r1', 100, 'Trading 212 EUR']]);
  assert.ok(tab(book, 'raw_balances').some((b) => b.connection === 'trading212-alex' && b.valueKind === 'cash'));
  assert.match(BS.SyncAll.summarize(first), /1 holdings, 1 daily values\./);

  const second = runSync(book, store);
  assert.deepStrictEqual(second.results[1].transactions, { added: 0, updated: 0, removed: 0, complete: true });
  assert.strictEqual(tab(book, 'daily_values').length, 1, 'same day: replaced, not added');
  runSync(book, store, { today: '2026-11-03' });
  assert.deepStrictEqual(tab(book, 'daily_values').map((d) => d.date), ['2026-11-03', '2026-11-02']);
});

test('sync all: a failing Trading 212 key never blocks the bank, and a failing bank never blocks Trading 212', () => {
  const book = fakeSheetBook();
  const a = runSync(book, memoryStore(), { t212Fails: true });
  assert.deepStrictEqual(a.results.map((r) => r.status), ['ok', 'error']);
  assert.match(a.results[1].error, /HTTP 401/);
  assert.strictEqual(book.sheets.holdings, undefined, 'no holdings tab without holdings read');
  const b = runSync(fakeSheetBook(), memoryStore(), { bankFails: true });
  assert.deepStrictEqual(b.results.map((r) => r.status), ['error', 'ok']);
});

test('sync all: holdings of a connection that now holds nothing are cleared; other connections keep theirs', () => {
  const book = fakeSheetBook();
  const writer = BS.RawWriter.create(book);
  BS.RawSync.write(writer, { accounts: [], balances: [], holdings: [{ connection: 'trading212-sam', accountId: 's', ticker: 'KEEP_EQ' }, { connection: 'trading212-alex', accountId: 'a', ticker: 'SOLD_EQ' }], holdingConnections: ['trading212-sam', 'trading212-alex'] }, ['trading212-sam', 'trading212-alex'], '2026-11-02');
  BS.RawSync.write(writer, { accounts: [], balances: [], holdings: [], holdingConnections: ['trading212-alex'] }, ['trading212-alex'], '2026-11-02');
  assert.deepStrictEqual(tab(book, 'holdings').map((h) => h.ticker), ['KEEP_EQ']);
});
