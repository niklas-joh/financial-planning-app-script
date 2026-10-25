// Trading 212 provider: normalization, paging, auth header, rate-limit retries, credentials
// and the provider registry. All API responses below are made up.
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');

const T212 = BS.Config.TRADING212;
const NOW = Date.parse('2026-11-02T06:30:00Z');
const toLocalDate = (iso) => new Date(iso).toISOString().slice(0, 10);

const SUMMARY = {
  id: 1234567,
  currency: 'EUR',
  cash: { availableToTrade: 100.5, inPies: 20, reservedForOrders: 4.5 },
  investments: { currentValue: 875, totalCost: 800, realizedProfitLoss: 12, unrealizedProfitLoss: 75 },
  totalValue: 1000
};
const POSITION = {
  instrument: { ticker: 'DEMO_US_EQ', isin: 'US0000000001', name: 'Demo Corp', currency: 'USD' },
  quantity: 3, averagePricePaid: 90, currentPrice: 100, createdAt: '2025-01-02T10:00:00Z',
  walletImpact: { currency: 'EUR', currentValue: 270, totalCost: 250, unrealizedProfitLoss: 20, fxImpact: -1.5 }
};

const fakeProps = (init) => {
  const map = Object.assign({}, init);
  return {
    getProperty: (k) => (k in map ? map[k] : null),
    setProperty: (k, v) => { map[k] = v; },
    getKeys: () => Object.keys(map)
  };
};

/** A client whose history endpoints serve the given pages in order, recording the paths asked for. */
function fakeClient(over) {
  const calls = { summary: 0, positions: 0, transactions: [], dividends: [] };
  const pager = (pages, log) => (path) => {
    log.push(path);
    return pages[log.length - 1] || { items: [], nextPagePath: null };
  };
  const o = over || {};
  return {
    calls,
    summary: () => { calls.summary++; return o.summary || SUMMARY; },
    positions: () => { calls.positions++; return o.positions || [POSITION]; },
    transactionsPage: pager(o.cashPages || [], calls.transactions),
    dividendsPage: pager(o.dividendPages || [], calls.dividends)
  };
}

// Built by the real module from property names, as in production.
const CONN = BS.T212Credentials.connections(fakeProps({ T212_API_KEY_ALEX: 'x' }))[0];

const provider = (client) => BS.Trading212Provider.create({ connection: CONN, client, nowMs: () => NOW, toLocalDate });

test('t212 normalize: the summary becomes one account named by bank and currency, id kept as is', () => {
  const a = BS.T212Normalize.account(SUMMARY, CONN);
  assert.deepStrictEqual([a.provider, a.accountId, a.bank, a.connection, a.accountName, a.currency],
    ['trading212', '1234567', 'Trading 212', 'trading212-alex', 'Trading 212 EUR', 'EUR']);
  assert.deepStrictEqual(a.apiFields, { id: 1234567, currency: 'EUR' });
  assert.throws(() => BS.T212Normalize.account({ currency: 'EUR' }, CONN), /missing id or currency/);
});

test('t212 normalize: summary figures become balances; total and investments are tagged for the daily value', () => {
  const b = BS.T212Normalize.balances(SUMMARY, 'acc', '2026-11-02T06:30:00.000Z');
  // The daily value takes Trading 212's total and investments value (which already contains the cash in
  // pies); the cash rows are detail only, so pie cash is never counted twice.
  assert.deepStrictEqual(b.map((x) => [x.balanceType, x.valueKind, x.amount]), [
    ['cash.availableToTrade', '', 100.5],
    ['cash.inPies', '', 20],
    ['cash.reservedForOrders', '', 4.5],
    ['investments.currentValue', 'investments', 875],
    ['totalValue', 'total', 1000]
  ]);
  assert.ok(b.every((x) => x.currency === 'EUR' && x.provider === 'trading212' && x.accountId === 'acc'));
  const partial = BS.T212Normalize.balances({ currency: 'EUR', cash: { availableToTrade: 1 } }, 'acc', '');
  assert.deepStrictEqual(partial.map((x) => x.balanceType), ['cash.availableToTrade'], 'absent figures are skipped');
  assert.throws(() => BS.T212Normalize.balances({ currency: 'EUR', totalValue: 'n/a' }, 'acc', ''), /not numeric/);
});

test('t212 normalize: positions become holdings with value in the account currency and prices in the instrument currency', () => {
  const [h] = BS.T212Normalize.holdings([POSITION], 'acc', 'EUR', 'now');
  assert.deepStrictEqual(
    [h.ticker, h.isin, h.name, h.instrumentCurrency, h.quantity, h.averagePrice, h.currentPrice, h.amount, h.currency, h.totalCost, h.unrealizedPnl, h.fxImpact, h.openedAt, h.asOf],
    ['DEMO_US_EQ', 'US0000000001', 'Demo Corp', 'USD', 3, 90, 100, 270, 'EUR', 250, 20, -1.5, '2025-01-02T10:00:00Z', 'now']);
  assert.strictEqual(h.apiFields['instrument.ticker'], 'DEMO_US_EQ');
  const noFx = BS.T212Normalize.holdings([Object.assign({}, POSITION, { walletImpact: { currentValue: 5, fxImpact: null } })], 'acc', 'EUR', 'now')[0];
  assert.deepStrictEqual([noFx.fxImpact, noFx.currency], [null, 'EUR'], 'null fx impact and missing currency fall back');
  assert.throws(() => BS.T212Normalize.holdings([{ quantity: 1 }], 'acc', 'EUR', 'now'), /ticker/);
});

test('t212 normalize: cash movement signs follow the type; transfers and unknown types keep the API sign', () => {
  const sign = (type, amount) => BS.T212Normalize.cashMovement({ type, amount, currency: 'EUR', dateTime: '2026-10-30T23:30:00Z', reference: 'r' }, 'acc', toLocalDate).amount;
  assert.deepStrictEqual(
    [sign('WITHDRAW', 10), sign('WITHDRAW', -10), sign('FEE', 1), sign('DEPOSIT', -50), sign('INTEREST_ON_FREE_CASH', 0.2),
      sign('LENDING_INTEREST', 0.1), sign('TRANSFER', -7), sign('TRANSFER', 7), sign('SOMETHING_NEW', -3)],
    [-10, -10, -1, 50, 0.2, 0.1, -7, 7, -3]);
  const t = BS.T212Normalize.cashMovement({ type: 'DEPOSIT', amount: 50, currency: 'EUR', dateTime: '2026-10-30T23:30:00Z', reference: 'ref-1' }, 'acc', toLocalDate);
  assert.deepStrictEqual([t.transactionId, t.status, t.bookingDate, t.valueDate, t.description, t.counterparty, t.amountEur],
    ['ref-1', 'booked', '2026-10-30', '2026-10-30', 'DEPOSIT', '', null]);
  assert.strictEqual(t.apiFields.type, 'DEPOSIT');
  const noRef = (amount) => BS.T212Normalize.cashMovement({ type: 'FEE', amount, dateTime: '2026-10-30T10:00:00Z' }, 'acc', toLocalDate).transactionId;
  assert.match(noRef(1), /^h:[0-9a-f]{8}$/);
  assert.strictEqual(noRef(1), noRef(1), 'the fallback id is stable');
  assert.throws(() => BS.T212Normalize.cashMovement({ type: 'FEE', amount: 1, dateTime: 'nope' }, 'acc', toLocalDate), /no valid date/);
});

test('t212 normalize: dividends are positive, prefixed and named after the instrument', () => {
  const d = BS.T212Normalize.dividend({ amount: 1.23, currency: 'EUR', paidOn: '2026-10-15T08:00:00Z', reference: '99', ticker: 'DEMO_US_EQ', instrument: { name: 'Demo Corp' }, type: 'ORDINARY' }, 'acc', toLocalDate);
  assert.deepStrictEqual([d.transactionId, d.amount, d.bookingDate, d.counterparty, d.description], ['div:99', 1.23, '2026-10-15', 'Demo Corp', 'Dividend ORDINARY']);
  const bare = BS.T212Normalize.dividend({ amount: -2, paidOn: '2026-10-15T08:00:00Z', reference: '1', ticker: 'X_EQ' }, 'acc', toLocalDate);
  assert.deepStrictEqual([bare.amount, bare.counterparty, bare.description], [2, 'X_EQ', 'Dividend']);
});

test('t212 provider: reads the summary once for the account, balances and holdings', () => {
  const client = fakeClient();
  const p = provider(client);
  const [account] = p.listAccounts({ outOfTime: () => false });
  assert.strictEqual(account.accountId, '1234567');
  assert.strictEqual(p.getBalances(account.accountId).length, 5);
  const holdings = p.getHoldings(account.accountId);
  assert.deepStrictEqual([holdings.length, holdings[0].currency, holdings[0].asOf], [1, 'EUR', new Date(NOW).toISOString()]);
  assert.deepStrictEqual([client.calls.summary, client.calls.positions], [1, 1]);
  assert.strictEqual(p.maxWindowDays, Infinity);
  assert.throws(() => p.authorize(), /T212_API_KEY_<HOLDER>/);
});

test('t212 provider: history pages are followed until one reaches the start date', () => {
  const item = (day, ref, type) => ({ type: type || 'DEPOSIT', amount: 10, currency: 'EUR', dateTime: day + 'T12:00:00Z', reference: ref });
  const client = fakeClient({
    cashPages: [
      { items: [item('2026-11-01', 'c1'), item('2026-10-20', 'c2')], nextPagePath: '/api/v0/equity/history/transactions?cursor=2&limit=50' },
      { items: [item('2026-10-05', 'c3', 'WITHDRAW'), item('2026-09-28', 'c4')], nextPagePath: '/api/v0/equity/history/transactions?cursor=4&limit=50' },
      { items: [item('2026-09-01', 'never-read')], nextPagePath: null }
    ],
    dividendPages: [{ items: [{ amount: 1, currency: 'EUR', paidOn: '2026-10-10T08:00:00Z', reference: 'd1', ticker: 'X_EQ' }], nextPagePath: null }]
  });
  const got = provider(client).getTransactions('1234567', { dateFrom: '2026-10-01', dateTo: '2026-11-02' });
  assert.deepStrictEqual(client.calls.transactions, [null, '/api/v0/equity/history/transactions?cursor=2&limit=50'],
    'stops after the page that reaches back past dateFrom');
  assert.deepStrictEqual(client.calls.dividends, [null]);
  assert.deepStrictEqual(got.map((t) => [t.transactionId, t.amount]), [['c1', 10], ['c2', 10], ['c3', -10], ['div:d1', 1]],
    'items before dateFrom are left out');
});

test('t212 provider: a history longer than the page cap is an error, not a silent gap', () => {
  const endless = { items: [{ type: 'DEPOSIT', amount: 1, dateTime: '2026-11-01T00:00:00Z', reference: 'x' }], nextPagePath: '/api/v0/next' };
  const pages = Array(BS.Config.TRANSACTIONS.MAX_PAGES + 1).fill(endless);
  assert.throws(() => provider(fakeClient({ cashPages: pages })).getTransactions('a', { dateFrom: '2026-01-01', dateTo: '2026-11-02' }),
    /did not reach 2026-01-01/);
});

test('t212 client: Basic auth from key and secret, paths on the live origin only, first page at the maximum size', () => {
  const requests = [];
  const http = { request: (req) => { requests.push(req); return { items: [], nextPagePath: null }; } };
  const client = BS.T212Client.create(http, { key: 'fake-key', secret: 'fake-secret' }, (s) => Buffer.from(s, 'utf8').toString('base64'));
  client.summary();
  client.transactionsPage(null);
  client.dividendsPage('/api/v0/equity/history/dividends?cursor=5&limit=50');
  assert.deepStrictEqual(requests.map((r) => r.url), [
    'https://live.trading212.com/api/v0/equity/account/summary',
    'https://live.trading212.com/api/v0/equity/history/transactions?limit=' + T212.PAGE_LIMIT,
    'https://live.trading212.com/api/v0/equity/history/dividends?cursor=5&limit=50'
  ]);
  assert.strictEqual(requests[0].headers.Authorization, 'Basic ' + Buffer.from('fake-key:fake-secret').toString('base64'));
  assert.ok(requests.every((r) => r.method === 'GET'));
  ['https://evil.example/x', '//evil.example/x', 'relative'].forEach((p) => assert.throws(() => BS.T212Client.urlFor(p), /page path/));
});

test('t212 rate limit: a 429 waits until the reset header, in seconds or ms, and gives up past the deadline or the cap', () => {
  const now = 1000000000000;
  const d = BS.T212RateLimit.delayMs;
  assert.strictEqual(d(429, { 'X-RateLimit-Reset': String((now + 4000) / 1000) }, 0, now, 60000), 4250);
  assert.strictEqual(d(429, { 'x-ratelimit-reset': now + 1000 }, 0, now, 60000), 1250);
  assert.strictEqual(d(429, { 'x-ratelimit-reset': (now - 5000) / 1000 }, 0, now, 60000), 250, 'a reset in the past retries at once');
  assert.strictEqual(d(429, {}, 2, now, 60000), BS.Http.backoffMs(2), 'no header: standard backoff');
  assert.strictEqual(d(503, { 'x-ratelimit-reset': (now + 30000) / 1000 }, 1, now, 60000), BS.Http.backoffMs(1), '5xx ignores the reset');
  assert.strictEqual(d(429, { 'x-ratelimit-reset': (now + 30000) / 1000 }, 0, now, 10000), null, 'past the run deadline');
  assert.strictEqual(d(429, { 'x-ratelimit-reset': (now + T212.MAX_RETRY_WAIT_MS + 1000) / 1000 }, 0, now, Infinity), null, 'past the cap');
});

test('t212 rate limit: the hook makes Http wait for the reset, then succeed', () => {
  const sleeps = [];
  let calls = 0;
  const now = 1000000000000;
  const http = BS.Http.create({
    fetch: () => {
      calls++;
      return calls === 1
        ? { getResponseCode: () => 429, getContentText: () => '', getAllHeaders: () => ({ 'x-ratelimit-reset': String((now + 5000) / 1000) }) }
        : { getResponseCode: () => 200, getContentText: () => '{"ok":true}', getAllHeaders: () => ({}) };
    },
    sleep: (ms) => sleeps.push(ms),
    retryDelayMs: BS.T212RateLimit.hook(() => now, () => 60000)
  });
  assert.deepStrictEqual(http.request({ method: 'GET', url: 'u' }), { ok: true });
  assert.deepStrictEqual(sleeps, [5250]);
});

test('t212 credentials: one connection per holder, status shows set/MISSING only, never values', () => {
  const props = fakeProps({
    T212_API_KEY_ALEX: 'k1', T212_API_SECRET_ALEX: 's1',
    T212_API_KEY_SAM_ISA: 'k2',
    EB_APP_ID: 'x'
  });
  const conns = BS.T212Credentials.connections(props);
  assert.deepStrictEqual(conns.map((c) => [c.key, c.holder, c.provider, c.propertySuffix]),
    [['trading212-alex', 'alex', 'trading212', 'ALEX'], ['trading212-sam-isa', 'sam-isa', 'trading212', 'SAM_ISA']]);
  assert.deepStrictEqual(conns.map((c) => BS.T212Credentials.isComplete(props, c)), [true, false]);
  assert.deepStrictEqual(BS.T212Credentials.credentials(props, conns[0]), { key: 'k1', secret: 's1' });
  assert.throws(() => BS.T212Credentials.credentials(props, conns[1]), /Set T212_API_SECRET_SAM_ISA/);
  const lines = BS.T212Credentials.statusLines(props);
  assert.deepStrictEqual(lines, [
    'Trading 212 (alex) - T212_API_KEY_ALEX: set; T212_API_SECRET_ALEX: set',
    'Trading 212 (sam-isa) - T212_API_KEY_SAM_ISA: set; T212_API_SECRET_SAM_ISA: MISSING'
  ]);
  assert.ok(!lines.join().includes('k1') && !lines.join().includes('s1'));
  assert.deepStrictEqual(BS.T212Credentials.statusLines(fakeProps({})), ['Trading 212: no API key set (optional; see the README)']);
});

test('provider registry: lists both providers, routes each connection, keeps consents to Enable Banking', () => {
  const eb = { key: 'demo-bank-alex', holder: 'alex', bank: { aspspName: 'Demo Bank', country: 'LU' } };
  const props = fakeProps({ T212_API_KEY_ALEX: 'k', T212_API_SECRET_ALEX: 's', EB_APP_ID: 'x' });
  const wired = {
    props,
    connections: { all: () => [eb] },
    store: { loadSession: (k) => (k === eb.key ? { accessValidUntil: '2027-01-31T00:00:00Z' } : null) },
    nowMs: () => NOW
  };
  const R = BS.ProviderRegistry;
  assert.deepStrictEqual(R.connections(wired).map((c) => c.key), ['demo-bank-alex', 'trading212-alex']);
  assert.deepStrictEqual(R.connections(wired).map((c) => R.isConnected(c, wired)), [true, true]);
  assert.deepStrictEqual(R.consents(wired).map((c) => c.connection.key), ['demo-bank-alex']);
  const lines = R.statusLines(wired);
  assert.deepStrictEqual(lines.slice(0, 3), ['EB_APP_ID: set', 'EB_PRIVATE_KEY: MISSING', 'EB_REDIRECT_URL: MISSING']);
  assert.match(lines[3], /^Demo Bank \(alex\): 2027-01-31T00:00:00Z \(89 days left\)/);
  assert.strictEqual(lines[4], 'Trading 212 (alex) - T212_API_KEY_ALEX: set; T212_API_SECRET_ALEX: set');
  assert.strictEqual(lines[5], 'Trading 212 (alex) history: older history loads with the next sync');
  const env = { http: () => ({ request: () => SUMMARY }), base64: (s) => s, toLocalDate, msLeft: () => 1000 };
  const t212 = R.create(R.connections(wired)[1], wired, env);
  assert.strictEqual(t212.listAccounts()[0].accountId, '1234567');
  assert.throws(() => R.entryFor({ key: 'x', provider: 'nope' }), /Unknown provider: nope/);
});

// --- One-time history backfill ---

const pages = (n, prefix) => Array.from({ length: n }, (_, i) => ({
  items: [{ type: 'DEPOSIT', amount: 1, currency: 'EUR', dateTime: '2024-0' + (9 - Math.min(i, 8)) + '-01T10:00:00Z', reference: prefix + i }],
  nextPagePath: i < n - 1 ? '/api/v0/page-' + prefix + (i + 1) : null
}));
const reader = (list) => {
  const asked = [];
  const read = (path) => {
    asked.push(path);
    const i = path === null ? 0 : Number(path.split('-').pop().replace(/\D/g, ''));
    return list[i];
  };
  return { read, asked };
};

test('t212 history: a step reads up to the page budget from the saved position and marks the last page', () => {
  const r = reader(pages(5, 'c'));
  const first = BS.T212History.step(r.read, { next: null, done: false }, 3);
  assert.deepStrictEqual([first.items.length, first.feed], [3, { next: '/api/v0/page-c3', done: false }]);
  const second = BS.T212History.step(r.read, first.feed, 3);
  assert.deepStrictEqual([second.items.length, second.feed], [2, { next: null, done: true }]);
  assert.deepStrictEqual(r.asked, [null, '/api/v0/page-c1', '/api/v0/page-c2', '/api/v0/page-c3', '/api/v0/page-c4']);
  assert.deepStrictEqual(BS.T212History.step(() => { throw new Error('never'); }, second.feed, 3).items, [], 'a done feed reads nothing');
});

test('t212 history: a rate limit, server error or the time budget pauses at the last good page; other errors throw', () => {
  const list = pages(3, 'c');
  const failingOn = (n, error) => (path) => { if (path === '/api/v0/page-c' + n) throw error; return reader(list).read(path); };
  const paused = BS.T212History.step(failingOn(1, BS.Errors.http(429, 'Limited')), { next: null, done: false }, 10);
  assert.deepStrictEqual([paused.items.length, paused.feed], [1, { next: '/api/v0/page-c1', done: false }]);
  assert.strictEqual(BS.T212History.step(failingOn(1, BS.Errors.http(503, 'down')), { next: null, done: false }, 10).feed.next, '/api/v0/page-c1');
  assert.strictEqual(BS.T212History.step(failingOn(1, BS.Errors.create('TimeBudget', 't')), { next: null, done: false }, 10).feed.done, false);
  assert.throws(() => BS.T212History.step(failingOn(1, BS.Errors.create('DataError', 'bad')), { next: null, done: false }, 10), /bad/);
});

test('t212 history: the state lives in one Script Property per holder and an unreadable one restarts', () => {
  const props = fakeProps({});
  const st = BS.T212History.store(props, CONN);
  assert.deepStrictEqual(st.load(), { transactions: { next: null, done: false }, dividends: { next: null, done: false } });
  assert.strictEqual(BS.T212History.statusLine(props, CONN), 'Trading 212 (alex) history: older history loads with the next sync');
  st.save({ transactions: { next: '/api/v0/x', done: false }, dividends: { next: null, done: true } });
  assert.ok(props.getKeys().includes('T212_HISTORY_ALEX'));
  assert.match(BS.T212History.statusLine(props, CONN), /loading older history/);
  st.save({ transactions: { next: null, done: true }, dividends: { next: null, done: true } });
  assert.strictEqual(BS.T212History.statusLine(props, CONN), 'Trading 212 (alex) history: complete');
  props.setProperty('T212_HISTORY_ALEX', '{not json');
  assert.deepStrictEqual(st.load().transactions, { next: null, done: false });
});

test('t212 provider: the backfill adds older history over several runs, then stops reading it', () => {
  const props = fakeProps({});
  const history = BS.T212History.store(props, CONN);
  const cashList = pages(BS.Config.TRADING212.BACKFILL_PAGES_PER_RUN + 2, 'c');
  const cash = reader(cashList);
  const client = {
    summary: () => SUMMARY,
    positions: () => [],
    transactionsPage: cash.read,
    dividendsPage: () => ({ items: [{ amount: 1, paidOn: '2024-10-01T10:00:00Z', reference: 'd1', ticker: 'X_EQ' }], nextPagePath: null })
  };
  const run = () => BS.Trading212Provider.create({ connection: CONN, client, nowMs: () => NOW, toLocalDate, history })
    .getTransactions('acc', { dateFrom: '2026-10-01', dateTo: '2026-11-02' });
  const first = run();
  const perRun = BS.Config.TRADING212.BACKFILL_PAGES_PER_RUN;
  assert.strictEqual(first.filter((t) => t.transactionId.startsWith('c')).length, perRun, 'old items come in although they are before dateFrom');
  assert.ok(first.some((t) => t.transactionId === 'div:d1'));
  assert.strictEqual(new Set(first.map((t) => t.transactionId)).size, first.length, 'no duplicates between recent and backfilled');
  assert.deepStrictEqual(history.load(), { transactions: { next: '/api/v0/page-c' + perRun, done: false }, dividends: { next: null, done: true } });
  const second = run();
  assert.deepStrictEqual(second.filter((t) => t.transactionId.startsWith('c')).map((t) => t.transactionId), ['c' + perRun, 'c' + (perRun + 1)],
    'the second run continues where the first stopped (page 0 is older than dateFrom, so the incremental part skips it)');
  assert.strictEqual(history.load().transactions.done, true);
  cash.asked.length = 0;
  run();
  assert.deepStrictEqual(cash.asked, [null], 'once done, only the incremental first page is read');
});

test('t212 provider: an item that fails to normalize never moves the backfill position past it', () => {
  const props = fakeProps({});
  const history = BS.T212History.store(props, CONN);
  const client = {
    summary: () => SUMMARY,
    positions: () => [],
    transactionsPage: (path) => (path === null
      ? { items: [{ type: 'DEPOSIT', amount: 1, dateTime: '2026-11-01T10:00:00Z', reference: 'ok' }], nextPagePath: '/api/v0/p1' }
      : { items: [{ type: 'DEPOSIT', amount: 'x', dateTime: '2024-01-01T10:00:00Z', reference: 'bad' }], nextPagePath: null }),
    dividendsPage: () => ({ items: [], nextPagePath: null })
  };
  const provider = BS.Trading212Provider.create({ connection: CONN, client, nowMs: () => NOW, toLocalDate, history });
  assert.throws(() => provider.getTransactions('acc', { dateFrom: '2026-10-01', dateTo: '2026-11-02' }), /not numeric/);
  assert.deepStrictEqual(history.load().transactions, { next: null, done: false });
});

test('t212 provider: when a later feed fails, no feed position is saved, so the read pages are not lost', () => {
  const props = fakeProps({});
  const history = BS.T212History.store(props, CONN);
  const client = {
    summary: () => SUMMARY,
    positions: () => [],
    transactionsPage: (path) => (path === null
      ? { items: [{ type: 'DEPOSIT', amount: 1, dateTime: '2026-11-01T10:00:00Z', reference: 'new' }], nextPagePath: '/api/v0/p1' }
      : { items: [{ type: 'DEPOSIT', amount: 1, dateTime: '2024-01-01T10:00:00Z', reference: 'old' }], nextPagePath: null }),
    dividendsPage: () => ({ items: [{ amount: 'x', paidOn: '2024-01-01T10:00:00Z', reference: 'bad' }], nextPagePath: null })
  };
  const provider = BS.Trading212Provider.create({ connection: CONN, client, nowMs: () => NOW, toLocalDate, history });
  assert.throws(() => provider.getTransactions('acc', { dateFrom: '2026-10-01', dateTo: '2026-11-02' }), /not numeric/);
  assert.strictEqual(props.getProperty('T212_HISTORY_ALEX'), null, 'the transactions feed did not move either');
  assert.deepStrictEqual(history.load().transactions, { next: null, done: false });
});

test('t212 provider: a run that reads no new backfill pages writes no state', () => {
  const props = fakeProps({});
  const history = BS.T212History.store(props, CONN);
  history.save({ transactions: { next: null, done: true }, dividends: { next: null, done: true } });
  let writes = 0;
  const counting = { load: history.load, save: (st) => { writes++; history.save(st); } };
  const client = { summary: () => SUMMARY, positions: () => [], transactionsPage: () => ({ items: [], nextPagePath: null }), dividendsPage: () => ({ items: [], nextPagePath: null }) };
  BS.Trading212Provider.create({ connection: CONN, client, nowMs: () => NOW, toLocalDate, history: counting })
    .getTransactions('acc', { dateFrom: '2026-10-01', dateTo: '2026-11-02' });
  assert.strictEqual(writes, 0);
});
