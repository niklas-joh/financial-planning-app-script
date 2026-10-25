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

test('t212 normalize: cash and investment figures become balances tagged for the daily value', () => {
  const b = BS.T212Normalize.balances(SUMMARY, 'acc', '2026-11-02T06:30:00.000Z');
  assert.deepStrictEqual(b.map((x) => [x.balanceType, x.valueKind, x.amount]), [
    ['cash.availableToTrade', 'cash', 100.5],
    ['cash.inPies', 'cash', 20],
    ['cash.reservedForOrders', 'cash', 4.5],
    ['investments.currentValue', 'investments', 875],
    ['totalValue', '', 1000]
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
  const env = { http: () => ({ request: () => SUMMARY }), base64: (s) => s, toLocalDate, msLeft: () => 1000 };
  const t212 = R.create(R.connections(wired)[1], wired, env);
  assert.strictEqual(t212.listAccounts()[0].accountId, '1234567');
  assert.throws(() => R.entryFor({ key: 'x', provider: 'nope' }), /Unknown provider: nope/);
});
