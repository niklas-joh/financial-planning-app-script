// Pluxee provider: normalization, token refresh and rotation, the re-login error, credentials and
// the provider registry. All API responses below are made up (synthetic merchants and card ids).
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');

const PLUXEE = BS.Config.PLUXEE;
const toLocalDate = (iso) => new Date(iso).toISOString().slice(0, 10);

const fakeProps = (init) => {
  const map = Object.assign({}, init);
  return {
    map,
    getProperty: (k) => (k in map ? map[k] : null),
    setProperty: (k, v) => { map[k] = v; },
    getKeys: () => Object.keys(map)
  };
};

const CARD = { cardId: 'LUX0000000000000000001', maskedPan: 'XXXX 1234', name: 'Demo Lunch', state: 'ACTIVE' };
const money = (value) => ({ value, exponent: 2, currency: 'EUR' });
const ITEMS = [
  { id: 'LUXa1', description: 'Corner Deli', date: '2026-10-25T09:42:04.000', status: 'APPROVED', type: 'DEBIT',
    amount: money(-1250), merchantName: 'Corner Deli   Demo Town ', merchantCity: 'DEMO TOWN' },
  { id: 'LUXa2', description: 'Music Store', date: '2026-10-22T19:10:20.000', status: 'DECLINED', type: 'DEBIT',
    amount: money(-999), merchantName: 'MUSIC STORE' },
  { id: 'LUXa3', description: 'Monthly credit', date: '2026-10-01T00:05:00.000', status: 'APPROVED', type: 'CREDIT',
    amount: money(22000) },
  { id: 'LUXa4', description: 'Bakery', date: '2026-10-26T12:00:00.000', status: 'AUTHORIZED', type: 'DEBIT',
    amount: money(-480), merchantName: 'Bakery' }
];

const CONN = BS.PluxeeCredentials.connections(fakeProps({ PLUXEE_CARD_ALEX: 'x' }))[0];

test('pluxee credentials: one connection per holder suffix; complete only with card and token', () => {
  const props = fakeProps({ PLUXEE_CARD_ALEX: 'LUX1', PLUXEE_REFRESH_TOKEN_ALEX: 'secret-token', PLUXEE_CARD_SAM: 'LUX2' });
  const all = BS.PluxeeCredentials.connections(props);
  assert.deepStrictEqual(all.map((c) => [c.key, c.provider, c.propertySuffix]),
    [['pluxee-alex', 'pluxee', 'ALEX'], ['pluxee-sam', 'pluxee', 'SAM']]);
  assert.deepStrictEqual(all.map((c) => BS.PluxeeCredentials.isComplete(props, c)), [true, false]);
  const lines = BS.PluxeeCredentials.statusLines(props).join('\n');
  assert.match(lines, /PLUXEE_REFRESH_TOKEN_SAM: MISSING/);
  assert.ok(!lines.includes('secret-token'), 'never shows a value');
  assert.match(BS.PluxeeCredentials.statusLines(fakeProps({}))[0], /not set up/);
});

test('pluxee credentials: the card id must be letters and digits (it goes into the URL path)', () => {
  assert.strictEqual(BS.PluxeeCredentials.cardId(fakeProps({ PLUXEE_CARD_ALEX: ' LUX01 ' }), CONN), 'LUX01');
  assert.throws(() => BS.PluxeeCredentials.cardId(fakeProps({ PLUXEE_CARD_ALEX: '../x' }), CONN), /letters and digits/);
  assert.throws(() => BS.PluxeeCredentials.cardId(fakeProps({}), CONN), /Set PLUXEE_CARD_ALEX/);
});

test('pluxee normalize: amounts are minor units with an exponent', () => {
  assert.deepStrictEqual(BS.PluxeeNormalize.amountOf(money(-6787)), { amount: -67.87, currency: 'EUR' });
  assert.deepStrictEqual(BS.PluxeeNormalize.amountOf({ value: 5, exponent: 0, currency: 'EUR' }), { amount: 5, currency: 'EUR' });
  assert.throws(() => BS.PluxeeNormalize.amountOf({ value: 'n/a', exponent: 2 }), /not numeric/);
  assert.throws(() => BS.PluxeeNormalize.amountOf({ value: 1, exponent: -1 }), /exponent/);
});

test('pluxee normalize: zone-less local times keep their date; zoned ones go through the time zone', () => {
  assert.strictEqual(BS.PluxeeNormalize.localDate('2026-10-25T23:42:04.000', toLocalDate), '2026-10-25');
  assert.strictEqual(BS.PluxeeNormalize.localDate('2026-10-25T23:42:04Z', () => '2026-10-26'), '2026-10-26');
  assert.throws(() => BS.PluxeeNormalize.localDate('', toLocalDate), /no valid date/);
});

test('pluxee normalize: approved rows are booked card rows, declined ones dropped, others pending', () => {
  const rows = BS.PluxeeNormalize.transactions(ITEMS, 'LUX1', toLocalDate);
  assert.deepStrictEqual(rows.map((t) => [t.transactionId, t.bookingDate, t.status, t.amount, t.counterparty, t.category]), [
    ['LUXa1', '2026-10-25', 'booked', -12.5, 'Corner Deli Demo Town', 'DEBIT'],
    ['LUXa3', '2026-10-01', 'booked', 220, 'Monthly credit', 'CREDIT'],
    ['LUXa4', '2026-10-26', 'pending', -4.8, 'Bakery', 'DEBIT']
  ]);
  assert.ok(rows.every((t) => t.provider === 'pluxee' && t.classification === 'card' && t.currency === 'EUR'));
  assert.strictEqual(BS.TxnRouting.tabFor(rows[0]), BS.Config.TABS.TRANSACTIONS);
  assert.throws(() => BS.PluxeeNormalize.transaction({ status: 'APPROVED', amount: money(1), date: '2026-10-01' }, 'LUX1', toLocalDate), /no id/);
});

test('pluxee normalize: the card becomes one account named after the product', () => {
  const a = BS.PluxeeNormalize.account(CARD, 'LUXcfg', CONN);
  assert.deepStrictEqual([a.provider, a.accountId, a.bank, a.connection, a.accountName, a.currency, a.ibanMasked],
    ['pluxee', CARD.cardId, 'Pluxee', 'pluxee-alex', 'Demo Lunch', 'EUR', 'XXXX 1234']);
  assert.strictEqual(BS.PluxeeNormalize.account(null, 'LUXcfg', CONN).accountId, 'LUXcfg');
});

/** HTTP fake: discovery, token endpoint and transactions, recording every request. */
function fakeHttp(over) {
  const o = over || {};
  const calls = [];
  return {
    calls,
    request: (req) => {
      calls.push(req);
      if (req.url.endsWith('/.well-known/openid-configuration')) {
        return { token_endpoint: o.tokenEndpoint || PLUXEE.ISSUER + '/token' };
      }
      if (req.method === 'POST') {
        if (o.tokenError) throw o.tokenError;
        return o.tokenResponse || { access_token: 'at-1', refresh_token: 'rt-2', expires_in: 1800 };
      }
      return { card: CARD, transactions: ITEMS };
    }
  };
}

const tokenStore = (props) => BS.PluxeeCredentials.tokenStore(props, CONN);

test('pluxee client: refreshes once, saves a rotated refresh token, sends the bearer token', () => {
  const props = fakeProps({ PLUXEE_REFRESH_TOKEN_ALEX: 'rt-1' });
  const http = fakeHttp();
  const client = BS.PluxeeClient.create(http, tokenStore(props));
  client.transactions('LUX1');
  client.transactions('LUX1');
  const posts = http.calls.filter((c) => c.method === 'POST');
  assert.strictEqual(posts.length, 1, 'one refresh per run');
  assert.deepStrictEqual(posts[0].form, { grant_type: 'refresh_token', refresh_token: 'rt-1', client_id: PLUXEE.CLIENT_ID });
  assert.strictEqual(props.map.PLUXEE_REFRESH_TOKEN_ALEX, 'rt-2', 'rotated token saved');
  const get = http.calls.filter((c) => c.url.includes('/transactions'))[0];
  assert.strictEqual(get.url, PLUXEE.API_ORIGIN + PLUXEE.TRANSACTIONS_PATH + 'LUX1/transactions?limit=' + PLUXEE.PAGE_LIMIT);
  assert.strictEqual(get.headers.Authorization, 'Bearer at-1');
});

test('pluxee client: keeps the stored token when none is returned; a refused refresh asks for a new login', () => {
  const props = fakeProps({ PLUXEE_REFRESH_TOKEN_ALEX: 'rt-1' });
  BS.PluxeeClient.create(fakeHttp({ tokenResponse: { access_token: 'at' } }), tokenStore(props)).transactions('LUX1');
  assert.strictEqual(props.map.PLUXEE_REFRESH_TOKEN_ALEX, 'rt-1');

  const refused = BS.Errors.http(400, 'grant request is invalid', 'invalid_grant');
  assert.throws(() => BS.PluxeeClient.create(fakeHttp({ tokenError: refused }), tokenStore(props)).transactions('LUX1'),
    (e) => e.name === 'AuthError' && /PLUXEE_REFRESH_TOKEN_ALEX/.test(e.message) && !/rt-1/.test(e.message));
  const down = BS.Errors.http(503, 'unavailable');
  assert.throws(() => BS.PluxeeClient.create(fakeHttp({ tokenError: down }), tokenStore(props)).transactions('LUX1'),
    (e) => e.name === 'HttpError');
});

test('pluxee client: a token endpoint off the PluxeeConnect host is refused (the token is never sent there)', () => {
  const props = fakeProps({ PLUXEE_REFRESH_TOKEN_ALEX: 'rt-1' });
  const http = fakeHttp({ tokenEndpoint: 'https://elsewhere.example/token' });
  assert.throws(() => BS.PluxeeClient.create(http, tokenStore(props)).transactions('LUX1'), /no usable token endpoint/);
  assert.ok(!http.calls.some((c) => c.method === 'POST'));
});

test('pluxee provider: one API call serves the account and the transactions, filtered to the range', () => {
  let reads = 0;
  const client = { transactions: (id) => { reads++; assert.strictEqual(id, 'LUX1'); return { card: CARD, transactions: ITEMS }; } };
  const p = BS.PluxeeProvider.create({ connection: CONN, cardId: 'LUX1', client, toLocalDate });
  assert.strictEqual(p.maxWindowDays, Infinity);
  assert.strictEqual(p.listAccounts()[0].accountId, CARD.cardId);
  assert.deepStrictEqual(p.getBalances(CARD.cardId), []);
  const rows = p.getTransactions(CARD.cardId, { dateFrom: '2026-10-20', dateTo: '2026-10-31' });
  assert.deepStrictEqual(rows.map((t) => t.transactionId), ['LUXa1', 'LUXa4']);
  assert.strictEqual(reads, 1);
  assert.throws(() => p.authorize(), /PLUXEE_CARD_/);
});

test('pluxee registry: connections, readiness and status lines come from Script Properties', () => {
  const props = fakeProps({ PLUXEE_CARD_ALEX: 'LUX1', PLUXEE_REFRESH_TOKEN_ALEX: 'rt-secret' });
  const wired = { props, connections: { all: () => [] }, store: { loadSession: () => null }, nowMs: () => 0 };
  const c = BS.ProviderRegistry.connections(wired).filter((x) => x.provider === 'pluxee')[0];
  assert.strictEqual(c.key, 'pluxee-alex');
  assert.ok(BS.ProviderRegistry.isConnected(c, wired));
  const p = BS.ProviderRegistry.create(c, wired, { http: () => fakeHttp(), toLocalDate });
  assert.strictEqual(p.listAccounts()[0].accountName, 'Demo Lunch');
  const status = BS.ProviderRegistry.statusLines(wired).join('\n');
  assert.match(status, /PLUXEE_CARD_ALEX: set; PLUXEE_REFRESH_TOKEN_ALEX: set/);
  assert.ok(!status.includes('rt-secret'));
});
