const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load');

const NOW = Date.parse('2026-01-01T00:00:00Z');
const memProps = () => {
  const m = {};
  return { getProperty: (k) => (k in m ? m[k] : null), setProperty: (k, v) => { m[k] = v; } };
};

test('normalizeAccount masks the IBAN and requires uid', () => {
  const a = BS.EbNormalize.normalizeAccount({
    uid: 'u1', name: 'Main', currency: 'EUR', cash_account_type: 'CACC',
    account_id: { iban: 'GB82WEST12345698765432' }
  });
  assert.strictEqual(a.ibanMasked, '****5432');
  assert.ok(!JSON.stringify(a).includes('GB82'));
  assert.throws(() => BS.EbNormalize.normalizeAccount({}), { name: 'DataError' });
});

test('normalizeBalance converts the decimal string amount', () => {
  const b = BS.EbNormalize.normalizeBalance(
    { name: 'x', balance_type: 'CLBD', balance_amount: { amount: '12.50', currency: 'SEK' } }, 'u1', 't');
  assert.strictEqual(b.amount, 12.5);
  assert.strictEqual(b.currency, 'SEK');
  assert.throws(() => BS.EbNormalize.normalizeBalance({ balance_amount: { amount: 'x' } }, 'u1', 't'));
});

test('consentExpiry clamps to the bank maximum', () => {
  const days = (iso) => Math.round((Date.parse(iso) - NOW) / 86400000);
  assert.strictEqual(days(BS.EbAuth.consentExpiry(NOW, 90, undefined)), 90);
  assert.strictEqual(days(BS.EbAuth.consentExpiry(NOW, 90, 30 * 86400)), 30);
  assert.ok(Date.parse(BS.EbAuth.consentExpiry(NOW, 30, 30 * 86400)) < NOW + 30 * 86400000);
});

test('findAspsp matches case-insensitively and suggests alternatives', () => {
  const list = [{ name: 'Revolut' }, { name: 'Revolut Business' }];
  assert.strictEqual(BS.EbAuth.findAspsp(list, 'revolut').name, 'Revolut');
  assert.throws(() => BS.EbAuth.findAspsp(list, 'Revolutx'), /Revolut Business/);
});

test('session store: pending is single use and expires', () => {
  let now = NOW;
  const store = BS.EbSessionStore.create(memProps(), () => now);
  store.savePending('s1', { bankKey: 'revolut' });
  assert.strictEqual(store.consumePending('s1').bankKey, 'revolut');
  assert.throws(() => store.consumePending('s1'), { name: 'AuthError' });
  store.savePending('s2', {});
  now += 16 * 60000;
  assert.throws(() => store.consumePending('s2'), { name: 'AuthError' });
});

function wired(overrides) {
  const props = memProps();
  const store = BS.EbSessionStore.create(props, () => NOW);
  const calls = [];
  const client = {
    listAspsps: () => ({ aspsps: [{ name: 'Revolut', country: 'LU', maximum_consent_validity: 7776000 }] }),
    startAuth: (body) => { calls.push(body); return { url: 'https://bank.example/login' }; },
    createSession: () => ({
      session_id: 'sess-1',
      access: { valid_until: '2026-03-01T00:00:00Z' },
      accounts: [{ uid: 'a1', currency: 'EUR', name: 'Main', account_id: { iban: 'GB82WEST12345698765432' } }]
    }),
    getSession: () => ({ status: 'AUTHORIZED' }),
    getBalances: () => ({ balances: [{ name: 'b', balance_type: 'CLBD', balance_amount: { amount: '5', currency: 'EUR' } }] }),
    ...overrides
  };
  const auth = BS.EbAuth.create({
    client, store, nowMs: () => NOW, uuid: () => 'state-1',
    getRedirectUrl: () => 'https://script.example/exec'
  });
  const bank = BS.Config.BANKS.REVOLUT;
  const provider = BS.EnableBankingProvider.create({ bank, client, store, auth, nowMs: () => NOW });
  return { provider, store, calls, client };
}

test('authorize builds the /auth body and stores the pending state', () => {
  const { provider, calls, store } = wired();
  const r = provider.authorize();
  assert.deepStrictEqual(r, { url: 'https://bank.example/login', state: 'state-1' });
  assert.strictEqual(calls[0].redirect_url, 'https://script.example/exec');
  assert.strictEqual(calls[0].psu_type, 'personal');
  assert.strictEqual(store.consumePending('state-1').bankKey, 'revolut');
});

test('complete persists the session with masked accounts; list/balances then work', () => {
  const { provider, store } = wired();
  const auth = BS.EbAuth.create({
    client: wired().client, store, nowMs: () => NOW, uuid: () => 's', getRedirectUrl: () => 'x'
  });
  store.savePending('st', { bankKey: 'revolut', validUntil: '2026-02-01T00:00:00Z' });
  const summary = auth.complete('code', 'st');
  assert.strictEqual(summary.accountCount, 1);
  assert.strictEqual(summary.validUntil, '2026-03-01T00:00:00Z');
  const accounts = provider.listAccounts();
  assert.strictEqual(accounts[0].ibanMasked, '****5432');
  const balances = provider.getBalances('a1');
  assert.strictEqual(balances[0].amount, 5);
  assert.ok(!JSON.stringify(store.loadSession('revolut')).includes('GB82'));
});

test('provider rejects use before authorization, after expiry, and when revoked', () => {
  const { provider, store } = wired();
  assert.throws(() => provider.listAccounts(), { name: 'AuthError' });
  store.saveSession('revolut', { sessionId: 's', accessValidUntil: '2025-01-01T00:00:00Z', accounts: [] });
  assert.throws(() => provider.listAccounts(), /expired/);
  store.saveSession('revolut', { sessionId: 's', accessValidUntil: '2027-01-01T00:00:00Z', accounts: [] });
  const revoked = wired({ getSession: () => ({ status: 'CLOSED' }) });
  revoked.store.saveSession('revolut', { sessionId: 's', accessValidUntil: '2027-01-01T00:00:00Z', accounts: [] });
  assert.throws(() => revoked.provider.listAccounts(), /CLOSED/);
});

test('getTransactions is explicitly not implemented yet', () => {
  assert.throws(() => wired().provider.getTransactions(), { name: 'NotImplementedError' });
});

test('assertBankProvider names missing methods', () => {
  assert.throws(() => BS.BankProvider.assertBankProvider({ authorize() {} }), /listAccounts/);
});

test('token provider caches the JWT and errors clearly on missing credentials', () => {
  const props = memProps();
  const cache = (() => { const m = {}; return { get: (k) => m[k] || null, put: (k, v) => { m[k] = v; } }; })();
  let signed = 0;
  const token = BS.EbToken.create({
    props, cache, nowSeconds: () => 1000,
    jwtDeps: { base64UrlEncode: (x) => Buffer.from(x).toString('base64url'), signRs256: () => { signed++; return Buffer.alloc(1); } }
  });
  assert.throws(() => token.getToken(), /EB_APP_ID is not set/);
  props.setProperty('EB_APP_ID', 'id');
  props.setProperty('EB_PRIVATE_KEY', '-----BEGIN PRIVATE KEY-----\nQUJD\n-----END PRIVATE KEY-----');
  assert.strictEqual(token.getToken(), token.getToken());
  assert.strictEqual(signed, 1);
});

test('raw sync writes both tabs in single batched writes', () => {
  const written = {};
  const writer = { writeTable: (tab, cols, rows) => { written[tab] = rows; return rows.length; } };
  const provider = {
    listAccounts: () => [{ accountId: 'a1' }, { accountId: 'a2' }],
    getBalances: (id) => [{ accountId: id }]
  };
  assert.deepStrictEqual(BS.RawSync.run(provider, writer), { accounts: 2, balances: 2 });
  assert.strictEqual(written.raw_balances.length, 2);
});

test('toMatrix emits a header row and blanks for missing values', () => {
  const m = BS.RawWriter.toMatrix([{ header: 'a', key: 'a' }, { header: 'b', key: 'b' }], [{ a: 1 }]);
  assert.deepStrictEqual(m, [['a', 'b'], [1, '']]);
});
