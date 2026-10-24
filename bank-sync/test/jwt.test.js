const test = require('node:test');
const assert = require('node:assert');
const crypto = require('node:crypto');
const { Jwt } = require('./load');

// Throwaway key generated per run; not related to any real credential.
const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', {
  modulusLength: 2048,
  privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
  publicKeyEncoding: { type: 'spki', format: 'pem' }
});

const b64url = (input) => Buffer.from(input).toString('base64url');
const deps = {
  base64UrlEncode: b64url,
  signRs256: (data, key) => crypto.sign('sha256', Buffer.from(data), key)
};
const params = {
  appId: 'fake-app-id', privateKey, nowSeconds: 1000, ttlSeconds: 3600,
  issuer: 'enablebanking.com', audience: 'api.enablebanking.com'
};

test('creates a verifiable RS256 JWT with the expected header and claims', () => {
  const [h, c, s] = Jwt.create(params, deps).split('.');
  assert.deepStrictEqual(JSON.parse(Buffer.from(h, 'base64url')), { typ: 'JWT', alg: 'RS256', kid: 'fake-app-id' });
  assert.deepStrictEqual(JSON.parse(Buffer.from(c, 'base64url')), {
    iss: 'enablebanking.com', aud: 'api.enablebanking.com', iat: 1000, exp: 4600
  });
  assert.ok(crypto.verify('sha256', Buffer.from(h + '.' + c), publicKey, Buffer.from(s, 'base64url')));
});

test('normalizePem converts literal \\n sequences and trims', () => {
  assert.strictEqual(Jwt.normalizePem('  A\\nB  '), 'A\nB');
});

test('signer receives the normalized PEM', () => {
  let seen;
  Jwt.create({ ...params, privateKey: 'X\\nY' }, { ...deps, signRs256: (d, k) => { seen = k; return Buffer.alloc(1); } });
  assert.strictEqual(seen, 'X\nY');
});
