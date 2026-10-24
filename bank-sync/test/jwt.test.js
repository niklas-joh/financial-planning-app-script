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

test('normalizePem keeps a well-formed key byte-for-byte (apart from a trailing newline)', () => {
  assert.strictEqual(Jwt.normalizePem(privateKey), privateKey.trim() + '\n');
});

test('normalizePem repairs a key whose line breaks were pasted as spaces', () => {
  const flattened = privateKey.trim().split('\n').join(' ');
  assert.strictEqual(Jwt.normalizePem(flattened), privateKey.trim() + '\n');
});

test('normalizePem repairs literal \\n sequences and surrounding whitespace', () => {
  const escaped = '  ' + privateKey.trim().split('\n').join('\\n') + '  ';
  assert.strictEqual(Jwt.normalizePem(escaped), privateKey.trim() + '\n');
});

test('a key pasted as one line still signs a verifiable JWT', () => {
  const flattened = privateKey.trim().split('\n').join(' ');
  const [h, c, s] = Jwt.create({ ...params, privateKey: flattened }, deps).split('.');
  assert.ok(crypto.verify('sha256', Buffer.from(h + '.' + c), publicKey, Buffer.from(s, 'base64url')));
});

test('normalizePem rejects text without a PEM block, without echoing it', () => {
  assert.throws(() => Jwt.normalizePem('secret-ish-text'), (e) => /not a PEM key/.test(e.message) && !/secret-ish/.test(e.message));
});

test('normalizePem rejects a PKCS#1 key with a conversion hint', () => {
  const pkcs1 = '-----BEGIN RSA PRIVATE KEY-----\nQUJD\n-----END RSA PRIVATE KEY-----';
  assert.throws(() => Jwt.normalizePem(pkcs1), /PKCS#8[\s\S]*openssl pkcs8/);
});
