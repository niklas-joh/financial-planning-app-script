const test = require('node:test');
const assert = require('node:assert');
const { Redact } = require('./load');

test('redacts PEM blocks, JWTs, bearer tokens and IBANs', () => {
  const input = '-----BEGIN PRIVATE KEY-----\nabc\n-----END PRIVATE KEY----- ' +
    'Bearer abc.def-ghi eyJhbGci.eyJpc3Mi.c2ln GB82WEST12345698765432';
  const out = Redact.text(input);
  assert.ok(!/abc|eyJ|GB82/.test(out), out);
  assert.match(out, /\[REDACTED:pem\]/);
  assert.match(out, /\[REDACTED:iban\]/);
});

test('maskIban keeps only the last four characters', () => {
  assert.strictEqual(Redact.maskIban('GB82 WEST 1234 5698 7654 32'), '****5432');
  assert.strictEqual(Redact.maskIban(undefined), '');
});

test('redacts HTTP Basic credentials but leaves the word "basic" in ordinary text alone', () => {
  const out = Redact.text('Authorization: Basic ZmFrZS1rZXk6ZmFrZS1zZWNyZXQ= failed; basic information is fine');
  assert.ok(!out.includes('ZmFrZS1rZXk6ZmFrZS1zZWNyZXQ='));
  assert.match(out, /Basic \[REDACTED\] failed; basic information is fine/);
});
