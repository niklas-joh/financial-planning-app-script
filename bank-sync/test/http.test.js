const test = require('node:test');
const assert = require('node:assert');
const { Http } = require('./load');

const reply = (status, body) => ({ getResponseCode: () => status, getContentText: () => (body ? JSON.stringify(body) : '') });

function client(responses, maxRetries) {
  const sleeps = [];
  let calls = 0;
  const http = Http.create({
    fetch: () => responses[Math.min(calls++, responses.length - 1)],
    sleep: (ms) => sleeps.push(ms),
    maxRetries
  });
  return { http, sleeps, calls: () => calls };
}

test('returns parsed JSON on success', () => {
  const { http } = client([reply(200, { ok: 1 })]);
  assert.deepStrictEqual(http.request({ method: 'GET', url: 'u' }), { ok: 1 });
});

test('retries 429/5xx with exponential backoff then succeeds', () => {
  const { http, sleeps } = client([reply(429), reply(503), reply(200, { ok: 1 })]);
  assert.deepStrictEqual(http.request({ method: 'GET', url: 'u' }), { ok: 1 });
  assert.deepStrictEqual(sleeps, [500, 1000]);
});

test('does not retry 4xx and throws a redacted HttpError', () => {
  const { http, calls } = client([reply(401, { message: 'bad Bearer abc123' })]);
  assert.throws(() => http.request({ method: 'GET', url: 'u' }),
    (e) => e.name === 'HttpError' && e.status === 401 && !e.message.includes('abc123'));
  assert.strictEqual(calls(), 1);
});

test('gives up after maxRetries', () => {
  const { http, calls } = client([reply(500)], 2);
  assert.throws(() => http.request({ method: 'GET', url: 'u' }), { status: 500 });
  assert.strictEqual(calls(), 3);
});
