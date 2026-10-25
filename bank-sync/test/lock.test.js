const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');
require('../src/entry/web-app.js');

const fakeLockService = (free) => {
  const log = [];
  return { log, getScriptLock: () => ({
    tryLock: (ms) => { log.push('try ' + ms); return free; },
    releaseLock: () => log.push('release')
  }) };
};

test('lock: runs the action and releases, also when it throws', () => {
  const svc = fakeLockService(true);
  assert.strictEqual(BS.Lock.run(svc, 5, () => 'done'), 'done');
  assert.throws(() => BS.Lock.run(svc, 5, () => { throw new Error('boom'); }), /boom/);
  assert.deepStrictEqual(svc.log, ['try 5', 'release', 'try 5', 'release']);
});

test('lock: a held lock raises LockError and the action never runs', () => {
  const svc = fakeLockService(false);
  let ran = false;
  assert.throws(() => BS.Lock.run(svc, 5, () => { ran = true; }), { name: 'LockError' });
  assert.strictEqual(ran, false);
  assert.deepStrictEqual(svc.log, ['try 5']);
});

test('redirect: reports busy and completes nothing while the script lock is held', () => {
  let completed = false;
  BS.Container = { services: () => ({ auth: { complete: () => { completed = true; return {}; } } }) };
  globalThis.LockService = fakeLockService(false);
  const text = BS.WebApp.handleRedirect({ parameter: { code: 'c', state: 's' } });
  assert.match(text, /try again/);
  assert.strictEqual(completed, false);
});

test('redirect: completes under the lock and releases it', () => {
  BS.Container = { services: () => ({ auth: { complete: () => ({ connectionKey: 'k', accountCount: 2, validUntil: 'x' }) } }) };
  const svc = globalThis.LockService = fakeLockService(true);
  assert.match(BS.WebApp.handleRedirect({ parameter: { code: 'c', state: 's' } }), /^Authorized k \(2 accounts\)/);
  assert.deepStrictEqual(svc.log.slice(-1), ['release']);
});
