const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');

const DAY = 86400000;
const NOW = Date.parse('2026-12-01T08:00:00Z');
const A = BS.Config.AUTOMATION;
const WARN = BS.Config.CONSENT.WARN_DAYS;
const CONN = { key: 'demo-bank-alex', holder: 'alex', bank: { aspspName: 'Demo Bank', country: 'LU' } };
const inDays = (d) => new Date(NOW + d * DAY).toISOString();

const fakeProps = (init) => {
  const map = Object.assign({}, init);
  return { map, getProperty: (k) => (k in map ? map[k] : null), setProperty: (k, v) => { map[k] = v; } };
};

const fakeTriggers = (initial) => {
  let next = 100;
  const list = (initial || []).slice();
  const log = [];
  return {
    log, all: list,
    list: () => list.slice(),
    createDaily: (handler, hour) => { log.push('daily ' + handler + ' ' + hour); list.push({ id: String(next++), handler }); },
    createAfter: (handler, ms) => { log.push('after ' + handler + ' ' + ms); list.push({ id: String(next++), handler }); },
    remove: (id) => { log.push('remove ' + id); list.splice(list.findIndex((t) => t.id === id), 1); }
  };
};

const okOutcome = (complete) => ({
  results: [{ connection: 'demo-bank-alex', status: 'ok', transactions: { added: 2, updated: 1, removed: 0, complete } }],
  raw: null,
  fx: { changed: 3, missing: 0, waiting: 0, balancesMissing: 0, failed: [] }
});

// ── Schedule ──────────────────────────────────────────────────────────────

test('schedule: plan creates one trigger, keeps an existing one and removes duplicates', () => {
  const other = { id: '9', handler: 'somethingElse' };
  assert.deepStrictEqual(BS.Schedule.plan([other], 'h', true), { create: true, remove: [] });
  assert.deepStrictEqual(BS.Schedule.plan([other, { id: '1', handler: 'h' }], 'h', true), { create: false, remove: [] });
  assert.deepStrictEqual(BS.Schedule.plan([{ id: '1', handler: 'h' }, other, { id: '2', handler: 'h' }], 'h', true),
    { create: false, remove: ['2'] });
  assert.deepStrictEqual(BS.Schedule.plan([{ id: '1', handler: 'h' }, other, { id: '2', handler: 'h' }], 'h', false),
    { create: false, remove: ['1', '2'] });
});

test('schedule: a run is incomplete when the budget skipped a connection or cut its transactions short', () => {
  assert.strictEqual(BS.Schedule.isIncomplete(okOutcome(true)), false);
  assert.strictEqual(BS.Schedule.isIncomplete(okOutcome(false)), true);
  assert.strictEqual(BS.Schedule.isIncomplete({ results: [{ status: 'skipped: time budget' }] }), true);
  assert.strictEqual(BS.Schedule.isIncomplete({ results: [{ status: 'error', error: 'x' }, { status: 'not connected' }] }), false);
});

test('schedule: follow-ups chain only while work is left and up to the maximum', () => {
  assert.strictEqual(BS.Schedule.shouldContinue(true, 0, 3), true);
  assert.strictEqual(BS.Schedule.shouldContinue(true, 2, 3), true);
  assert.strictEqual(BS.Schedule.shouldContinue(true, 3, 3), false);
  assert.strictEqual(BS.Schedule.shouldContinue(false, 0, 3), false);
});

// ── Consent warnings ──────────────────────────────────────────────────────

test('consent: no warning while the expiry is further away than the largest threshold', () => {
  assert.strictEqual(BS.ConsentWarnings.due(inDays(31), null, NOW, WARN), null);
  assert.strictEqual(BS.ConsentWarnings.daysLeft(inDays(31), NOW), 31);
});

test('consent: each threshold is warned once, in turn', () => {
  const at30 = BS.ConsentWarnings.due(inDays(30), null, NOW, WARN);
  assert.strictEqual(at30.threshold, 30);
  assert.deepStrictEqual(at30.record, { validUntil: inDays(30), sent: [30] });
  // Same day again: already sent.
  assert.strictEqual(BS.ConsentWarnings.due(inDays(30), at30.record, NOW, WARN), null);
  // 14 days before, with the 30-day warning on record.
  const until = inDays(30);
  const at14 = BS.ConsentWarnings.due(until, at30.record, NOW + 16 * DAY, WARN);
  assert.strictEqual(at14.threshold, 14);
  assert.deepStrictEqual(at14.record.sent, [30, 14]);
});

test('consent: reaching several thresholds at once sends one warning and records them all', () => {
  const w = BS.ConsentWarnings.due(inDays(5), null, NOW, WARN);
  assert.strictEqual(w.threshold, 7);
  assert.strictEqual(w.daysLeft, 5);
  assert.deepStrictEqual(w.record.sent, [30, 14, 7]);
  assert.strictEqual(BS.ConsentWarnings.due(inDays(5), w.record, NOW, WARN), null);
});

test('consent: a lapsed consent is warned once more; a renewed consent starts afresh', () => {
  const until = inDays(-2);
  const sentBefore = { validUntil: until, sent: [30, 14, 7, 1] };
  const lapsed = BS.ConsentWarnings.due(until, sentBefore, NOW, WARN);
  assert.strictEqual(lapsed.threshold, 0);
  assert.deepStrictEqual(lapsed.record.sent, [30, 14, 7, 1, 0]);
  assert.strictEqual(BS.ConsentWarnings.due(until, lapsed.record, NOW, WARN), null);
  // Renewed: a new expiry far away needs no warning, and a later 30-day mark warns again.
  assert.strictEqual(BS.ConsentWarnings.due(inDays(180), lapsed.record, NOW, WARN), null);
  assert.strictEqual(BS.ConsentWarnings.due(inDays(29), lapsed.record, NOW, WARN).threshold, 30);
});

test('consent: an invalid date warns nothing', () => {
  assert.strictEqual(BS.ConsentWarnings.due('not a date', null, NOW, WARN), null);
});

test('consent: the email names the connection, the date and the renewal steps, nothing else', () => {
  const until = '2027-01-22T10:00:00.000Z';
  const w = BS.ConsentWarnings.due(until, null, Date.parse(until) - 14 * DAY, WARN);
  const mail = BS.ConsentWarnings.message(CONN, until, w, WARN);
  assert.strictEqual(mail.subject, 'Bank Sync: consent for Demo Bank (alex) lapses on 2027-01-22');
  assert.match(mail.body, /lapses on 2027-01-22 \(in 14 days\)/);
  assert.match(mail.body, /Connect a bank\.\.\.: enter the country \(LU\), the bank \(Demo Bank\) and the same holder name \(alex\)/);
  assert.match(mail.body, /30, 14, 7, 1 days before/);
  const lapsed = BS.ConsentWarnings.message(CONN, until, { threshold: 0, daysLeft: -1 }, WARN);
  assert.strictEqual(lapsed.subject, 'Bank Sync: consent for Demo Bank (alex) has lapsed');
  assert.match(lapsed.body, /lapsed on 2027-01-22/);
});

test('consent: status note flags lapsing-soon and lapsed consents', () => {
  assert.strictEqual(BS.ConsentWarnings.statusNote(inDays(60), NOW, WARN), '');
  assert.match(BS.ConsentWarnings.statusNote(inDays(20), NOW, WARN), /renew soon/);
  assert.match(BS.ConsentWarnings.statusNote(inDays(-1), NOW, WARN), /LAPSED/);
});

test('consent: sendDue emails each due warning once and records it only after sending', () => {
  const props = fakeProps();
  const sent = [];
  const deps = { consents: [{ connection: CONN, validUntil: inDays(10) }, { connection: Object.assign({}, CONN, { key: 'far' }), validUntil: inDays(90) }],
    props, send: (s) => sent.push(s), nowMs: () => NOW };
  assert.strictEqual(BS.ConsentWarnings.sendDue(deps), 1);
  assert.strictEqual(BS.ConsentWarnings.sendDue(deps), 0);
  assert.strictEqual(sent.length, 1);
  assert.deepStrictEqual(JSON.parse(props.map['BS_CONSENT_WARNED_demo-bank-alex']).sent, [30, 14]);

  const failing = fakeProps();
  assert.throws(() => BS.ConsentWarnings.sendDue(Object.assign({}, deps, { props: failing, send: () => { throw new Error('quota'); } })), /quota/);
  assert.deepStrictEqual(failing.map, {});
});

// ── Last run ──────────────────────────────────────────────────────────────

test('last run: records counts per connection and marks failures', () => {
  const meta = { source: 'daily', at: '2026-12-01 06:10', chain: 0, warningsSent: 1, followUp: false };
  const ok = BS.LastRun.fromOutcome(okOutcome(true), meta);
  assert.strictEqual(ok.ok, true);
  assert.deepStrictEqual(ok.connections[0], { connection: 'demo-bank-alex', status: 'ok', added: 2, updated: 1, removed: 0, complete: true, error: undefined });
  const text = BS.LastRun.describe(ok).join('\n');
  assert.match(text, /^Last sync: 2026-12-01 06:10 \(daily\): OK/);
  assert.match(text, /demo-bank-alex: 2 added, 1 updated, 0 pending removed/);
  assert.match(text, /Consent warnings emailed: 1/);

  const bad = BS.LastRun.fromOutcome({ results: [{ connection: 'c', status: 'error', error: 'AuthError: Consent has expired' }], fx: null }, meta);
  assert.strictEqual(bad.ok, false);
  assert.match(BS.LastRun.describe(bad).join('\n'), /FAILED[\s\S]*c: error \(AuthError: Consent has expired\)/);
  assert.strictEqual(BS.LastRun.fromOutcome(okOutcome(true), Object.assign({}, meta, { warningsError: 'x' })).ok, false);
  assert.deepStrictEqual(BS.LastRun.describe(null), ['Last sync: none recorded yet']);
});

// ── AutoSync ──────────────────────────────────────────────────────────────

test('auto sync: enable creates one daily trigger and never a duplicate; disable removes it and follow-ups', () => {
  const t = fakeTriggers([{ id: '1', handler: 'otherFunction' }]);
  assert.strictEqual(BS.AutoSync.enable(t), true);
  assert.strictEqual(BS.AutoSync.enable(t), false);
  assert.deepStrictEqual(t.log, ['daily ' + A.DAILY_HANDLER + ' ' + A.DAILY_HOUR]);
  t.all.push({ id: '7', handler: A.DAILY_HANDLER }, { id: '8', handler: A.CONTINUE_HANDLER });
  assert.strictEqual(BS.AutoSync.enable(t), false);
  assert.deepStrictEqual(t.log.slice(-1), ['remove 7']);
  assert.strictEqual(BS.AutoSync.disable(t), 2);
  assert.deepStrictEqual(t.all, [{ id: '1', handler: 'otherFunction' }]);
  assert.strictEqual(BS.AutoSync.disable(t), 0);
});

const autoDeps = (overrides) => {
  const props = fakeProps();
  const store = { saved: null, load: () => store.saved, save: (r) => { store.saved = r; } };
  return Object.assign({
    triggers: fakeTriggers(), lastRun: store, sync: () => okOutcome(true),
    consents: () => [{ connection: CONN, validUntil: inDays(100) }], send: () => {}, props,
    nowMs: () => NOW, nowLabel: () => '2026-12-01 06:10'
  }, overrides);
};

test('auto sync: a finished run is recorded and schedules no follow-up', () => {
  const deps = autoDeps();
  const record = BS.AutoSync.runTriggered(deps, 'daily');
  assert.strictEqual(record.ok, true);
  assert.strictEqual(record.followUp, false);
  assert.strictEqual(deps.lastRun.saved, record);
  assert.deepStrictEqual(deps.triggers.log, []);
});

test('auto sync: an unfinished run schedules one follow-up and old follow-ups are cleared first', () => {
  const deps = autoDeps({ sync: () => okOutcome(false), triggers: fakeTriggers([{ id: '5', handler: A.CONTINUE_HANDLER }]) });
  const record = BS.AutoSync.runTriggered(deps, 'daily');
  assert.strictEqual(record.followUp, true);
  assert.strictEqual(record.chain, 0);
  assert.deepStrictEqual(deps.triggers.log, ['remove 5', 'after ' + A.CONTINUE_HANDLER + ' ' + A.CONTINUE_AFTER_MS]);
  assert.strictEqual(deps.triggers.all.filter((t) => t.handler === A.CONTINUE_HANDLER).length, 1);
});

test('auto sync: follow-ups stop after the maximum chain even if work is left', () => {
  const deps = autoDeps({ sync: () => okOutcome(false) });
  BS.AutoSync.runTriggered(deps, 'daily');
  for (let i = 1; i <= A.MAX_CONTINUATIONS; i++) {
    const r = BS.AutoSync.runTriggered(deps, 'follow-up');
    assert.strictEqual(r.chain, i);
    assert.strictEqual(r.followUp, i < A.MAX_CONTINUATIONS);
  }
  assert.strictEqual(deps.triggers.all.length, 0);
  // The next daily run starts a fresh chain.
  assert.strictEqual(BS.AutoSync.runTriggered(deps, 'daily').followUp, true);
});

test('auto sync: a busy lock is recorded and retried by a follow-up, without throwing', () => {
  const deps = autoDeps({ sync: () => { throw BS.Errors.create('LockError', 'Another sync or login is running'); } });
  const record = BS.AutoSync.runTriggered(deps, 'daily');
  assert.strictEqual(record.ok, false);
  assert.strictEqual(record.followUp, true);
  assert.match(BS.LastRun.describe(record).join('\n'), /FAILED, unfinished; a follow-up run is scheduled\n {2}LockError/);
});

test('auto sync: any other failure is recorded, then rethrown so Apps Script reports it too', () => {
  const deps = autoDeps({ sync: () => { throw new Error('sheet gone'); } });
  assert.throws(() => BS.AutoSync.runTriggered(deps, 'daily'), /sheet gone/);
  assert.strictEqual(deps.lastRun.saved.ok, false);
  assert.match(deps.lastRun.saved.error, /sheet gone/);
  assert.deepStrictEqual(deps.triggers.log, []);
});

test('auto sync: due consent warnings are emailed; a mail failure fails the record but not the sync', () => {
  const mails = [];
  const deps = autoDeps({ consents: () => [{ connection: CONN, validUntil: inDays(6) }], send: (s) => mails.push(s) });
  assert.strictEqual(BS.AutoSync.runTriggered(deps, 'daily').warningsSent, 1);
  assert.strictEqual(BS.AutoSync.runTriggered(deps, 'daily').warningsSent, 0);
  assert.strictEqual(mails.length, 1);

  const failing = autoDeps({ consents: () => [{ connection: CONN, validUntil: inDays(6) }], send: () => { throw new Error('quota'); } });
  const record = BS.AutoSync.runTriggered(failing, 'daily');
  assert.strictEqual(record.ok, false);
  assert.match(record.warningsError, /quota/);
  assert.strictEqual(record.connections[0].added, 2);
});

test('auto sync: status lines show whether the daily sync is on and the last run', () => {
  const off = BS.AutoSync.statusLines([], null);
  assert.match(off[0], /^Daily sync: off/);
  const on = BS.AutoSync.statusLines([{ id: '1', handler: A.DAILY_HANDLER }], null);
  assert.match(on[0], /^Daily sync: on \(between 6:00 and 7:00, Europe\/Luxembourg\)/);
  assert.strictEqual(on[1], 'Last sync: none recorded yet');
});

test('auto sync: a manual run is recorded with chain 0', () => {
  const store = { save: (r) => { store.saved = r; } };
  BS.AutoSync.recordManual(store, okOutcome(true), '2026-12-01 12:00');
  assert.strictEqual(store.saved.source, 'menu');
  assert.strictEqual(store.saved.chain, 0);
});
