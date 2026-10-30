// Trading 212 report export: the client calls, the request and the checks that save a finished report
// to the exports folder, the triggers they schedule, and the CSV import reading that folder. Fake data only.
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');

const A = BS.Config.AUTOMATION;
const T212 = BS.Config.TRADING212;
const NOW = Date.parse('2026-11-01T04:10:00Z');
const base64 = (s) => Buffer.from(s, 'utf8').toString('base64');

const fakeProps = (init) => {
  const map = Object.assign({}, init);
  return { map, getProperty: (k) => (k in map ? map[k] : null), setProperty: (k, v) => { map[k] = v; }, getKeys: () => Object.keys(map) };
};

// Connections and credentials are built by the real modules from fake property names, as in production.
const PROPS_INIT = { T212_API_KEY_ALEX: 'k1', T212_API_SECRET_ALEX: 's1' };
const FOLDER_PROP = BS.Config.PROPERTY_KEYS.T212_EXPORT_FOLDER_PREFIX + 'ALEX';
const STATE_PROP = BS.Config.PROPERTY_KEYS.T212_EXPORT_PREFIX + 'ALEX';
const CONN = BS.T212Credentials.connections(fakeProps(PROPS_INIT))[0];

const REPORT = 'Action,Time (UTC),Notes,ID,Total,Currency (Total),Merchant name,Merchant category\n' +
  'Card debit,2026-10-21 09:27:53+00:00,,id-card-1,-5.00,"EUR","DEMO CAFE","RESTAURANTS"\n';

const fakeTriggers = () => {
  const list = [];
  const log = [];
  let next = 1;
  return {
    log, list: () => list.slice(),
    createAfter: (handler, ms) => { log.push('after ' + handler + ' ' + ms); list.push({ id: String(next++), handler }); },
    remove: (ids) => ids.forEach((id) => { log.push('remove ' + id); list.splice(list.findIndex((t) => t.id === id), 1); })
  };
};

/** A client whose reports list answers with the given statuses in turn; records what was asked. */
function fakeClient(statuses, over) {
  const calls = { requests: [], lists: 0, downloads: [] };
  let nextId = 41;
  return Object.assign({
    calls,
    requestReport: (body) => { calls.requests.push(body); return { reportId: ++nextId }; },
    reports: () => {
      const status = statuses[Math.min(calls.lists++, statuses.length - 1)];
      return [{ reportId: 1, status: 'Finished', downloadLink: 'https://storage.example/old.csv' },
        { reportId: nextId, status, downloadLink: status === 'Finished' ? 'https://storage.example/r.csv?sig=abc' : null }];
    },
    download: (link) => { calls.downloads.push(link); return REPORT; }
  }, over || {});
}

function deps(client, props) {
  const files = [];
  const triggers = fakeTriggers();
  return {
    files, triggers,
    connections: [CONN],
    clientFor: () => client,
    props: props || fakeProps(Object.assign({ [FOLDER_PROP]: 'folder-1' }, PROPS_INIT)),
    folder: (id) => ({ create: (name, text) => files.push({ id, name, text }) }),
    withLock: (action) => action(),
    nowMs: () => NOW,
    todayIso: () => '2026-11-01'
  };
}

test('t212 client: a report request posts what to include and the period; the list is read with the key', () => {
  const requests = [];
  const http = { request: (req) => { requests.push(req); return req.method === 'POST' ? { reportId: 7 } : []; } };
  const client = BS.T212Client.create(http, BS.T212Credentials.credentials(fakeProps(PROPS_INIT), CONN), base64);
  assert.deepStrictEqual(client.requestReport({ dataIncluded: {}, timeFrom: 'a', timeTo: 'b' }), { reportId: 7 });
  client.reports();
  assert.deepStrictEqual(requests.map((r) => [r.method, r.url]), [
    ['POST', 'https://live.trading212.com/api/v0/equity/history/exports'],
    ['GET', 'https://live.trading212.com/api/v0/equity/history/exports']
  ]);
  assert.deepStrictEqual(requests[0].body, { dataIncluded: {}, timeFrom: 'a', timeTo: 'b' });
  assert.ok(requests.every((r) => r.headers.Authorization === 'Basic ' + base64('k1:s1')));
});

test('t212 client: a report is downloaded as text, with the key only on Trading 212 hosts', () => {
  const requests = [];
  const http = { request: (req) => { requests.push(req); return 'csv text'; } };
  const client = BS.T212Client.create(http, BS.T212Credentials.credentials(fakeProps(PROPS_INIT), CONN), base64);
  assert.strictEqual(client.download('https://storage.example/r.csv?sig=1'), 'csv text');
  client.download('https://live.trading212.com/reports/r.csv');
  assert.deepStrictEqual(requests.map((r) => [r.text, 'Authorization' in r.headers]), [[true, false], [true, true]]);
  assert.strictEqual(BS.T212Client.isOwnHost('https://trading212.com:443/x'), true);
  assert.strictEqual(BS.T212Client.isOwnHost('https://eviltrading212.com/x'), false);
  assert.strictEqual(BS.T212Client.isOwnHost('https://live.trading212.com.evil.example/x'), false);
  ['http://live.trading212.com/x', 'https://live.trading212.com@evil.example/x', '/x', ''].forEach((link) =>
    assert.throws(() => client.download(link), /Unexpected Trading 212 report link/));
});

test('http: a text request returns the body as written; errors are still parsed', () => {
  const http = BS.Http.create({
    fetch: (url) => ({ getResponseCode: () => (url === 'ok' ? 200 : 404), getContentText: () => (url === 'ok' ? 'a;b\n1;2' : '{"message":"gone"}') }),
    sleep: () => {}
  });
  assert.strictEqual(http.request({ method: 'GET', url: 'ok', text: true }), 'a;b\n1;2');
  assert.throws(() => http.request({ method: 'GET', url: 'missing', text: true }), /gone/);
});

test('t212 export: the first report covers the full history, later ones reach back a week before the last one ended', () => {
  const R = BS.T212Export.rangeFor;
  assert.deepStrictEqual(R({ coveredTo: '' }, '2026-11-01'), { from: T212.EXPORT_FIRST_FROM, to: '2026-11-01' });
  assert.deepStrictEqual(R({ coveredTo: '2026-10-01' }, '2026-11-01'), { from: '2026-09-24', to: '2026-11-01' });
  assert.deepStrictEqual(R({ coveredTo: '2026-11-05' }, '2026-11-01'), { from: '2026-10-29', to: '2026-11-01' });
});

test('t212 export: a request asks for the period and the data the report format reads, and schedules one check', () => {
  const client = fakeClient(['Queued']);
  const d = deps(client);
  const results = BS.T212Export.requestAndSchedule(d);
  assert.deepStrictEqual(results, [{ connection: 'trading212-alex', status: 'requested', from: T212.EXPORT_FIRST_FROM, to: '2026-11-01' }]);
  assert.deepStrictEqual(client.calls.requests, [{
    dataIncluded: { includeDividends: true, includeInterest: true, includeOrders: false, includeTransactions: true },
    timeFrom: T212.EXPORT_FIRST_FROM + 'T00:00:00Z', timeTo: '2026-11-01T04:10:00.000Z'
  }]);
  assert.deepStrictEqual(JSON.parse(d.props.map[STATE_PROP]).pending,
    { reportId: 42, from: T212.EXPORT_FIRST_FROM, to: '2026-11-01', requestedAt: '2026-11-01T04:10:00.000Z', checks: 0 });
  assert.deepStrictEqual(d.triggers.log, ['after ' + A.REPORT_FETCH_HANDLER + ' ' + A.REPORT_FETCH_AFTER_MS]);
  // Asking again replaces the pending report and leaves exactly one check scheduled.
  BS.T212Export.requestAndSchedule(d);
  assert.strictEqual(JSON.parse(d.props.map[STATE_PROP]).pending.reportId, 43);
  assert.deepStrictEqual(d.triggers.list().map((t) => t.handler), [A.REPORT_FETCH_HANDLER]);
});

test('t212 export: without an exports folder nothing is requested; a failed request is recorded, redacted', () => {
  const client = fakeClient(['Queued']);
  const noFolder = deps(client, fakeProps(PROPS_INIT));
  const results = BS.T212Export.requestAndSchedule(noFolder);
  assert.match(results[0].status, new RegExp('^no folder: set ' + FOLDER_PROP));
  assert.strictEqual(client.calls.requests.length, 0);
  assert.deepStrictEqual(noFolder.triggers.log, []);

  const failing = deps(fakeClient([], { requestReport: () => { throw BS.Errors.http(403, 'Missing Permissions', ''); } }));
  const out = BS.T212Export.requestAndSchedule(failing);
  assert.match(out[0].status, /^error: HttpError/);
  assert.match(JSON.parse(failing.props.map[STATE_PROP]).error, /Missing Permissions/);
  assert.deepStrictEqual(failing.triggers.log, []);
});

test('t212 export: a check waits while the report is built, then saves it to the folder and stops checking', () => {
  const client = fakeClient(['Queued', 'Running', 'Finished']);
  const d = deps(client);
  BS.T212Export.requestAndSchedule(d);
  assert.deepStrictEqual(BS.T212Export.fetchAndSchedule(d), { results: [{ connection: 'trading212-alex', status: 'waiting' }], again: true });
  assert.strictEqual(BS.T212Export.fetchAndSchedule(d).again, true);
  assert.deepStrictEqual(d.files, []);
  const done = BS.T212Export.fetchAndSchedule(d);
  assert.deepStrictEqual(done, { results: [{ connection: 'trading212-alex', status: 'saved' }], again: false });
  const name = 'trading212-report-alex-' + T212.EXPORT_FIRST_FROM + '-to-2026-11-01.csv';
  assert.deepStrictEqual(d.files, [{ id: 'folder-1', name, text: REPORT }]);
  assert.deepStrictEqual(client.calls.downloads, ['https://storage.example/r.csv?sig=abc']);
  const state = JSON.parse(d.props.map[STATE_PROP]);
  assert.deepStrictEqual([state.pending, state.coveredTo, state.saved.file, state.error], [null, '2026-11-01', name, '']);
  assert.ok(!d.props.map[STATE_PROP].includes('storage.example'), 'the download link is never stored');
  // Every check cleared the one before; none is left once the report is saved.
  assert.deepStrictEqual(d.triggers.list(), []);
  // Nothing pending: a stray check does nothing; the next request starts a week before the saved end.
  assert.deepStrictEqual(BS.T212Export.fetchPending(d), { results: [], again: false });
  BS.T212Export.request(d);
  assert.strictEqual(client.calls.requests[1].timeFrom, '2026-10-25T00:00:00Z');
});

test('t212 export: a failed or cancelled report ends the wait without moving the covered period', () => {
  ['Failed', 'Canceled'].forEach((status) => {
    const d = deps(fakeClient([status]));
    BS.T212Export.requestAndSchedule(d);
    const out = BS.T212Export.fetchAndSchedule(d);
    assert.deepStrictEqual(out, { results: [{ connection: 'trading212-alex', status: 'failed: Trading 212 report ' + status.toLowerCase() }], again: false });
    const state = JSON.parse(d.props.map[STATE_PROP]);
    assert.deepStrictEqual([state.pending, state.coveredTo], [null, '']);
  });
});

test('t212 export: checks stop after the limit (also when they fail); the report stays pending until the next request', () => {
  let lists = 0;
  const d = deps(fakeClient(['Queued'], { reports: () => { lists++; if (lists % 2) throw BS.Errors.http(500, 'down', ''); return []; } }));
  BS.T212Export.requestAndSchedule(d);
  let out;
  for (let i = 0; i < A.REPORT_MAX_CHECKS; i++) out = BS.T212Export.fetchAndSchedule(d);
  assert.strictEqual(out.again, false);
  assert.match(out.results[0].status, new RegExp('^gave up after ' + A.REPORT_MAX_CHECKS + ' checks'));
  assert.deepStrictEqual(d.triggers.list(), []);
  assert.deepStrictEqual(BS.T212Export.fetchPending(d), { results: [], again: false });
  assert.strictEqual(lists, A.REPORT_MAX_CHECKS);
  assert.match(BS.T212Export.statusLine(d.props, CONN), /no more checks until the next request/);
});

test('t212 export: status shows the folder, the saved period, a pending report and the last error, never values', () => {
  const props = fakeProps(PROPS_INIT);
  assert.strictEqual(BS.T212Export.statusLine(props, CONN), 'Trading 212 (alex) report export - ' + FOLDER_PROP + ': MISSING; none saved yet');
  const d = deps(fakeClient(['Finished']));
  BS.T212Export.requestAndSchedule(d);
  assert.match(BS.T212Export.statusLine(d.props, CONN), /: set; none saved yet; requested 2026-11-01 04:10 UTC, 0 checks$/);
  BS.T212Export.fetchAndSchedule(d);
  const line = BS.T212Export.statusLine(d.props, CONN);
  assert.strictEqual(line, 'Trading 212 (alex) report export - ' + FOLDER_PROP + ': set; last saved 2026-11-01 (' + T212.EXPORT_FIRST_FROM + ' to 2026-11-01)');
  ['k1', 's1', 'folder-1', 'storage.example'].forEach((secret) => assert.ok(!line.includes(secret)));
});

test('t212 export: the menu summary names the period, or why nothing was requested', () => {
  assert.match(BS.T212Export.summarize([{ connection: 'trading212-alex', status: 'requested', from: '2026-10-01', to: '2026-11-01' }]),
    /^trading212-alex: report requested for 2026-10-01 to 2026-11-01\. It is saved to the exports folder in about 5 minutes/);
  assert.strictEqual(BS.T212Export.summarize([]), 'No Trading 212 key with its secret is set (see the README).');
});

test('t212 export: the exports folder is also a CSV folder, so the saved report is applied like one uploaded by hand', () => {
  const props = fakeProps(Object.assign({ [FOLDER_PROP]: 'folder-1' }, PROPS_INIT));
  const csv = BS.CsvFolders.connections(props);
  assert.deepStrictEqual(csv.map((c) => c.key), ['csv-t212-alex']);
  assert.strictEqual(BS.CsvFolders.folderId(props, csv[0]), 'folder-1');
  const name = BS.T212Export.fileName(CONN, { from: '2026-10-01', to: '2026-11-01' });
  const provider = BS.CsvImportProvider.create({
    connection: csv[0],
    folder: { files: () => [{ name, size: REPORT.length, text: () => REPORT }] },
    record: { save: () => {} },
    nowMs: () => NOW
  });
  assert.deepStrictEqual(provider.getEnrichments().map((e) => [e.provider, e.transactionId, e.counterparty, e.classification]),
    [['trading212', 'id-card-1', 'DEMO CAFE', 'card']]);
});

// --- Review fixes (#84 stranded check, #85 lock, #86 redirects) ---

test('t212 export: a failed repeat request keeps a check scheduled for the report still pending', () => {
  let calls = 0;
  const d = deps(fakeClient(['Queued'], {
    requestReport: () => { calls++; if (calls > 1) throw BS.Errors.http(429, 'Limited: 1 / 30s', ''); return { reportId: 42 }; }
  }));
  BS.T212Export.requestAndSchedule(d);
  const again = BS.T212Export.requestAndSchedule(d);
  assert.match(again[0].status, /^error: HttpError/);
  assert.strictEqual(JSON.parse(d.props.map[STATE_PROP]).pending.reportId, 42);
  assert.deepStrictEqual(d.triggers.list().map((t) => t.handler), [A.REPORT_FETCH_HANDLER]);
});

test('t212 export: requests and checks run under the script lock; a busy lock requests nothing and only reschedules the check', () => {
  const order = [];
  const client = fakeClient(['Queued']);
  const d = deps(client);
  d.withLock = (action) => { order.push('lock'); const out = action(); order.push('unlock'); return out; };
  BS.T212Export.requestAndSchedule(d);
  assert.deepStrictEqual(order, ['lock', 'unlock']);
  assert.strictEqual(client.calls.requests.length, 1);

  const busy = () => { throw BS.Errors.create('LockError', 'Another sync or login is running; try again in a minute'); };
  d.withLock = busy;
  assert.throws(() => BS.T212Export.requestAndSchedule(d), /Another sync/);
  assert.strictEqual(client.calls.requests.length, 1);
  assert.deepStrictEqual(BS.T212Export.fetchAndSchedule(d), { results: [], again: true, busy: true });
  assert.strictEqual(client.calls.lists, 0);
  assert.strictEqual(JSON.parse(d.props.map[STATE_PROP]).pending.checks, 0);
  assert.deepStrictEqual(d.triggers.list().map((t) => t.handler), [A.REPORT_FETCH_HANDLER]);
});

test('http: with followRedirects false a redirect is returned with its target, not followed', () => {
  const seen = [];
  const http = BS.Http.create({
    fetch: (url, params) => { seen.push(params.followRedirects); return { getResponseCode: () => 302, getContentText: () => '', getAllHeaders: () => ({ location: 'https://storage.example/r.csv' }) }; },
    sleep: () => {}
  });
  assert.deepStrictEqual(http.request({ method: 'GET', url: 'https://live.trading212.com/r', text: true, followRedirects: false }),
    { redirect: 'https://storage.example/r.csv' });
  assert.deepStrictEqual(seen, [false]);
});

test('t212 client: a download follows redirects itself and sends the key only to Trading 212 hops', () => {
  const hops = {
    'https://live.trading212.com/reports/1': { redirect: '/files/1' },
    'https://live.trading212.com/files/1': { redirect: 'https://storage.example/r.csv?sig=1' },
    'https://storage.example/r.csv?sig=1': 'csv text',
    'https://storage.example/back': { redirect: 'https://live.trading212.com/x.csv' },
    'https://live.trading212.com/x.csv': 'own text',
    'https://live.trading212.com/insecure': { redirect: 'http://storage.example/r.csv' },
    'https://live.trading212.com/nowhere': { redirect: '' },
    'https://live.trading212.com/loop': { redirect: '/loop' }
  };
  const requests = [];
  const http = { request: (req) => { requests.push(req); return hops[req.url]; } };
  const client = BS.T212Client.create(http, BS.T212Credentials.credentials(fakeProps(PROPS_INIT), CONN), base64);
  assert.strictEqual(client.download('https://live.trading212.com/reports/1'), 'csv text');
  assert.deepStrictEqual(requests.map((r) => [r.url, 'Authorization' in r.headers, r.followRedirects]), [
    ['https://live.trading212.com/reports/1', true, false],
    ['https://live.trading212.com/files/1', true, false],
    ['https://storage.example/r.csv?sig=1', false, false]
  ]);
  requests.length = 0;
  assert.strictEqual(client.download('https://storage.example/back'), 'own text');
  assert.deepStrictEqual(requests.map((r) => 'Authorization' in r.headers), [false, true]);
  assert.throws(() => client.download('https://live.trading212.com/insecure'), /Unexpected Trading 212 report link/);
  assert.throws(() => client.download('https://live.trading212.com/nowhere'), /redirect has no usable target/);
  requests.length = 0;
  assert.throws(() => client.download('https://live.trading212.com/loop'), /redirects too often/);
  assert.strictEqual(requests.length, T212.MAX_REDIRECTS + 1);
});
