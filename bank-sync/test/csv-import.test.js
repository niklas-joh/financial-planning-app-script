// CSV import provider: parsing, the BIL format, stable ids, the Drive folder provider and a
// full Sync all run. Every file below is made up (fake account, card and people).
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');

const P = BS.CsvParse;
const NOW = Date.parse('2026-09-01T08:00:00Z');

const fakeProps = (init) => {
  const map = Object.assign({}, init);
  return {
    map,
    getProperty: (k) => (k in map ? map[k] : null),
    setProperty: (k, v) => { map[k] = v; },
    getKeys: () => Object.keys(map)
  };
};

const CONN = BS.CsvFolders.connections(fakeProps({ CSV_FOLDER_ALEX: 'folder-1' }))[0];

const BIL_HEAD = [
  'BILnet',
  'LIST OF ACCOUNT MOVEMENTS',
  'LU120000000000001234 Current Private',
  '',
  'Date;Value date;Name;Message;Amount'
];
const ROWS = {
  rent: '28/08/2026;31/08/2026;Payment order;Standing order Remittance Info: Rent Beneficiary: DEMO PERSON  Beneficiary account: LU990000000000005678 At: DEMOLULL - DEMO BANK;-824.50',
  pay: '26/08/2026;26/08/2026;Credit;"Credit Remittance Info: SALARY By order of: DEMO EMPLOYER  Paying bank: DEMOLULL - DEMO BANK ";8,828.03',
  card: '14/08/2026;14/08/2026;Card payment;DEMO SHOP LUX 12/08/2026 10:58 4000000000001234;-9.99',
  instant: '07/08/2026;07/08/2026;Payment order;Instant Payment Remittance Info: Savings;-1,500',
  july: '31/07/2026;31/07/2026;Payment order;Standing order Remittance Info: Rent;-824.50'
};
const bil = (...rows) => BIL_HEAD.concat(rows).join('\r\n') + '\r\n';
const utf8 = (text) => () => text;

// --- Parsing ---

test('csv parse: quoted fields keep separators, quotes and line breaks; BOM and CRLF are handled', () => {
  assert.deepStrictEqual(P.rows('\uFEFFa;b\r\n"x;y";"say ""hi"""\n"two\nlines";\n', ';'),
    [['a', 'b'], ['x;y', 'say "hi"'], ['two\nlines', '']]);
  assert.deepStrictEqual(P.rows('a;b', ';'), [['a', 'b']]);
  assert.deepStrictEqual(P.rows("it's;5", ';'), [["it's", '5']]);
  assert.throws(() => P.rows('a;"open\n', ';'), /quoted field is not closed/);
});

test('csv parse: the header is found below a preamble; blank rows drop; extra fields are an error', () => {
  const t = P.table(P.rows('title\n\nA;B;\n1;2\n;\n3;4;\n', ';'), ['A', 'B']);
  assert.deepStrictEqual(t.preamble, [['title'], ['']]);
  assert.deepStrictEqual(t.header, ['A', 'B']);
  assert.deepStrictEqual(t.records, [{ A: '1', B: '2' }, { A: '3', B: '4' }]);
  assert.strictEqual(P.table(P.rows('X;Y\n1;2\n', ';'), ['A']), null);
  assert.throws(() => P.table(P.rows('A;B\n1;2;3\n', ';'), ['A', 'B']), /Row 2 has more fields than the header/);
});

test('csv parse: numbers follow the declared marks; a misplaced separator is an error, not a wrong value', () => {
  const dot = { decimal: '.', thousands: ',' };
  assert.deepStrictEqual(['-824.50', '8,828.03', '-1,500', '+12', ' 1 234.5 ', '1\u00A0234\u202F567', '\u22125', '1,234,567.1'].map((x) => P.number(x, dot)),
    [-824.5, 8828.03, -1500, 12, 1234.5, 1234567, -5, 1234567.1]);
  assert.strictEqual(P.number('  ', dot), null);
  assert.strictEqual(P.number('1.234,56', { decimal: ',', thousands: '.' }), 1234.56);
  ['1,5', '12,34.0', 'abc', '1.2.3', '--1'].forEach((x) => assert.throws(() => P.number(x, dot), /Not a number/, x));
});

test('csv parse: dates follow the declared layout and must exist', () => {
  assert.strictEqual(P.date('28/08/2026', 'DD/MM/YYYY'), '2026-08-28');
  assert.strictEqual(P.date(' 1/2/2026 ', 'DD/MM/YYYY'), '2026-02-01');
  assert.strictEqual(P.date('2026-08-28', 'YYYY-MM-DD'), '2026-08-28');
  assert.throws(() => P.date('31/02/2026', 'DD/MM/YYYY'), /Not a real date/);
  assert.throws(() => P.date('2026-08-28', 'DD/MM/YYYY'), /Not a date/);
});

// --- BIL format and normalization ---

const normalize = (text) => BS.CsvNormalize.file(BS.CsvNormalize.detect(utf8(text)), CONN);

test('csv bil: the account comes from the line above the header; the number is hashed and masked', () => {
  const { account } = normalize(bil(ROWS.rent));
  assert.match(account.accountId, /^bil-[0-9a-f]{8}$/);
  assert.deepStrictEqual([account.provider, account.bank, account.connection, account.accountName, account.currency, account.ibanMasked],
    ['csv', 'BIL', 'csv-alex', 'Current Private', 'EUR', '****1234']);
  assert.throws(() => normalize(BIL_HEAD.slice(3).concat(ROWS.rent).join('\n')), /no account line/);
});

test('csv bil: signed amounts with thousands commas, dates, counterparty, numbers masked everywhere stored', () => {
  const t = normalize(bil(ROWS.rent, ROWS.pay, ROWS.card, ROWS.instant)).transactions;
  assert.deepStrictEqual(t.map((x) => [x.bookingDate, x.valueDate, x.amount, x.currency, x.status]), [
    ['2026-08-28', '2026-08-31', -824.5, 'EUR', 'booked'],
    ['2026-08-26', '2026-08-26', 8828.03, 'EUR', 'booked'],
    ['2026-08-14', '2026-08-14', -9.99, 'EUR', 'booked'],
    ['2026-08-07', '2026-08-07', -1500, 'EUR', 'booked']
  ]);
  assert.deepStrictEqual(t.map((x) => x.counterparty), ['DEMO PERSON', 'DEMO EMPLOYER', 'DEMO SHOP LUX', '']);
  assert.match(t[0].description, /^Payment order: Standing order .* Beneficiary account: \*\*\*\*5678 At:/);
  assert.match(t[2].description, /10:58 \*\*\*\*1234$/);
  const stored = JSON.stringify(t);
  assert.ok(!stored.includes('4000000000001234') && !stored.includes('LU990000000000005678') && !stored.includes('LU120000000000001234'));
  assert.strictEqual(t[2].apiFields.Message, 'DEMO SHOP LUX 12/08/2026 10:58 ****1234');
  assert.strictEqual(t[1].apiFields.Amount, '8,828.03');
});

test('csv bil: ids are stable across files; identical rows stay distinct; a bad row names its record only', () => {
  const a = normalize(bil(ROWS.rent, ROWS.pay)).transactions.map((x) => x.transactionId);
  const b = normalize(bil(ROWS.card, ROWS.rent, ROWS.pay, ROWS.july)).transactions.map((x) => x.transactionId);
  assert.deepStrictEqual(b.slice(1, 3), a);
  assert.match(a[0], /^h20260828-[0-9a-f]{8}-1$/);
  const twice = normalize(bil(ROWS.rent, ROWS.rent)).transactions.map((x) => x.transactionId);
  assert.deepStrictEqual(twice.map((id) => id.slice(-2)), ['-1', '-2']);
  assert.strictEqual(twice[0], a[0]);
  assert.throws(() => normalize(bil(ROWS.rent, '28/08/2026;28/08/2026;Credit;x;12,34')), /Record 2: Not a number/);
});

test('csv detect: unknown headers are not guessed; an invalid UTF-8 decode falls back to Windows-1252', () => {
  assert.strictEqual(BS.CsvNormalize.detect(utf8('Datum;Belopp\n2026-01-01;5\n')), null);
  const asked = [];
  const decode = (charset) => { asked.push(charset); return charset === 'UTF-8' ? bil(ROWS.rent).replace('Rent', 'R\uFFFDnt') : bil(ROWS.rent); };
  const found = BS.CsvNormalize.detect(decode);
  assert.strictEqual(found.format.id, 'bil');
  assert.deepStrictEqual(asked, ['UTF-8', 'windows-1252']);
});

// --- Provider over a Drive folder, and a full Sync all ---

const file = (name, text) => ({ name, size: text.length, text: () => text });

function folderRun(files, init) {
  const props = fakeProps(Object.assign({ CSV_FOLDER_ALEX: 'folder-1' }, init));
  let rows = [];
  const raw = {};
  const env = { driveFolder: (id) => { assert.strictEqual(id, 'folder-1'); return { files: () => files }; } };
  const wired = { props, nowMs: () => NOW, connections: { all: () => [] }, store: { loadSession: () => null } };
  const deps = {
    connections: BS.ProviderRegistry.connections(wired),
    isConnected: (c) => BS.ProviderRegistry.isConnected(c, wired),
    providerFor: (c) => BS.ProviderRegistry.create(c, wired, env),
    store: { load: () => rows, save: (r) => { rows = r; return r.length; } },
    rawWriter: { writeTable: (tab, cols, r) => { raw[tab] = r; return r.length; } },
    nowMs: () => NOW,
    todayIso: () => '2026-09-01'
  };
  return { props, wired, deps, rows: () => rows, raw, run: () => BS.SyncAll.run(deps) };
}

test('csv provider: Sync all imports the folder; the same files again add and update nothing', () => {
  const files = [file('bil-august.csv', bil(ROWS.rent, ROWS.pay, ROWS.card)), file('notes.txt', 'not a csv')];
  const r = folderRun(files);
  const first = r.run();
  assert.deepStrictEqual(first.results.map((x) => [x.connection, x.status, x.transactions.added, x.transactions.updated]),
    [['csv-alex', 'ok', 3, 0]]);
  assert.deepStrictEqual(r.rows().map((t) => [t.bank, t.connection, t.accountName, t.ibanMasked]),
    Array(3).fill(['BIL', 'csv-alex', 'Current Private', '****1234']));
  assert.ok(r.rows().every((t) => !('apiFields' in t)));
  assert.strictEqual(r.raw.raw_transactions[0]['api.Name'], 'Payment order');
  assert.strictEqual(r.raw.raw_accounts.length, 1);
  const again = r.run();
  assert.deepStrictEqual([again.results[0].transactions.added, again.results[0].transactions.updated], [0, 0]);
  assert.strictEqual(r.rows().length, 3);
});

test('csv provider: an overlapping file adds only its new rows, also when both files stay in the folder', () => {
  const files = [file('a.csv', bil(ROWS.rent, ROWS.pay))];
  const r = folderRun(files);
  r.run();
  files.push(file('b.csv', bil(ROWS.pay, ROWS.card, ROWS.instant)));
  const second = r.run();
  assert.deepStrictEqual([second.results[0].transactions.added, second.results[0].transactions.updated], [2, 0]);
  assert.strictEqual(r.rows().length, 4);
  const record = JSON.parse(r.props.map.CSV_IMPORTS_ALEX);
  assert.deepStrictEqual(record.sources.bil, { files: 2, rows: 4, from: '2026-08-07', to: '2026-08-28' });
});

test('csv provider: an unreadable or unknown file is skipped and named; the others still import', () => {
  const files = [
    file('good.csv', bil(ROWS.rent)),
    file('LU120000000000001234-bad.csv', bil('28/08/2026;28/08/2026;Credit;x;not-a-number')),
    file('other.csv', 'Datum;Belopp\n2026-01-01;5\n')
  ];
  const r = folderRun(files);
  const out = r.run();
  assert.strictEqual(out.results[0].transactions.added, 1);
  const record = JSON.parse(r.props.map.CSV_IMPORTS_ALEX);
  assert.deepStrictEqual(record.skipped, [
    '[REDACTED:iban]-bad.csv (Record 1: Not a number in the expected format)',
    'other.csv (its header matches no known export)'
  ]);
  const lines = BS.CsvFolders.statusLines(r.props);
  assert.deepStrictEqual(lines, [
    'CSV import (alex) - CSV_FOLDER_ALEX: set; last read: 2026-09-01 08:00 UTC',
    '  bil: 1 file, 1 transactions, 2026-08-28 to 2026-08-28',
    '  skipped [REDACTED:iban]-bad.csv (Record 1: Not a number in the expected format)',
    '  skipped other.csv (its header matches no known export)'
  ]);
  assert.ok(!lines.join().includes('824'));
});

test('csv provider: EUR rows get rate 1 at their booking date through the normal FX step', () => {
  const r = folderRun([file('a.csv', bil(ROWS.rent))]);
  r.deps.fx = BS.FxService.create({
    fetchRates: () => { throw new Error('EUR needs no rates'); },
    rates: { load: () => ({}), save: () => 0 },
    cache: { get: () => null, put: () => {}, remove: () => {} },
    todayIso: () => '2026-09-01'
  });
  const out = r.run();
  assert.strictEqual(out.fx.changed, 1);
  assert.deepStrictEqual([r.rows()[0].amountEur, r.rows()[0].fxRate, r.rows()[0].fxRateDate], [-824.5, 1, '2026-08-28']);
});

test('csv registry: one connection per folder property; status without a folder; no consents; authorize explains', () => {
  const props = fakeProps({ CSV_FOLDER_ALEX: 'f', CSV_FOLDER_SAM: '', T212_API_KEY_ALEX: 'k' });
  const wired = { props, nowMs: () => NOW, connections: { all: () => [] }, store: { loadSession: () => null } };
  const R = BS.ProviderRegistry;
  const conns = R.connections(wired);
  assert.deepStrictEqual(conns.map((c) => c.key), ['trading212-alex', 'csv-alex', 'csv-sam']);
  assert.deepStrictEqual(conns.slice(1).map((c) => R.isConnected(c, wired)), [true, false]);
  assert.deepStrictEqual(R.consents(wired), []);
  assert.deepStrictEqual(BS.CsvFolders.statusLines(props), [
    'CSV import (alex) - CSV_FOLDER_ALEX: set; last read: never',
    'CSV import (sam) - CSV_FOLDER_SAM: MISSING; last read: never'
  ]);
  assert.deepStrictEqual(BS.CsvFolders.statusLines(fakeProps({})), ['CSV import: no folder set (optional; see the README)']);
  const provider = R.create(conns[1], wired, { driveFolder: () => ({ files: () => [] }) });
  assert.throws(() => provider.authorize(), /set CSV_FOLDER_ALEX to its id/);
  assert.deepStrictEqual(provider.listAccounts(), []);
});
