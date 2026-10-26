// CSV import provider: parsing, the BIL format, stable ids, the Drive folder provider and a
// full Sync all run. Every file below is made up (fake account, card and people).
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');
const { fakeSheetBook } = require('./fake-sheet');

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
  const comma = { decimal: ',', thousands: '' };
  assert.deepStrictEqual(['-12,345678', '4500,5', '-0,25', '700', '-12,000001'].map((x) => P.number(x, comma)),
    [-12.345678, 4500.5, -0.25, 700, -12.000001]);
  ['-570.00', '1.5', '1 234,5,6', ',5', '5,'].forEach((x) => assert.throws(() => P.number(x, comma), /Not a number/, x));
  ['.5', '5.', '1.5,0'].forEach((x) => assert.throws(() => P.number(x, dot), /Not a number/, x));
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
  const { accounts } = normalize(bil(ROWS.rent));
  assert.strictEqual(accounts.length, 1);
  const account = accounts[0];
  assert.match(account.accountId, /^bil-[0-9a-f]{8}$/);
  assert.strictEqual(account.accountId, 'bil-' + BS.Hash.fnv1a('LU120000000000001234'));
  // A file with a header and no rows still names its account.
  assert.strictEqual(normalize(bil()).accounts[0].accountId, account.accountId);
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
  assert.deepStrictEqual(t.map((x) => [x.category, x.counterparty, x.description, x.counterpartyAccount, x.counterpartyBank, x.reference]), [
    ['STANDING_ORDER', 'DEMO PERSON', 'Rent', '****5678', 'DEMOLULL - DEMO BANK', ''],
    ['CREDIT', 'DEMO EMPLOYER', 'SALARY', '', 'DEMOLULL - DEMO BANK', ''],
    ['CARD_PAYMENT', 'DEMO SHOP LUX', 'DEMO SHOP LUX 12/08/2026 10:58 ****1234', '', '', ''],
    ['INSTANT_PAYMENT', '', 'Savings', '', '', '']
  ]);
  assert.strictEqual(t[0].apiFields['parsed.Beneficiary account'], '****5678');
  assert.strictEqual(t[0].apiFields['parsed.Remittance Info'], 'Rent');
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

test('csv bil: the message is split on BIL labels only; a colon inside the remittance stays in it', () => {
  const msg = 'Credit Remittance Info: REFUND - 18 BENEFICIAIRE: 1111111111111 DATE CISS: 19.08.2026 By order of: DEMO FUND ' +
    'Paying bank: DEMOLULL - DEMO BANK End-to-end Identification: 22222222222222222 Payment initiation reference: NSCT0001';
  const p = BS.CsvBil.parts(msg);
  assert.strictEqual(p.lead, 'Credit');
  assert.deepStrictEqual(p.fields, {
    'Remittance Info': 'REFUND - 18 BENEFICIAIRE: 1111111111111 DATE CISS: 19.08.2026',
    'By order of': 'DEMO FUND',
    'Paying bank': 'DEMOLULL - DEMO BANK',
    'End-to-end Identification': '22222222222222222',
    'Payment initiation reference': 'NSCT0001'
  });
  const row = normalize(bil('20/08/2026;20/08/2026;Credit;"' + msg + ' ";517.34')).transactions[0];
  assert.deepStrictEqual([row.category, row.counterparty, row.counterpartyBank, row.reference],
    ['CREDIT', 'DEMO FUND', 'DEMOLULL - DEMO BANK', '****2222']);
  assert.strictEqual(row.description, 'REFUND - 18 BENEFICIAIRE: ****1111 DATE CISS: 19.08.2026');
  // Without any label: the whole message is the description and the BIL type the category.
  assert.deepStrictEqual(BS.CsvBil.describe('Payment order', 'Fee for July').description, 'Fee for July');
  assert.deepStrictEqual(BS.CsvBil.describe('Payment order', 'Fee for July').category, 'PAYMENT_ORDER');
});

test('csv bil: a label typed inside the remittance stays in it; the real labels after it still count', () => {
  const typed = BS.CsvBil.parts('Standing order Remittance Info: Meet At: the club Beneficiary: DEMO PERSON ' +
    'Beneficiary account: LU990000000000005678 At: DEMOLULL - DEMO BANK');
  assert.deepStrictEqual(typed.fields, {
    'Remittance Info': 'Meet At: the club',
    Beneficiary: 'DEMO PERSON',
    'Beneficiary account': 'LU990000000000005678',
    At: 'DEMOLULL - DEMO BANK'
  });
  const twice = BS.CsvBil.parts('Credit Remittance Info: re Beneficiary: someone By order of: DEMO FUND Paying bank: DEMOLULL - DEMO BANK');
  assert.deepStrictEqual(twice.fields, {
    'Remittance Info': 're Beneficiary: someone',
    'By order of': 'DEMO FUND',
    'Paying bank': 'DEMOLULL - DEMO BANK'
  });
  assert.strictEqual(twice.lead, 'Credit');
});

test('csv bil: ids depend on the raw columns only, so a new way of splitting the message keeps them', () => {
  const id = normalize(bil(ROWS.rent)).transactions[0].transactionId;
  const print = JSON.stringify(['2026-08-28', '2026-08-31', 'Payment order',
    'Standing order Remittance Info: Rent Beneficiary: DEMO PERSON Beneficiary account: LU990000000000005678 At: DEMOLULL - DEMO BANK', -824.5]);
  assert.strictEqual(id, 'h20260828-' + BS.Hash.fnv1a(print) + '-1');
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

// --- Trading 212 report: details for the rows the API already stores ---

const REPORT_HEAD = 'Action,Time (UTC),Notes,ID,Total,Currency (Total),Currency conversion fee,Currency (Currency conversion fee),Merchant name,Merchant category';
const REPORT = [
  REPORT_HEAD,
  'Spending cashback,2026-08-19 01:12:04+00:00,,id-cashback-1,0.31,"EUR",,,,',
  'Card debit,2026-08-21 09:27:53+00:00,,id-card-1,-53.46,"EUR",0.00,"USD","DEMO CAFE","RESTAURANTS"',
  'Deposit,2026-08-22 09:15:28+00:00,"Bank Transfer",id-deposit-1,100.00,"EUR",,,,',
  'Card debit,2026-08-22 09:25:18+00:00,,id-card-2,-20.00,"EUR",,,"DEMO SCHOOL, LUX","SERVICE_PROVIDERS"'
].join('\n') + '\n';

test('csv t212 report: every row gives a category (merchant category, else the action) and a counterparty (merchant, else notes)', () => {
  const found = BS.CsvNormalize.detect(utf8(REPORT));
  assert.deepStrictEqual([found.format.id, found.format.kind, found.format.enriches], ['trading212-report', 'details', 'trading212']);
  assert.deepStrictEqual(BS.CsvNormalize.details(found), [
    { provider: 'trading212', transactionId: 'id-cashback-1', counterparty: '', category: 'SPENDING_CASHBACK', classification: 'card', date: '2026-08-19' },
    { provider: 'trading212', transactionId: 'id-card-1', counterparty: 'DEMO CAFE', category: 'RESTAURANTS', classification: 'card', date: '2026-08-21' },
    { provider: 'trading212', transactionId: 'id-deposit-1', counterparty: 'Bank Transfer', category: 'DEPOSIT', classification: 'investment', date: '2026-08-22' },
    { provider: 'trading212', transactionId: 'id-card-2', counterparty: 'DEMO SCHOOL, LUX', category: 'SERVICE_PROVIDERS', classification: 'card', date: '2026-08-22' }
  ]);
  const other = REPORT_HEAD + '\nCard debit,2026-08-21 09:27:53+00:00,,id-card-3,-1.00,"EUR",,,,\nInterest on cash,2026-08-21 01:00:00+00:00,,id-int-1,0.01,"EUR",,,,\n';
  assert.deepStrictEqual(BS.CsvNormalize.details(BS.CsvNormalize.detect(utf8(other))).map((d) => [d.counterparty, d.category, d.classification]),
    [['', 'CARD_DEBIT', 'card'], ['', 'INTEREST_ON_CASH', 'investment']]);
  const short = 'Action,Time (UTC),Notes,ID,Total,Currency (Total)\nSpending cashback,2026-08-23 01:08:44+00:00,,id-cashback-2,0.80,"EUR"\n';
  assert.deepStrictEqual(BS.CsvNormalize.details(BS.CsvNormalize.detect(utf8(short))).map((d) => d.transactionId), ['id-cashback-2']);
  const noId = REPORT_HEAD + '\nCard debit,2026-08-21 09:27:53+00:00,,,-1.00,"EUR",,,"X","Y"\n';
  assert.throws(() => BS.CsvNormalize.details(BS.CsvNormalize.detect(utf8(noId))), /Record 1: A Trading 212 report row has no ID/);
});

test('csv t212 report: Sync all adds merchant and category to the API rows, adds no rows, and later API syncs keep them', () => {
  const api = (id, type, amount) => ({ provider: 'trading212', accountId: '1234567', transactionId: id, bookingDate: '2026-08-21', valueDate: '2026-08-21',
    status: 'booked', amount, currency: 'EUR', amountEur: null, fxRate: null, fxRateDate: '', counterparty: '', description: type });
  // Built by the real module from a property name, as in production.
  const t212 = BS.T212Credentials.connections(fakeProps({ T212_API_KEY_ALEX: 'x' }))[0];
  const r = folderRun([file('report.csv', REPORT), file('bil.csv', bil(ROWS.rent))]);
  r.deps.store = BS.TransactionsStore.create(fakeSheetBook(), (d) => d.toISOString().slice(0, 10));
  const csvProviderFor = r.deps.providerFor;
  r.deps.connections = [t212].concat(r.deps.connections);
  r.deps.isConnected = (c) => c === t212 || BS.ProviderRegistry.isConnected(c, r.wired);
  r.deps.providerFor = (c) => (c !== t212 ? csvProviderFor(c) : {
    listAccounts: () => [{ provider: 'trading212', accountId: '1234567', bank: 'Trading 212', connection: t212.key, currency: 'EUR' }],
    getBalances: () => [],
    maxWindowDays: Infinity,
    getTransactions: () => [api('id-card-1', 'WITHDRAW', -53.46), api('id-cashback-1', 'DEPOSIT', 0.31), api('id-deposit-1', 'DEPOSIT', 100)]
  });

  const first = r.run();
  assert.deepStrictEqual(first.results.map((x) => [x.connection, x.transactions.added]), [['trading212-alex', 3], ['csv-alex', 1]]);
  assert.deepStrictEqual(first.details, { read: 4, enriched: 3 });
  assert.match(BS.SyncAll.summarize(first), /Details from reports: 3 transactions updated \(4 read\)\./);
  const byId = () => Object.fromEntries(r.deps.store.load().map((t) => [t.transactionId, t]));
  assert.strictEqual(r.deps.store.load().length, 4);
  assert.deepStrictEqual([byId()['id-card-1'].counterparty, byId()['id-card-1'].category, byId()['id-card-1'].description],
    ['DEMO CAFE', 'RESTAURANTS', 'WITHDRAW']);
  assert.deepStrictEqual([byId()['id-cashback-1'].category, byId()['id-deposit-1'].category, byId()['id-deposit-1'].counterparty],
    ['SPENDING_CASHBACK', 'DEPOSIT', 'Bank Transfer']);

  const second = r.run();
  assert.deepStrictEqual(second.results.map((x) => [x.transactions.added, x.transactions.updated]), [[0, 0], [0, 0]]);
  assert.deepStrictEqual(second.details, { read: 4, enriched: 0 });
  assert.strictEqual(byId()['id-card-1'].counterparty, 'DEMO CAFE');

  const record = JSON.parse(r.props.map.CSV_IMPORTS_ALEX);
  assert.deepStrictEqual(record.sources['trading212-report'], { files: 1, rows: 4, from: '2026-08-19', to: '2026-08-22' });
});

test('csv t212 report (B1): unlabelled API rows wait in unclassified_transactions; Import CSV files moves them by the report', () => {
  const api = (id, type, amount) => ({ provider: 'trading212', accountId: '1234567', transactionId: id, bookingDate: '2026-08-21', valueDate: '2026-08-21',
    status: 'booked', amount, currency: 'EUR', amountEur: null, fxRate: null, fxRateDate: '', counterparty: '', description: type });
  const T = BS.Config.TABS;
  const book = fakeSheetBook();
  const iso = (d) => d.toISOString().slice(0, 10);
  const tab = (name) => BS.TransactionsStore.create(book, iso, name).load().map((t) => t.transactionId).sort();
  const routedStore = () => {
    const stores = {};
    BS.TxnRouting.tabs().forEach((name) => { stores[name] = BS.TransactionsStore.create(book, iso, name); });
    return BS.TxnRouting.createStore(stores);
  };
  const t212 = BS.T212Credentials.connections(fakeProps({ T212_API_KEY_ALEX: 'x' }))[0];
  const files = [file('bil.csv', bil(ROWS.rent))];
  const r = folderRun(files);
  const csvProviderFor = r.deps.providerFor;
  const csvConnections = r.deps.connections;
  r.deps.connections = [t212].concat(csvConnections);
  r.deps.isConnected = (c) => c === t212 || BS.ProviderRegistry.isConnected(c, r.wired);
  r.deps.providerFor = (c) => (c !== t212 ? csvProviderFor(c) : {
    listAccounts: () => [{ provider: 'trading212', accountId: '1234567', bank: 'Trading 212', connection: t212.key, currency: 'EUR' }],
    getBalances: () => [],
    maxWindowDays: Infinity,
    getTransactions: () => [api('id-card-1', 'WITHDRAW', -53.46), api('id-cashback-1', 'DEPOSIT', 0.31), api('id-deposit-1', 'DEPOSIT', 100),
      api('int-1', 'INTEREST_ON_FREE_CASH', 0.02)]
  });

  // Daily sync before any report: the API cannot tell a card payment from a transfer.
  r.deps.store = routedStore();
  r.run();
  assert.deepStrictEqual(tab(T.UNCLASSIFIED_TRANSACTIONS), ['id-card-1', 'id-cashback-1', 'id-deposit-1']);
  assert.deepStrictEqual(tab(T.INVESTMENT_TRANSACTIONS), ['int-1']);
  assert.strictEqual(tab(T.TRANSACTIONS).length, 1);

  // The report lands; Import CSV files runs only the CSV connections and still finds the rows.
  files.push(file('report.csv', REPORT));
  r.deps.connections = csvConnections;
  r.deps.store = routedStore();
  const imported = r.run();
  assert.deepStrictEqual(imported.details, { read: 4, enriched: 3 });
  assert.deepStrictEqual(tab(T.UNCLASSIFIED_TRANSACTIONS), []);
  assert.deepStrictEqual(tab(T.INVESTMENT_TRANSACTIONS), ['id-deposit-1', 'int-1']);
  assert.strictEqual(tab(T.TRANSACTIONS).length, 3);
  assert.ok(['id-card-1', 'id-cashback-1'].every((id) => tab(T.TRANSACTIONS).includes(id)));
  assert.match(BS.SyncAll.summarize(imported), /Moved between tabs: 2 to transactions, 1 to investment_transactions\./);

  // The next daily sync re-reads the API rows: nothing added, updated or moved.
  r.deps.connections = [t212].concat(csvConnections);
  r.deps.store = routedStore();
  const again = r.run();
  assert.deepStrictEqual(again.results.map((x) => [x.transactions.added, x.transactions.updated]), [[0, 0], [0, 0]]);
  assert.strictEqual(BS.TxnRouting.movedText(again.moved), '');
  assert.strictEqual(tab(T.TRANSACTIONS).length + tab(T.INVESTMENT_TRANSACTIONS).length + tab(T.UNCLASSIFIED_TRANSACTIONS).length, 5);
});

test('csv provider: the last-read record names only the first skipped files, so it fits one Script Property', () => {
  const bad = Array.from({ length: 200 }, (_, i) => file('export-' + 'x'.repeat(300) + '-' + i + '.csv', 'Datum;Belopp\n2026-01-01;5\n'));
  const r = folderRun([file('good.csv', bil(ROWS.rent))].concat(bad));
  const out = r.run();
  assert.strictEqual(out.results[0].transactions.added, 1);
  const stored = r.props.map.CSV_IMPORTS_ALEX;
  assert.ok(stored.length < 9000, 'record is ' + stored.length + ' characters');
  const record = JSON.parse(stored);
  assert.strictEqual(record.skipped.length, BS.Config.CSV.MAX_SKIPPED_LISTED);
  assert.ok(record.skipped.every((s) => s.length <= BS.Config.CSV.MAX_SKIPPED_CHARS));
  assert.strictEqual(record.skippedMore, 190);
  assert.strictEqual(BS.CsvFolders.statusLines(r.props).pop(), '  and 190 more skipped files');
});

// --- Avanza (one file, several accounts; rows go to investment_transactions). Made-up accounts, funds and ISINs. ---

const AVANZA_HEADER = 'Datum;Konto;Typ av transaktion;Värdepapper/beskrivning;Antal;Kurs;Belopp;Transaktionsvaluta;Courtage;Valutakurs;Instrumentvaluta;ISIN;Resultat';
const AV = {
  sale: '2026-08-20;Demo ISK;Sälj;Demo Global Fund A;-10,5;200,25;2102,63;SEK;;;SEK;XS0000000011;150,123456',
  saleEarlier: '2026-06-02;Demo ISK;Sälj;Demo Small Caps;-1,25;80;100;SEK;;;SEK;XS0000000029;-4,5',
  buy: '2026-08-18;Demo ISK;Köp;Demo Global Fund A;3;190;-570;SEK;;;SEK;XS0000000011;',
  deposit: '2026-08-17;Demo ISK;Insättning;Insättning från bankkonto;;;1000;SEK;;;;;',
  dividend: '2026-08-11;Demo Pension;Utdelning;Demo EM Fund;120,5;0,25;30,13;SEK;;;SEK;XS0000000037;',
  reinvest: '2026-08-11;Demo Pension;Återinv. utdelning;Demo EM Fund;0,1;301,3;-30,13;SEK;;;SEK;XS0000000037;',
  risk: '2026-08-08;Demo Pension;Uttag av riskkostnad;Riskpremie;;;-0,5;SEK;;;;;',
  tax: '2026-07-27;Demo Pension;Avkastningsskatt;Avkastningsskatt;;;-100;SEK;;;;;',
  withholding: '2026-07-15;Demo ISK;Utländsk källskatt;Demo US Shares;;;-1,2;SEK;;;USD;XS0000000045;',
  unknown: '2026-07-10;Demo ISK;Byte av fond;Demo Global Fund A;5;;;SEK;;;SEK;XS0000000011;'
};
const avanza = (...rows) => '﻿' + [AVANZA_HEADER].concat(rows).join('\n') + '\n';

test('csv avanza: one file holds both accounts; each account id is the source plus a hash of its name', () => {
  const out = normalize(avanza(AV.sale, AV.dividend, AV.risk, AV.buy));
  assert.deepStrictEqual(out.accounts.map((a) => [a.accountId, a.bank, a.accountName, a.currency, a.ibanMasked, a.provider, a.connection]), [
    ['avanza-' + BS.Hash.fnv1a('Demo ISK'), 'Avanza', 'Demo ISK', 'SEK', '', 'csv', 'csv-alex'],
    ['avanza-' + BS.Hash.fnv1a('Demo Pension'), 'Avanza', 'Demo Pension', 'SEK', '', 'csv', 'csv-alex']
  ]);
  assert.deepStrictEqual(out.transactions.map((t) => t.accountId), [0, 1, 1, 0].map((i) => out.accounts[i].accountId));
});

test('csv avanza: amount from Belopp as signed, type as description and code, security as counterparty, ISIN as reference', () => {
  const t = normalize(avanza(AV.sale, AV.buy, AV.deposit, AV.dividend, AV.reinvest, AV.risk, AV.tax, AV.withholding)).transactions;
  assert.deepStrictEqual(t.map((x) => [x.bookingDate, x.valueDate, x.amount, x.currency, x.status]), [
    ['2026-08-20', '', 2102.63, 'SEK', 'booked'],
    ['2026-08-18', '', -570, 'SEK', 'booked'],
    ['2026-08-17', '', 1000, 'SEK', 'booked'],
    ['2026-08-11', '', 30.13, 'SEK', 'booked'],
    ['2026-08-11', '', -30.13, 'SEK', 'booked'],
    ['2026-08-08', '', -0.5, 'SEK', 'booked'],
    ['2026-07-27', '', -100, 'SEK', 'booked'],
    ['2026-07-15', '', -1.2, 'SEK', 'booked']
  ]);
  assert.deepStrictEqual(t.map((x) => [x.description, x.category, x.counterparty, x.reference]), [
    ['Sälj', 'SALJ', 'Demo Global Fund A', 'XS0000000011'],
    ['Köp', 'KOP', 'Demo Global Fund A', 'XS0000000011'],
    ['Insättning', 'INSATTNING', 'Insättning från bankkonto', ''],
    ['Utdelning', 'UTDELNING', 'Demo EM Fund', 'XS0000000037'],
    ['Återinv. utdelning', 'ATERINV_UTDELNING', 'Demo EM Fund', 'XS0000000037'],
    ['Uttag av riskkostnad', 'UTTAG_AV_RISKKOSTNAD', 'Riskpremie', ''],
    ['Avkastningsskatt', 'AVKASTNINGSSKATT', 'Avkastningsskatt', ''],
    ['Utländsk källskatt', 'UTLANDSK_KALLSKATT', 'Demo US Shares', 'XS0000000045']
  ]);
  // Quantity, price, fee, exchange rate and result stay on the raw tab only, as written.
  assert.deepStrictEqual([t[0].apiFields.Antal, t[0].apiFields.Kurs, t[0].apiFields.Resultat, t[0].apiFields.Courtage, t[0].apiFields.Instrumentvaluta],
    ['-10,5', '200,25', '150,123456', '', 'SEK']);
  assert.ok(t.every((x) => x.counterpartyAccount === '' && x.counterpartyBank === ''));
});

test('csv parse: codes drop accents, so Swedish types become plain codes; English labels are unchanged', () => {
  assert.deepStrictEqual(['Sälj', 'Köp', 'Återinv. utdelning', 'Utländsk källskatt', 'Ränta', 'Övrigt', 'Spending cashback', 'Card payment']
    .map(P.code), ['SALJ', 'KOP', 'ATERINV_UTDELNING', 'UTLANDSK_KALLSKATT', 'RANTA', 'OVRIGT', 'SPENDING_CASHBACK', 'CARD_PAYMENT']);
});

test('csv avanza: ids are stable; the reinvested dividend pair never collides; identical rows count per account', () => {
  const ids = (text) => normalize(text).transactions.map((x) => x.transactionId);
  const pair = ids(avanza(AV.dividend, AV.reinvest));
  assert.notStrictEqual(pair[0], pair[1]);
  assert.ok(pair.every((id) => /^h20260811-[0-9a-f]{8}-1$/.test(id)));
  // The same row twice in one account: -1 and -2; the same row in the other account is that account's -1.
  const other = AV.risk.replace('Demo Pension', 'Demo ISK');
  const t = normalize(avanza(AV.risk, AV.risk, other)).transactions;
  assert.deepStrictEqual(t.map((x) => x.transactionId.slice(-2)), ['-1', '-2', '-1']);
  assert.notStrictEqual(t[0].accountId, t[2].accountId);
  // Columns Avanza may recompute or rename are not part of the id.
  const recomputed = AV.sale.replace('150,123456', '151').replace(';;;SEK;XS', ';1,5;;SEK;XS');
  assert.deepStrictEqual(ids(avanza(recomputed)), ids(avanza(AV.sale)));
  // A longer export ("Sedan start") and a later month give the same ids for the days they share.
  const all = ids(avanza(AV.sale, AV.buy, AV.dividend, AV.reinvest, AV.tax, AV.saleEarlier));
  assert.deepStrictEqual(ids(avanza(AV.sale, AV.buy, AV.dividend, AV.reinvest)), all.slice(0, 4));
});

test('csv avanza: an older header with "Courtage (SEK)" is read; an unknown type is imported as written', () => {
  const old = avanza(AV.sale).replace(';Courtage;', ';Courtage (SEK);');
  const found = BS.CsvNormalize.detect(utf8(old));
  assert.strictEqual(found.format.id, 'avanza');
  assert.strictEqual(BS.CsvNormalize.file(found, CONN).transactions[0].amount, 2102.63);
  // A type not seen before keeps its own sign and text; a blank amount is a movement without cash.
  const t = normalize(avanza(AV.unknown)).transactions[0];
  assert.deepStrictEqual([t.description, t.category, t.amount, t.counterparty], ['Byte av fond', 'BYTE_AV_FOND', 0, 'Demo Global Fund A']);
});

test('csv avanza: a row that cannot be read skips the file and names the record only', () => {
  assert.throws(() => normalize(avanza(AV.sale, AV.buy.replace('-570', '-570.00'))), /Record 2: Not a number/);
  assert.throws(() => normalize(avanza(AV.sale.replace('2026-08-20', '20/08/2026'))), /Record 1: Not a date/);
  assert.throws(() => normalize(avanza(AV.sale.replace(';Demo ISK;', ';;'))), /Record 1: An Avanza row has no account/);
  assert.throws(() => normalize(avanza(AV.sale.replace(';SEK;;;SEK;', ';;;;SEK;'))), /Record 1: An Avanza row has no currency/);
  assert.throws(() => normalize(avanza(AV.sale.replace(';SEK;;;SEK;', ';kr;;;SEK;'))), /Record 1: An Avanza row has no currency code/);
});

test('csv avanza: Sync all puts Avanza rows on investment_transactions and BIL rows on transactions; overlapping exports add nothing', () => {
  const T = BS.Config.TABS;
  const book = fakeSheetBook();
  const iso = (d) => d.toISOString().slice(0, 10);
  const tab = (name) => BS.TransactionsStore.create(book, iso, name).load();
  const routedStore = () => {
    const stores = {};
    BS.TxnRouting.tabs().forEach((name) => { stores[name] = BS.TransactionsStore.create(book, iso, name); });
    return BS.TxnRouting.createStore(stores);
  };
  const files = [file('bil.csv', bil(ROWS.rent)), file('avanza-sedan-start.csv', avanza(AV.sale, AV.dividend, AV.reinvest, AV.risk, AV.tax))];
  const r = folderRun(files);
  r.deps.store = routedStore();
  const first = r.run();
  assert.deepStrictEqual(first.results.map((x) => [x.transactions.added, x.transactions.updated]), [[6, 0]]);
  assert.deepStrictEqual(tab(T.TRANSACTIONS).map((t) => t.bank), ['BIL']);
  assert.deepStrictEqual(tab(T.INVESTMENT_TRANSACTIONS).map((t) => [t.bank, t.accountName, t.category]), [
    ['Avanza', 'Demo ISK', 'SALJ'],
    ['Avanza', 'Demo Pension', 'UTDELNING'],
    ['Avanza', 'Demo Pension', 'ATERINV_UTDELNING'],
    ['Avanza', 'Demo Pension', 'UTTAG_AV_RISKKOSTNAD'],
    ['Avanza', 'Demo Pension', 'AVKASTNINGSSKATT']
  ]);
  assert.strictEqual(tab(T.UNCLASSIFIED_TRANSACTIONS).length, 0);
  assert.deepStrictEqual(r.raw.raw_accounts.map((a) => a.accountName).sort(), ['Current Private', 'Demo ISK', 'Demo Pension']);
  assert.strictEqual(BS.TxnRouting.movedText(first.moved), '');

  // A later month's export overlapping the first: only its new row is added; nothing is updated or moved.
  files.push(file('avanza-august.csv', avanza(AV.sale, AV.buy, AV.dividend, AV.reinvest, AV.risk)));
  r.deps.store = routedStore();
  const second = r.run();
  assert.deepStrictEqual(second.results.map((x) => [x.transactions.added, x.transactions.updated]), [[1, 0]]);
  assert.strictEqual(BS.TxnRouting.movedText(second.moved), '');
  r.deps.store = routedStore();
  const third = r.run();
  assert.deepStrictEqual(third.results.map((x) => [x.transactions.added, x.transactions.updated]), [[0, 0]]);
  assert.strictEqual(tab(T.INVESTMENT_TRANSACTIONS).length, 6);

  // Show status: counts and dates per source only.
  const lines = BS.CsvFolders.statusLines(r.props);
  assert.deepStrictEqual(lines.slice(1), [
    '  avanza: 2 files, 6 transactions, 2026-07-27 to 2026-08-20',
    '  bil: 1 file, 1 transactions, 2026-08-28 to 2026-08-28'
  ]);
  assert.ok(!lines.join().includes('Demo') && !lines.join().includes('2102'));
});
