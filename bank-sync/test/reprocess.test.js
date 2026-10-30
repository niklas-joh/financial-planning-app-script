// MASK_ACCOUNT_NUMBERS and Re-process stored rows: the one masking helper, full values when the flag is
// off, ids that never depend on it, and a re-process over Enable Banking and CSV rows that rewrites them in
// place (same row counts on every tab, run twice). All data below is made up.
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load.js');
const { withMasking } = require('./masking');
const { fakeSheetBook } = require('./fake-sheet');

const NOW = Date.parse('2026-09-01T08:00:00Z');
const TODAY = '2026-09-01';
const IBAN = 'LU280019400644750000';
const OTHER_IBAN = 'LU990000000000005678';

const memProps = (init) => {
  const map = Object.assign({}, init);
  return {
    map,
    getProperty: (k) => (k in map ? map[k] : null),
    setProperty: (k, v) => { map[k] = v; },
    deleteProperty: (k) => { delete map[k]; },
    getKeys: () => Object.keys(map)
  };
};

// --- The helper ---

test('masking: off by default; maskAccount returns the value unchanged when off and masks when on', () => {
  assert.strictEqual(BS.Config.MASK_ACCOUNT_NUMBERS, false);
  assert.strictEqual(BS.Redact.maskAccount(IBAN), IBAN);
  assert.strictEqual(BS.Redact.maskAccount(1234567890123), '1234567890123');
  [null, undefined, ''].forEach((v) => assert.strictEqual(BS.Redact.maskAccount(v), ''));
  withMasking(BS, true, () => {
    assert.strictEqual(BS.Redact.maskAccount(IBAN), '****0000');
    assert.strictEqual(BS.Redact.maskAccount(undefined), '');
  });
  // Log scrubbing does not depend on the flag.
  assert.strictEqual(BS.Redact.text('id ' + IBAN), 'id [REDACTED:iban]');
});

test('masking: isMasked finds values written while masking was on, and nothing else', () => {
  assert.ok(BS.Redact.isMasked('****8707'));
  assert.ok(BS.Redact.isMasked('REFUND BENEFICIAIRE: ****1439 DATE'));
  assert.ok(!BS.Redact.isMasked(IBAN));
  assert.ok(!BS.Redact.isMasked('5 * 4 stars'));
  assert.ok(!BS.Redact.isMasked(-8707));
});

// --- Every masking site keeps the full value when the flag is off ---

const BIL_HEAD = ['BILnet', 'LIST OF ACCOUNT MOVEMENTS', 'LU120000000000001234 Current Private', '', 'Date;Value date;Name;Message;Amount'];
const BIL = {
  rent: '28/08/2026;31/08/2026;Payment order;Standing order Remittance Info: Rent Beneficiary: DEMO PERSON  Beneficiary account: ' +
    OTHER_IBAN + ' At: DEMOLULL - DEMO BANK End-to-end Identification: 202608280000005545;-824.50',
  refund: '19/08/2026;19/08/2026;Credit;Credit Remittance Info: REFUND - 18 BENEFICIAIRE: LU120000000000001439 DATE CISS: 19.08.2026 ' +
    'By order of: DEMO FUND Paying bank: DEMOLULL - DEMO BANK;25.00',
  card: '14/08/2026;14/08/2026;Card payment;DEMO SHOP LUX 12/08/2026 10:58 4000000000001234;-9.99',
  july: '30/07/2026;30/07/2026;Payment order;Standing order Remittance Info: Old Beneficiary: DEMO LANDLORD Beneficiary account: ' +
    'LU110000000000008707 At: DEMOLULL - DEMO BANK;-700.00'
};
const bil = (...rows) => BIL_HEAD.concat(rows).join('\r\n') + '\r\n';
const CSV_CONN = BS.CsvFolders.connections(memProps({ CSV_FOLDER_ALEX: 'folder-1' }))[0];
const csvNormalize = (text) => BS.CsvNormalize.file(BS.CsvNormalize.detect(() => text), CSV_CONN);

test('masking off: the CSV account number, counterparty account, texts, reference and api.parsed fields stay whole', () => {
  const out = csvNormalize(bil(BIL.rent, BIL.refund, BIL.card));
  assert.strictEqual(out.accounts[0].ibanMasked, 'LU120000000000001234');
  const [rent, refund, card] = out.transactions;
  assert.strictEqual(rent.counterpartyAccount, OTHER_IBAN);
  assert.strictEqual(rent.reference, '202608280000005545');
  assert.strictEqual(rent.apiFields['parsed.Beneficiary account'], OTHER_IBAN);
  assert.strictEqual(rent.apiFields['parsed.End-to-end Identification'], '202608280000005545');
  assert.match(refund.description, /BENEFICIAIRE: LU120000000000001439 DATE/);
  assert.match(card.description, /4000000000001234/);
  assert.ok(!JSON.stringify(out).includes('****'));
});

test('masking on: the same fields are cut to their last four characters', () => withMasking(BS, true, () => {
  const out = csvNormalize(bil(BIL.rent, BIL.refund));
  assert.strictEqual(out.accounts[0].ibanMasked, '****1234');
  assert.strictEqual(out.transactions[0].counterpartyAccount, '****5678');
  assert.strictEqual(out.transactions[0].reference, '****5545');
  assert.strictEqual(out.transactions[0].apiFields['parsed.Beneficiary account'], '****5678');
  assert.match(out.transactions[1].description, /BENEFICIAIRE: \*\*\*\*1439 DATE/);
}));

const ebTxn = (o) => ({
  transaction_amount: { amount: '12.50', currency: 'EUR' }, credit_debit_indicator: 'DBIT', status: 'BOOK',
  creditor: { name: 'DEMO SHOP' }, creditor_account: { other: { identification: '12345678901234', scheme_name: 'BBAN' } },
  debtor_account_additional_identification: [{ identification: '98765432109876', scheme_name: 'BBAN' }],
  remittance_information: ['Order 1'], ...o
});
const ebAccount = () => ({
  uid: 'eb-1', currency: 'EUR', name: 'DEMO HOLDER', details: 'Main', cash_account_type: 'CACC',
  account_id: { iban: IBAN }, all_account_ids: [{ identification: IBAN, scheme_name: 'IBAN' }]
});

test('masking off: Enable Banking IBAN and account identifications stay whole', () => {
  const account = BS.EbNormalize.normalizeAccount(ebAccount(), '');
  assert.strictEqual(account.ibanMasked, IBAN);
  assert.strictEqual(account.apiFields['account_id.iban'], IBAN);
  assert.strictEqual(account.apiFields['all_account_ids[0].identification'], IBAN);
  const t = BS.EbNormalize.normalizeTransactions([ebTxn({ booking_date: '2026-08-20' })], 'eb-1').transactions[0];
  assert.strictEqual(t.apiFields['creditor_account.other.identification'], '12345678901234');
  assert.strictEqual(t.apiFields['debtor_account_additional_identification[0].identification'], '98765432109876');
  withMasking(BS, true, () => {
    const masked = BS.EbNormalize.normalizeTransactions([ebTxn({ booking_date: '2026-08-20' })], 'eb-1').transactions[0];
    assert.strictEqual(masked.apiFields['creditor_account.other.identification'], '****1234');
    assert.strictEqual(masked.apiFields['debtor_account_additional_identification[0].identification'], '****9876');
  });
});

test('ids never depend on the flag: CSV and Enable Banking transaction and account ids are the same both ways', () => {
  const ids = () => {
    const csv = csvNormalize(bil(BIL.rent, BIL.refund, BIL.card, BIL.card));
    const eb = BS.EbNormalize.normalizeTransactions([
      ebTxn({ booking_date: '2026-08-20' }), ebTxn({ booking_date: '2026-08-20' }), ebTxn({ booking_date: '2026-08-21', entry_reference: 'E1' })
    ], 'eb-1').transactions;
    return csv.accounts.map((a) => a.accountId).concat(csv.transactions.map((t) => t.transactionId), eb.map((t) => t.transactionId));
  };
  const off = ids();
  assert.deepStrictEqual(withMasking(BS, true, ids), off);
  assert.strictEqual(new Set(off).size, off.length);
});

test('enrich: a detail rewrites a value stored masked, and nothing else; idempotent both ways', () => {
  const row = { provider: 'trading212', transactionId: 'a', counterparty: 'SHOP ****1234', category: '' };
  const detail = { provider: 'trading212', transactionId: 'a', counterparty: 'SHOP 4000000000001234', category: 'GROCERIES' };
  const once = BS.TxnMerge.enrich([row], [detail]);
  assert.deepStrictEqual([once.rows[0].counterparty, once.rows[0].category], ['SHOP 4000000000001234', 'GROCERIES']);
  assert.strictEqual(BS.TxnMerge.enrich(once.rows, [detail]).enriched, 0);
  // With masking on the detail is masked too, so a masked row stays as it is.
  const masked = { ...detail, counterparty: 'SHOP ****1234', category: '' };
  assert.strictEqual(BS.TxnMerge.enrich([row], [masked]).enriched, 0);
});

// --- Full-history sync pieces ---

test('auth.refreshStale with force reads fresh details again', () => {
  const props = memProps();
  const store = BS.EbSessionStore.create(props, () => NOW);
  const connection = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' });
  store.saveSession(connection.key, { sessionId: 's', accessValidUntil: '2027-01-01T00:00:00Z',
    accounts: [{ accountId: 'eb-1', ibanMasked: '****0000', detailsAt: new Date(NOW).toISOString() }] });
  const asked = [];
  const client = { getAccountDetails: (uid) => { asked.push(uid); return ebAccount(); } };
  const auth = BS.EbAuth.create({ client, store, connections: BS.Connections.create(props), nowMs: () => NOW, uuid: () => 'u', getRedirectUrl: () => 'x' });
  auth.refreshStale(connection);
  assert.deepStrictEqual(asked, []);
  const accounts = auth.refreshStale(connection, undefined, true);
  assert.deepStrictEqual(asked, ['eb-1']);
  assert.strictEqual(accounts[0].ibanMasked, IBAN);
  assert.strictEqual(store.loadSession(connection.key).accounts[0].ibanMasked, IBAN);
});

test('sync fullHistory: starts at the oldest stored row; a window the source cannot serve is counted, not fatal', () => {
  const stored = [
    { provider: 'p', accountId: 'a1', transactionId: 'x', bookingDate: '2026-06-10', valueDate: '', status: 'booked', amount: -1, currency: 'EUR' },
    { provider: 'p', accountId: 'a1', transactionId: 'y', bookingDate: '2026-08-30', valueDate: '', status: 'booked', amount: -1, currency: 'EUR' }
  ];
  let rows = stored;
  const windows = [];
  const provider = {
    getTransactions: (id, range) => {
      windows.push(range.dateFrom);
      if (range.dateFrom < '2026-07-01') throw BS.Errors.http(400, 'too old');
      return [];
    }
  };
  const deps = (o) => ({ provider, accounts: [{ accountId: 'a1' }], store: { load: () => rows, save: (r) => { rows = r; } },
    nowMs: () => NOW, todayIso: () => TODAY, ...o });
  const out = BS.TxnSync.run(deps({ fullHistory: true }));
  assert.deepStrictEqual(windows, ['2026-06-10', '2026-07-11', '2026-08-11']);
  assert.strictEqual(out.unavailable, 1);
  assert.strictEqual(rows.length, 2);
  // Without fullHistory the same failure still stops the account, and it starts at the latest booked date minus the overlap.
  windows.length = 0;
  provider.getTransactions = (id, range) => { windows.push(range.dateFrom); throw BS.Errors.http(400, 'too old'); };
  assert.throws(() => BS.TxnSync.run(deps({})), /too old/);
  assert.deepStrictEqual(windows, ['2026-08-23']);
  provider.getTransactions = () => { throw BS.Errors.auth('Consent has expired'); };
  assert.throws(() => BS.TxnSync.run(deps({ fullHistory: true })), { name: 'AuthError' });
  assert.strictEqual(BS.TxnSync.earliestStored(stored, 'a1'), '2026-06-10');
  assert.strictEqual(BS.TxnSync.earliestStored(stored, 'nope'), '');
});

// --- Re-process over a whole sheet: Enable Banking plus a CSV folder ---

/** A Bank Sync sheet with one Enable Banking login and one CSV folder, wired like the container. */
function world() {
  const props = memProps({ CSV_FOLDER_ALEX: 'folder-1', T212_API_KEY_NIKLAS: 'k' });
  const book = fakeSheetBook();
  const ebStore = BS.EbSessionStore.create(props, () => NOW);
  const connections = BS.Connections.create(props);
  const ebConn = BS.Connections.build({ bankName: 'Revolut', country: 'LU', holder: 'niklas' });
  connections.save(ebConn);
  ebStore.saveSession(ebConn.key, { sessionId: 's', accessValidUntil: '2027-01-01T00:00:00Z',
    accounts: [{ provider: 'enablebanking', accountId: 'eb-1', currency: 'EUR', detailsAt: '' }] });
  const bank = {
    servedFrom: '2026-06-01',
    rows: [
      ebTxn({ booking_date: '2026-06-10', remittance_information: ['Early'] }),
      ebTxn({ booking_date: '2026-08-20', entry_reference: 'E1' })
    ]
  };
  const client = {
    getSession: () => ({ status: 'AUTHORIZED' }),
    getAccountDetails: () => ebAccount(),
    getBalances: () => ({ balances: [{ balance_amount: { amount: '10.00', currency: 'EUR' }, balance_type: 'CLBD', name: 'Booked' }] }),
    getTransactions: (uid, q) => {
      if (q.date_from < bank.servedFrom) throw BS.Errors.http(400, 'Wrong transactions period');
      return { transactions: bank.rows.filter((t) => t.booking_date >= q.date_from && t.booking_date <= q.date_to), continuation_key: null };
    }
  };
  const auth = BS.EbAuth.create({ client, store: ebStore, connections, nowMs: () => NOW, uuid: () => 'u', getRedirectUrl: () => 'x' });
  const wired = { props, nowMs: () => NOW, connections, store: ebStore, client, auth };
  const files = [
    { name: 'bil-august.csv', text: bil(BIL.rent, BIL.refund, BIL.card) },
    { name: 'bil-july.csv', text: bil(BIL.july) }
  ].map((f) => ({ name: f.name, size: f.text.length, text: () => f.text }));
  const env = { driveFolder: () => ({ files: () => files }) };
  const stores = {};
  BS.TxnRouting.tabs().forEach((tab) => { stores[tab] = BS.TransactionsStore.create(book, (d) => d.toISOString().slice(0, 10), tab); });
  const sync = () => ({
    connections: BS.ProviderRegistry.connections(wired),
    isConnected: (c) => BS.ProviderRegistry.isConnected(c, wired),
    providerFor: (c) => BS.ProviderRegistry.create(c, wired, env),
    store: BS.TxnRouting.createStore(stores),
    rawWriter: BS.RawWriter.create(book),
    nowMs: () => NOW,
    todayIso: () => TODAY
  });
  const reprocess = () => {
    const deps = sync();
    return BS.Reprocess.run({ sync: deps, rawReader: deps.rawWriter });
  };
  /** Data rows on each tab (header excluded). */
  const counts = () => {
    const out = {};
    Object.keys(book.sheets).sort().forEach((tab) => { out[tab] = Math.max(0, book.sheets[tab].getLastRow() - 1); });
    return out;
  };
  const read = (tab) => BS.RawWriter.create(book).readTable(tab, []);
  return { props, book, bank, files, sync, reprocess, counts, read };
}

const sortedIds = (rows) => rows.map((r) => r.provider + '|' + r.accountId + '|' + r.transactionId).sort();

test('re-process: rows synced with masking on are rewritten whole, in place; a second run changes nothing', () => {
  const w = world();
  withMasking(BS, true, () => BS.SyncAll.run(w.sync()));
  const before = w.counts();
  const idsBefore = sortedIds(w.sync().store.load());
  assert.ok(w.read('transactions').some((r) => r.ibanMasked === '****1234'));
  assert.deepStrictEqual(before, { raw_accounts: 2, raw_balances: 1, raw_transactions: 6, transactions: 6 });

  // The bank now serves less history, and the July export is no longer in the folder.
  w.bank.servedFrom = '2026-07-01';
  w.files.splice(1, 1);
  const first = w.reprocess();
  assert.deepStrictEqual(first.sync.results.map((r) => [r.connection, r.status, r.transactions.added]),
    [['revolut-niklas', 'ok', 0], ['csv-alex', 'ok', 0]], 'Trading 212 is not re-read');
  // CSV texts change, so those rows count as updated; Enable Banking rows only had the masked IBAN as their
  // account label, which every save stamps again from the refreshed accounts without counting an update.
  assert.ok(first.sync.results[1].transactions.updated > 0);
  assert.strictEqual(first.sync.results[0].transactions.unavailable, 1);
  assert.deepStrictEqual(w.counts(), before);
  assert.deepStrictEqual(sortedIds(w.sync().store.load()), idsBefore);

  const txns = w.read('transactions');
  const rent = txns.find((r) => r.counterparty === 'DEMO PERSON');
  assert.deepStrictEqual([rent.ibanMasked, rent.counterpartyAccount, rent.reference], ['LU120000000000001234', OTHER_IBAN, '202608280000005545']);
  assert.match(txns.find((r) => r.category === 'CREDIT').description, /BENEFICIAIRE: LU120000000000001439/);
  assert.ok(txns.filter((r) => r.provider === 'enablebanking').every((r) => r.ibanMasked === IBAN));
  const raw = w.read('raw_transactions');
  assert.strictEqual(raw.find((r) => r.counterparty === 'DEMO PERSON')['api.parsed.Beneficiary account'], OTHER_IBAN);
  const e1 = raw.find((r) => r.transactionId === 'E1');
  assert.strictEqual(e1['api.creditor_account.other.identification'], '12345678901234');
  assert.strictEqual(e1['api.debtor_account_additional_identification[0].identification'], '98765432109876');
  const accounts = w.read('raw_accounts');
  assert.deepStrictEqual(accounts.map((a) => a.ibanMasked).sort(), ['LU120000000000001234', IBAN]);
  assert.strictEqual(accounts.find((a) => a.provider === 'enablebanking')['api.account_id.iban'], IBAN);

  // What the sources no longer hold stays masked and is listed as not recoverable: the June bank row's API
  // fields, and the July CSV row (its account label is rewritten on every tab, its texts are not).
  const left = first.remaining.map((r) => [r.tab, r.connection, r.date, r.columns.join(','), r.recoverable]);
  assert.deepStrictEqual(left, [
    ['transactions', 'csv-alex', '2026-07-30', 'counterpartyAccount', false],
    ['raw_transactions', 'csv-alex', '2026-07-30', 'api.Message,api.parsed.Beneficiary account', false],
    ['raw_transactions', 'revolut-niklas', '2026-06-10',
      'api.creditor_account.other.identification,api.debtor_account_additional_identification[0].identification', false]
  ]);
  const text = BS.Reprocess.summarize(first);
  assert.match(text, /Still masked: 3 rows\. 3 cannot be recovered/);
  assert.match(text, /revolut-niklas: 0 added, \d+ updated, 0 pending removed, 1 date windows not served by the source/);
  assert.ok(!/824|9\.99|DEMO PERSON/.test(text), 'no amounts or text in the summary');

  const second = w.reprocess();
  assert.ok(second.sync.results.every((r) => r.transactions.added === 0 && r.transactions.updated === 0));
  assert.deepStrictEqual(w.counts(), before);
  assert.deepStrictEqual(sortedIds(w.sync().store.load()), idsBefore);
  assert.deepStrictEqual(second.remaining, first.remaining);
});

test('re-process: nothing left masked is said plainly; refused while masking is on', () => {
  const w = world();
  BS.SyncAll.run(w.sync());
  const out = w.reprocess();
  assert.deepStrictEqual(out.remaining, []);
  assert.match(BS.Reprocess.summarize(out), /Still masked: none/);
  withMasking(BS, true, () => assert.throws(() => w.reprocess(), { name: 'ConfigError' }));
});

test('re-process: rows of a connection cut short by the time budget are listed as recoverable', () => {
  const rows = [{ connection: 'csv-alex', accountName: 'A', transactionId: 't', bookingDate: '2026-08-01', description: 'x ****1234' },
    { connection: 'trading212-niklas', accountName: 'B', transactionId: 'u', bookingDate: '2026-08-02', description: '****9999' }];
  const deps = { sync: { store: { load: () => rows } }, rawReader: { readTable: () => [] } };
  const left = BS.Reprocess.remaining(deps, (key) => key !== 'csv-alex');
  assert.deepStrictEqual(left.map((r) => [r.connection, r.recoverable]), [['csv-alex', true], ['trading212-niklas', false]]);
  const text = BS.Reprocess.summarize({ sync: { results: [], raw: null, details: null, fx: null, moved: null }, remaining: left });
  assert.match(text, /1 cannot be recovered.*; 1 were not re-read in this run: run Re-process again/);
  assert.match(text, /\[run again\]/);
});
