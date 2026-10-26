// B1: which tab a row is stored on, and the routed store over the three transaction tabs. Fake data only.
const test = require('node:test');
const assert = require('node:assert');
const BS = require('./load');
const { fakeSheetBook } = require('./fake-sheet');

const TABS = BS.Config.TABS;
const iso = (d) => d.toISOString().slice(0, 10);
const tx = (o) => ({
  provider: 'enablebanking', bank: '', connection: '', holderName: '', accountName: '', ibanMasked: '', accountId: 'a1',
  transactionId: 't1', bookingDate: '2026-09-10', valueDate: '2026-09-10', status: 'booked', amount: -5, currency: 'EUR',
  amountEur: null, fxRate: null, fxRateDate: '', counterparty: '', description: '', category: '', counterpartyAccount: '',
  counterpartyBank: '', reference: '', kind: '', ...o
});
const t212 = (o) => tx({ provider: 'trading212', accountId: '1234567', ...o });

/** A routed store over a fake book; `saves` counts the writes per tab and `fail` makes a tab's write throw. */
function routed(book) {
  const saves = {};
  const fail = {};
  const stores = {};
  BS.TxnRouting.tabs().forEach((tab) => {
    const inner = BS.TransactionsStore.create(book, iso, tab);
    saves[tab] = 0;
    stores[tab] = {
      load: inner.load,
      save: (rows) => { if (fail[tab]) throw new Error('write failed'); saves[tab]++; return inner.save(rows); }
    };
  });
  return { store: BS.TxnRouting.createStore(stores), saves, fail, tab: (name) => BS.TransactionsStore.create(book, iso, name).load() };
}

const ids = (rows) => rows.map((t) => t.transactionId).sort();

test('routing: banks and card spending to transactions, Trading 212 to investments, its unclear types to unclassified', () => {
  const tabFor = BS.TxnRouting.tabFor;
  assert.strictEqual(tabFor(tx({})), TABS.TRANSACTIONS);
  assert.strictEqual(tabFor(tx({ provider: 'csv' })), TABS.TRANSACTIONS);
  assert.strictEqual(tabFor(t212({ description: 'WITHDRAW' })), TABS.UNCLASSIFIED_TRANSACTIONS);
  assert.strictEqual(tabFor(t212({ description: 'DEPOSIT' })), TABS.UNCLASSIFIED_TRANSACTIONS);
  ['INTEREST_ON_FREE_CASH', 'LENDING_INTEREST', 'FEE', 'TRANSFER', 'Dividend ORDINARY', 'NEW_TYPE'].forEach((type) =>
    assert.strictEqual(tabFor(t212({ description: type })), TABS.INVESTMENT_TRANSACTIONS, type));
  // A report label wins over the type.
  assert.strictEqual(tabFor(t212({ description: 'WITHDRAW', kind: 'card' })), TABS.TRANSACTIONS);
  assert.strictEqual(tabFor(t212({ description: 'DEPOSIT', kind: 'investment' })), TABS.INVESTMENT_TRANSACTIONS);
  // A provider name that is also an Object member is not an investment provider.
  assert.strictEqual(tabFor(tx({ provider: 'constructor' })), TABS.TRANSACTIONS);
  assert.ok(Object.prototype.hasOwnProperty.call(BS.Config.ROUTING.INVESTMENT_PROVIDERS, BS.Config.TRADING212.PROVIDER_ID));
});

test('report kind: card for a merchant, a card action or cashback; investment for everything else', () => {
  const isCard = BS.CsvT212Report.isCard;
  assert.ok(isCard('DEMO CAFE', 'CARD_DEBIT'));
  assert.ok(isCard('', 'CARD_DEBIT'));
  assert.ok(isCard('', 'CARD_REFUND'));
  assert.ok(isCard('', 'SPENDING_CASHBACK'));
  assert.ok(isCard('RESTAURANTS', 'SOMETHING_NEW'));
  assert.ok(!isCard('', 'DEPOSIT'));
  assert.ok(!isCard('', 'WITHDRAWAL'));
  assert.ok(!isCard('', 'INTEREST_ON_CASH'));
});

test('migration: the first save moves the Trading 212 rows out of transactions, keeping EUR columns and details; a second moves nothing', () => {
  const book = fakeSheetBook();
  // The tab as it is before B1: every row on transactions.
  BS.TransactionsStore.create(book, iso).save([
    tx({ transactionId: 'bank-1' }),
    t212({ transactionId: 'int-1', description: 'INTEREST_ON_FREE_CASH', amount: 0.04, amountEur: 0.04, fxRate: 1, fxRateDate: '2026-09-10' }),
    t212({ transactionId: 'wd-1', description: 'WITHDRAW', amount: -20, currency: 'GBP', amountEur: -23.1, fxRate: 0.865, fxRateDate: '2026-09-10' }),
    t212({ transactionId: 'card-1', description: 'WITHDRAW', counterparty: 'DEMO CAFE', category: 'RESTAURANTS', kind: 'card' })
  ]);
  const first = routed(book);
  const rows = first.store.load();
  assert.strictEqual(rows.length, 4);
  first.store.save(rows);
  assert.deepStrictEqual(ids(first.tab(TABS.TRANSACTIONS)), ['bank-1', 'card-1']);
  assert.deepStrictEqual(ids(first.tab(TABS.INVESTMENT_TRANSACTIONS)), ['int-1']);
  assert.deepStrictEqual(ids(first.tab(TABS.UNCLASSIFIED_TRANSACTIONS)), ['wd-1']);
  const moved = first.tab(TABS.UNCLASSIFIED_TRANSACTIONS)[0];
  assert.deepStrictEqual([moved.amountEur, moved.fxRate, moved.fxRateDate, moved.currency], [-23.1, 0.865, '2026-09-10', 'GBP']);
  assert.strictEqual(first.tab(TABS.TRANSACTIONS).find((t) => t.transactionId === 'card-1').counterparty, 'DEMO CAFE');
  assert.deepStrictEqual(first.store.moved(), { [TABS.TRANSACTIONS]: 0, [TABS.INVESTMENT_TRANSACTIONS]: 1, [TABS.UNCLASSIFIED_TRANSACTIONS]: 1 });
  assert.strictEqual(BS.TxnRouting.movedText(first.store.moved()), '1 to investment_transactions, 1 to unclassified_transactions');

  const second = routed(book);
  second.store.save(second.store.load());
  assert.deepStrictEqual(second.saves, { [TABS.TRANSACTIONS]: 0, [TABS.INVESTMENT_TRANSACTIONS]: 0, [TABS.UNCLASSIFIED_TRANSACTIONS]: 0 });
  assert.strictEqual(BS.TxnRouting.movedText(second.store.moved()), '');
});

test('routed store: a tab that did not change is not written, and tabs without rows are not created', () => {
  const book = fakeSheetBook();
  const r = routed(book);
  r.store.save([tx({ transactionId: 'bank-1' })]);
  assert.deepStrictEqual(Object.keys(book.sheets), [TABS.TRANSACTIONS]);
  r.store.save([tx({ transactionId: 'bank-1' }), t212({ transactionId: 'int-1', description: 'INTEREST_ON_FREE_CASH' })]);
  assert.deepStrictEqual(r.saves, { [TABS.TRANSACTIONS]: 1, [TABS.INVESTMENT_TRANSACTIONS]: 1, [TABS.UNCLASSIFIED_TRANSACTIONS]: 0 });
});

test('a report labels an unclassified row: it moves once; a later API re-read updates it where it is and never adds it again', () => {
  const book = fakeSheetBook();
  const r = routed(book);
  const api = t212({ transactionId: 'ref-1', description: 'WITHDRAW', amount: -9.5 });
  r.store.save([api]);
  assert.deepStrictEqual(ids(r.tab(TABS.UNCLASSIFIED_TRANSACTIONS)), ['ref-1']);

  const labelled = BS.TxnMerge.enrich(r.store.load(), [{ provider: 'trading212', transactionId: 'ref-1', counterparty: 'DEMO SHOP', category: 'RETAIL', kind: 'card' }]);
  assert.strictEqual(labelled.enriched, 1);
  r.store.save(labelled.rows);
  assert.deepStrictEqual(ids(r.tab(TABS.TRANSACTIONS)), ['ref-1']);
  assert.deepStrictEqual(r.tab(TABS.UNCLASSIFIED_TRANSACTIONS), []);

  // The API's 7-day re-read brings the same row without the report's details.
  const merged = BS.TxnMerge.merge(r.store.load(), [api], [{ accountId: '1234567', from: '2026-09-01', to: '2026-09-30' }]);
  assert.deepStrictEqual([merged.added, merged.updated], [0, 0]);
  r.store.save(merged.rows);
  const all = r.store.load();
  assert.strictEqual(all.length, 1);
  assert.deepStrictEqual([all[0].kind, all[0].counterparty], ['card', 'DEMO SHOP']);
  assert.deepStrictEqual(ids(r.tab(TABS.TRANSACTIONS)), ['ref-1']);
  assert.strictEqual(r.store.moved()[TABS.TRANSACTIONS], 1);
});

test('a save that stops halfway leaves a moving row in both tabs, never in none; the next load joins it', () => {
  const book = fakeSheetBook();
  const setup = routed(book);
  setup.store.save([t212({ transactionId: 'ref-1', description: 'DEPOSIT', amount: 0.3 })]);

  // The new tab is written first: if that write fails, the row is still where it was.
  const early = routed(book);
  early.fail[TABS.TRANSACTIONS] = true;
  assert.throws(() => early.store.save(early.store.load().map((t) => ({ ...t, kind: 'card' }))), /write failed/);
  assert.deepStrictEqual(ids(early.tab(TABS.UNCLASSIFIED_TRANSACTIONS)), ['ref-1']);

  // If the old tab's write fails, the row is in both tabs; the next load returns it once, with the label.
  const late = routed(book);
  late.fail[TABS.UNCLASSIFIED_TRANSACTIONS] = true;
  assert.throws(() => late.store.save(late.store.load().map((t) => ({ ...t, kind: 'card', category: 'SPENDING_CASHBACK' }))), /write failed/);
  assert.deepStrictEqual(ids(late.tab(TABS.TRANSACTIONS)), ['ref-1']);
  assert.deepStrictEqual(ids(late.tab(TABS.UNCLASSIFIED_TRANSACTIONS)), ['ref-1']);
  const next = routed(book);
  const rows = next.store.load();
  assert.strictEqual(rows.length, 1);
  assert.deepStrictEqual([rows[0].kind, rows[0].category], ['card', 'SPENDING_CASHBACK']);
  next.store.save(rows);
  assert.deepStrictEqual(next.tab(TABS.UNCLASSIFIED_TRANSACTIONS), []);
  assert.deepStrictEqual(ids(next.tab(TABS.TRANSACTIONS)), ['ref-1']);
});

test('routed store: save before any load still knows what the tabs hold', () => {
  const book = fakeSheetBook();
  BS.TransactionsStore.create(book, iso).save([t212({ transactionId: 'int-1', description: 'INTEREST_ON_FREE_CASH' })]);
  const r = routed(book);
  r.store.save([t212({ transactionId: 'int-1', description: 'INTEREST_ON_FREE_CASH' })]);
  assert.deepStrictEqual(r.tab(TABS.TRANSACTIONS), []);
  assert.deepStrictEqual(ids(r.tab(TABS.INVESTMENT_TRANSACTIONS)), ['int-1']);
});

test('status line: rows per tab, 0 for a tab not created yet', () => {
  const book = fakeSheetBook();
  routed(book).store.save([tx({ transactionId: 'b1' }), tx({ transactionId: 'b2' }), t212({ transactionId: 'w1', description: 'WITHDRAW' })]);
  assert.strictEqual(BS.TxnRouting.statusLine(book), 'Rows: transactions 2, investment_transactions 0, unclassified_transactions 1');
});

test('sync all: one EUR conversion over every tab; the summary and the last-run record name the rows moved', () => {
  const book = fakeSheetBook();
  const r = routed(book);
  const converted = [];
  const deps = {
    // Built by the real module from a property name, as in production.
    connections: BS.T212Credentials.connections({ getKeys: () => ['T212_API_KEY_ALEX'], getProperty: () => 'x' }),
    isConnected: () => true,
    providerFor: () => ({
      maxWindowDays: Infinity,
      listAccounts: () => [{ provider: 'trading212', accountId: '1234567', bank: 'Trading 212', currency: 'EUR' }],
      getBalances: () => [],
      getTransactions: () => [
        t212({ transactionId: 'int-1', description: 'INTEREST_ON_FREE_CASH', amount: 0.03 }),
        t212({ transactionId: 'wd-1', description: 'WITHDRAW', amount: -12 })
      ]
    }),
    store: r.store,
    rawWriter: { writeTable: (tab, cols, rows) => rows.length },
    nowMs: () => Date.parse('2026-09-15T06:00:00Z'),
    todayIso: () => '2026-09-15',
    fx: {
      convertBalances: (rows) => ({ rows, missing: 0, failed: [] }),
      convertTransactions: (rows) => {
        converted.push(ids(rows));
        return { rows: rows.map((t) => ({ ...t, amountEur: t.amount, fxRate: 1, fxRateDate: t.bookingDate })), changed: rows.length, missing: 0, waiting: 0, failed: [] };
      }
    }
  };
  const outcome = BS.SyncAll.run(deps);
  assert.deepStrictEqual(converted, [['int-1', 'wd-1']]);
  assert.strictEqual(r.tab(TABS.INVESTMENT_TRANSACTIONS)[0].amountEur, 0.03);
  assert.strictEqual(r.tab(TABS.UNCLASSIFIED_TRANSACTIONS)[0].amountEur, -12);
  assert.deepStrictEqual(r.tab(TABS.TRANSACTIONS), []);
  // New rows were never on another tab, so nothing counts as moved.
  assert.doesNotMatch(BS.SyncAll.summarize(outcome), /Moved between tabs/);

  const moving = { ...outcome, moved: { [TABS.TRANSACTIONS]: 2, [TABS.INVESTMENT_TRANSACTIONS]: 0, [TABS.UNCLASSIFIED_TRANSACTIONS]: 0 } };
  assert.match(BS.SyncAll.summarize(moving), /Moved between tabs: 2 to transactions\./);
  const record = BS.LastRun.fromOutcome(moving, { at: '2026-09-15 06:20', source: 'daily', warningsSent: 0 });
  assert.ok(BS.LastRun.describe(record).includes('  Moved between tabs: 2 to transactions'));
});
