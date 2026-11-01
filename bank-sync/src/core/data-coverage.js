/**
 * @fileoverview Where each account's data comes from and how far it reaches, so a missing month is seen and the
 * next manual export starts at the right date. Two parts:
 * - How an account is fetched (`source`, `frequency`, `howTo`) is typed by the owner on `account_owners`; a new
 *   account starts with the defaults below (`defaults`).
 * - What is covered is computed by every Sync all and written to the `data_coverage` tab (never typed): per
 *   account its rows, first and last date, the date its data is complete up to (`upTo`), for a CSV account the
 *   date to export from, and a status against its frequency. Items on the hand-kept `manual_balances` tab are
 *   listed too, with the date of their latest value.
 * `upTo` is: for an API account the last day its connection was read without an error (kept in
 * `BS_ACCOUNTS_READ`); for a CSV account the end of its newest export (`CsvImportProvider`: the latest date in
 * the files of that source uploaded the same day, since a quiet account's own last row can be months older than
 * the export); for a manual item the date of its latest value. Dates, counts and names only.
 */
(function (ns) {
  const COV = ns.Config.COVERAGE;

  /** Source labels a new account starts with (the owner may type anything instead). */
  const SOURCES = Object.freeze({
    API: 'API (automatic)',
    CSV_MANUAL: 'CSV (manual upload)',
    MANUAL: 'Manual entry'
  });

  /** Columns of the `data_coverage` tab. */
  const HEADER = ['owner', 'bank', 'accountName', 'source', 'frequency', 'rows', 'firstDate', 'lastDate', 'upTo',
    'exportFrom', 'status', 'nextStep', 'accountId'];

  /** How a new account of each provider is fetched; CSV accounts are set per source in `defaults`. */
  const HOW_TO = Object.freeze({
    enablebanking: 'Daily sync; log in again when the consent warning email comes (every 90-180 days)',
    trading212: 'Daily sync; a monthly report CSV is saved to its exports folder automatically (card details)',
    pluxee: 'Daily sync; run "Connect Pluxee" again when Show status says the login lapsed',
    csv: 'Export the account movements as CSV from "exportFrom" to today and upload the file to the CSV folder'
  });

  /**
   * @param {{provider: string, accountId: string}} account - Account.
   * @returns {{source: string, frequency: string, howTo: string}} What a new `account_owners` row starts with.
   */
  function defaults(account) {
    if (account.provider === ns.Config.CSV.PROVIDER_ID) {
      const sourceId = String(account.accountId).split('-')[0];
      return { source: SOURCES.CSV_MANUAL, frequency: COV.CSV_FREQUENCY[sourceId] || COV.DEFAULT_CSV_FREQUENCY, howTo: HOW_TO.csv };
    }
    return { source: SOURCES.API, frequency: 'daily', howTo: HOW_TO[account.provider] || '' };
  }

  /**
   * @param {string} frequency - As typed, e.g. 'Monthly'.
   * @returns {?number} Days between updates, or null for no schedule ('ad hoc', blank or unknown).
   */
  function frequencyDays(frequency) {
    const days = COV.FREQUENCY_DAYS[String(frequency || '').trim().toLowerCase()];
    return days || null;
  }

  /**
   * @param {string} iso - YYYY-MM-DD.
   * @returns {number} Days since 1970-01-01.
   */
  function dayNumber(iso) {
    return Math.round(Date.parse(iso + 'T00:00:00Z') / 86400000);
  }

  /**
   * Status of an account's data. OK while it is at most one period (plus a day) old, due up to two periods,
   * overdue after that.
   * @param {string} frequency - As typed.
   * @param {string} upTo - Date the data is complete up to, or ''.
   * @param {string} today - YYYY-MM-DD.
   * @returns {string} 'OK', 'due', 'overdue', 'no data yet' or 'no schedule'.
   */
  function statusOf(frequency, upTo, today) {
    const days = frequencyDays(frequency);
    if (!days) return 'no schedule';
    if (!upTo) return 'no data yet';
    const age = dayNumber(today) - dayNumber(upTo);
    if (age <= days + 1) return 'OK';
    return age <= 2 * days + 1 ? 'due' : 'overdue';
  }

  /**
   * @param {Object} row - Coverage row.
   * @returns {boolean} Whether it needs doing.
   */
  function needsAction(row) {
    return row.status === 'due' || row.status === 'overdue' || row.status === 'no data yet';
  }

  /**
   * @param {Object} row - Coverage row (status set).
   * @param {string} kind - 'csv', 'api' or 'manual'.
   * @returns {string} What to do now, or ''.
   */
  function nextStep(row, kind) {
    if (!needsAction(row)) return '';
    if (kind === 'csv') {
      return (row.exportFrom ? 'Export ' + row.exportFrom + ' to today' : 'Export the full history') + ' and upload it to the CSV folder';
    }
    if (kind === 'manual') return 'Add a dated row with the current value on ' + ns.Config.TABS.MANUAL_BALANCES;
    return (row.upTo ? 'Not read since ' + row.upTo : 'Never read') + ': check Show status (consent or login)';
  }

  /**
   * @param {Object} row - Row read from a tab.
   * @param {string} key - Column.
   * @returns {string} The cell as trimmed text, '' when missing.
   */
  function cell(row, key) {
    return String(row[key] === undefined || row[key] === null ? '' : row[key]).trim();
  }

  /**
   * First date, last date and row count per account.
   * @param {Array<{accountId: string, bookingDate: string}>} transactions - Rows of the three transaction tabs.
   * @returns {Object<string, {rows: number, first: string, last: string}>} Per account id.
   */
  function spans(transactions) {
    const out = {};
    transactions.forEach(function (t) {
      const date = cell(t, 'bookingDate');
      const s = out[t.accountId] || (out[t.accountId] = { rows: 0, first: '', last: '' });
      s.rows++;
      if (!date) return;
      if (!s.first || date < s.first) s.first = date;
      if (date > s.last) s.last = date;
    });
    return out;
  }

  /**
   * Coverage rows of the accounts on `account_owners`.
   * @param {Array<Object>} owners - Rows of `account_owners`.
   * @param {Object<string, {rows: number, first: string, last: string}>} byAccount - From `spans`.
   * @param {{read: Object<string,string>, csv: Object<string,string>, today: string}} dates - Last good read per
   *   API account, export end per CSV account, today.
   * @returns {Array<Object>} Rows.
   */
  function accountRows(owners, byAccount, dates) {
    return owners.filter(function (o) { return cell(o, 'accountId'); }).map(function (o) {
      const id = cell(o, 'accountId');
      const isCsv = cell(o, 'provider') === ns.Config.CSV.PROVIDER_ID;
      const fallback = defaults({ provider: cell(o, 'provider'), accountId: id });
      const span = byAccount[id] || { rows: 0, first: '', last: '' };
      const upTo = (isCsv ? dates.csv[id] : dates.read[id]) || '';
      const row = {
        owner: ns.Connections.slug(cell(o, 'owner')), bank: cell(o, 'bank'), accountName: cell(o, 'accountName'),
        source: cell(o, 'source') || fallback.source, frequency: cell(o, 'frequency') || fallback.frequency,
        rows: span.rows, firstDate: span.first, lastDate: span.last,
        upTo: isCsv && span.last > upTo ? span.last : upTo,
        exportFrom: '', accountId: id
      };
      if (isCsv) row.exportFrom = row.upTo;
      row.status = statusOf(row.frequency, row.upTo, dates.today);
      row.nextStep = nextStep(row, isCsv ? 'csv' : 'api');
      return row;
    });
  }

  /**
   * Coverage rows of the items on `manual_balances` (one per bank and name; the latest row with an amount counts).
   * @param {Array<Object>} manual - Rows of `manual_balances` (dates as YYYY-MM-DD).
   * @param {string} today - YYYY-MM-DD.
   * @returns {Array<Object>} Rows.
   */
  function manualRows(manual, today) {
    const items = {};
    const order = [];
    manual.forEach(function (m) {
      const name = cell(m, 'accountName');
      if (!name) return;
      const key = cell(m, 'bank') + '|' + name;
      let item = items[key];
      if (!item) {
        item = items[key] = { owner: '', bank: cell(m, 'bank'), accountName: name, frequency: '', rows: 0, first: '', last: '', latest: '' };
        order.push(key);
      }
      const date = cell(m, 'date');
      if (date >= item.latest) {
        item.latest = date;
        item.owner = ns.Connections.slug(cell(m, 'owner')) || item.owner;
        item.frequency = cell(m, 'frequency') || item.frequency;
      }
      if (!date || cell(m, 'amount') === '') return;
      item.rows++;
      if (!item.first || date < item.first) item.first = date;
      if (date > item.last) item.last = date;
    });
    return order.map(function (key) {
      const item = items[key];
      const row = {
        owner: item.owner, bank: item.bank, accountName: item.accountName, source: SOURCES.MANUAL,
        frequency: item.frequency || COV.DEFAULT_MANUAL_FREQUENCY, rows: item.rows, firstDate: item.first,
        lastDate: item.last, upTo: item.last, exportFrom: '', accountId: ''
      };
      row.status = statusOf(row.frequency, row.upTo, today);
      row.nextStep = nextStep(row, 'manual');
      return row;
    });
  }

  /**
   * Every coverage row: the accounts, then the manual items.
   * @param {{owners: Array<Object>, transactions: Array<Object>, manual: Array<Object>, read: Object<string,string>,
   *   csv: Object<string,string>, today: string}} input - What the rows are built from.
   * @returns {Array<Object>} Rows (`HEADER` keys).
   */
  function build(input) {
    return accountRows(input.owners, spans(input.transactions), input)
      .concat(manualRows(input.manual, input.today));
  }

  /**
   * @param {Object} row - Coverage row.
   * @returns {string} e.g. "BIL Current Private (export 2026-08-31 to today)".
   */
  function dueText(row) {
    const name = [row.bank, row.accountName].filter(Boolean).join(' ');
    if (row.exportFrom) return name + ' (export ' + row.exportFrom + ' to today)';
    return name + (row.upTo ? ' (up to ' + row.upTo + ')' : ' (no data yet)');
  }

  /**
   * @param {Array<Object>} rows - Coverage rows.
   * @returns {string} The ones needing action, first `MAX_LISTED`, then "and N more"; '' if none.
   */
  function dueList(rows) {
    const due = rows.filter(needsAction).map(dueText);
    const more = due.length - COV.MAX_LISTED;
    return due.slice(0, COV.MAX_LISTED).join(', ') + (more > 0 ? ' and ' + more + ' more' : '');
  }

  /**
   * Last good read per API account, in one Script Property.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): *}} props - Script Properties.
   * @returns {{load: function(): Object<string,string>, save: function(Object<string,string>): void}} Store.
   */
  function readStore(props) {
    const key = ns.Config.PROPERTY_KEYS.ACCOUNTS_READ;
    return {
      load: function () {
        try {
          return JSON.parse(props.getProperty(key) || '{}') || {};
        } catch (e) {
          return {};
        }
      },
      save: function (map) { props.setProperty(key, JSON.stringify(map)); }
    };
  }

  /**
   * @typedef {Object} CoverageOutcome
   * @property {number} rows - Rows written to `data_coverage`.
   * @property {string} due - Accounts needing action (`dueList`), or ''.
   * @property {string} [error] - Redacted message when the tab could not be built or written.
   */

  /**
   * Coverage for one run.
   * @param {{tab: {write: function(Array<Object>): void, loadManual: function(): Array<Object>},
   *   owners: {load: function(): Array<Object>}, store: {load: function(): Array<Object>},
   *   read: {load: function(): Object<string,string>, save: function(Object<string,string>): void},
   *   csvCoverage: function(): Object<string,string>, todayIso: function(): string}} deps - Services.
   * @returns {{update: function(Array<{accountId: string, provider: string}>): CoverageOutcome}} `update` records
   *   the accounts whose transactions were read to today in this run (API accounts; Sync all leaves out a connection
   *   that failed or ran out of time) and rewrites the tab. It never throws.
   */
  function create(deps) {
    return {
      update: function (readAccounts) {
        try {
          const today = deps.todayIso();
          const read = deps.read.load();
          const api = readAccounts.filter(function (a) { return a.provider !== ns.Config.CSV.PROVIDER_ID; });
          api.forEach(function (a) { read[a.accountId] = today; });
          if (api.length) deps.read.save(read);
          const rows = build({
            owners: deps.owners.load(), transactions: deps.store.load(), manual: deps.tab.loadManual(),
            read: read, csv: deps.csvCoverage(), today: today
          });
          deps.tab.write(rows);
          return { rows: rows.length, due: dueList(rows) };
        } catch (e) {
          return { rows: 0, due: '', error: ns.Redact.text(e.name + ': ' + e.message) };
        }
      }
    };
  }

  /**
   * One summary line for a run, or '' without an outcome.
   * @param {?CoverageOutcome} outcome - From `update`.
   * @returns {string} Line (names and dates only).
   */
  function summaryLine(outcome) {
    if (!outcome) return '';
    const tabName = ns.Config.TABS.DATA_COVERAGE;
    if (outcome.error) return 'Data coverage: ' + tabName + ' could not be updated (' + outcome.error + ').';
    return 'Data coverage: ' + outcome.rows + ' rows on ' + tabName + '; ' + (outcome.due ? 'to update: ' + outcome.due : 'all up to date') + '.';
  }

  /**
   * Show status line, from the tab as the last sync wrote it.
   * @param {{load: function(): Array<Object>}} tab - The `data_coverage` tab.
   * @returns {string} Line.
   */
  function statusLine(tab) {
    const tabName = ns.Config.TABS.DATA_COVERAGE;
    let rows;
    try {
      rows = tab.load();
    } catch (e) {
      return 'Data coverage: ' + tabName + ' cannot be read (' + ns.Redact.text(e.name + ': ' + e.message) + ')';
    }
    if (!rows.length) return 'Data coverage: none yet (the next sync fills ' + tabName + ')';
    const due = dueList(rows.map(function (r) {
      return { bank: cell(r, 'bank'), accountName: cell(r, 'accountName'), exportFrom: cell(r, 'exportFrom'),
        upTo: cell(r, 'upTo'), status: cell(r, 'status') };
    }));
    return 'Data coverage: ' + rows.length + ' rows on ' + tabName + '; ' + (due ? 'to update: ' + due : 'all up to date');
  }

  ns.DataCoverage = {
    SOURCES: SOURCES, HEADER: HEADER, defaults: defaults, frequencyDays: frequencyDays, statusOf: statusOf,
    build: build, readStore: readStore, create: create, summaryLine: summaryLine, statusLine: statusLine
  };
})(globalThis.BankSync);
