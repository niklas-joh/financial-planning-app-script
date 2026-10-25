/**
 * @fileoverview CsvImportProvider: the `BankProvider` for one Drive folder of CSV exports. Each
 * run reads every `.csv` file in the folder, so a sync merges the whole folder; the idempotent
 * upsert makes files already imported change nothing. Rows that appear in two overlapping files
 * have the same id and count once. A file that matches no format, or cannot be read, is skipped
 * and named in the last-read record; the other files still import.
 */
(function (ns) {
  const CSV = ns.Config.CSV;

  /**
   * @typedef {Object} CsvFile
   * @property {string} name - File name.
   * @property {number} size - Size in bytes.
   * @property {function(string): string} text - Content decoded with the given charset.
   */

  /**
   * @typedef {Object} CsvProviderDeps
   * @property {import('../../core/config.js').ConnectionConfig} connection - CSV connection.
   * @property {{files: function(): Array<CsvFile>}} folder - The connection's Drive folder.
   * @property {{save: function(Object): void}} record - Last-read record (`CsvFolders.imports`).
   * @property {function(): number} nowMs - Clock.
   */

  /**
   * Adds one file's transactions to the per-source counts.
   * @param {Object<string, import('./csv-folders.js').SourceRead>} sources - Counts so far.
   * @param {string} id - Format id.
   * @param {Array<import('../bank-provider.js').Transaction>} rows - The file's new rows.
   */
  function count(sources, id, rows) {
    const s = sources[id] || (sources[id] = { files: 0, rows: 0, from: '', to: '' });
    s.files++;
    s.rows += rows.length;
    rows.forEach(function (t) {
      if (!s.from || t.bookingDate < s.from) s.from = t.bookingDate;
      if (t.bookingDate > s.to) s.to = t.bookingDate;
    });
  }

  /**
   * Reads and normalizes every CSV file in the folder.
   * @param {CsvProviderDeps} deps - Services.
   * @returns {{accounts: Array<import('../bank-provider.js').Account>,
   *   transactions: Array<import('../bank-provider.js').Transaction>, record: import('./csv-folders.js').ImportRecord}}
   *   Accounts and transactions (overlap removed) and what was read.
   */
  function readFolder(deps) {
    const accounts = {};
    const byKey = {};
    const record = { at: new Date(deps.nowMs()).toISOString(), sources: {}, skipped: [] };
    deps.folder.files()
      .filter(function (f) { return /\.csv$/i.test(f.name); })
      .sort(function (a, b) { return a.name < b.name ? -1 : a.name > b.name ? 1 : 0; })
      .forEach(function (f) {
        try {
          if (f.size > CSV.MAX_FILE_BYTES) throw ns.Errors.create('DataError', 'larger than ' + CSV.MAX_FILE_BYTES + ' bytes');
          const found = ns.CsvNormalize.detect(f.text);
          if (!found) throw ns.Errors.create('DataError', 'its header matches no known export');
          const out = ns.CsvNormalize.file(found, deps.connection);
          const fresh = out.transactions.filter(function (t) {
            const key = t.accountId + '|' + t.transactionId;
            if (byKey[key]) return false;
            byKey[key] = t;
            return true;
          });
          accounts[out.account.accountId] = out.account;
          count(record.sources, found.format.id, fresh);
        } catch (e) {
          record.skipped.push(ns.Redact.text(f.name + ' (' + e.message + ')'));
        }
      });
    return {
      accounts: Object.keys(accounts).sort().map(function (id) { return accounts[id]; }),
      transactions: Object.keys(byKey).map(function (key) { return byKey[key]; }),
      record: record
    };
  }

  /**
   * @param {CsvProviderDeps} deps - Services.
   * @returns {import('../bank-provider.js').BankProvider} Provider over the folder's files.
   */
  function create(deps) {
    let read = null;
    const load = function () {
      if (!read) {
        read = readFolder(deps);
        deps.record.save(read.record);
      }
      return read;
    };
    return ns.BankProvider.assertBankProvider({
      /** One read covers the whole folder, whatever date range is asked for. */
      maxWindowDays: Infinity,
      authorize: function () {
        throw ns.Errors.config('CSV files are read from a Drive folder: set ' + ns.Config.PROPERTY_KEYS.CSV_FOLDER_PREFIX +
          deps.connection.propertySuffix + ' to its id in Script Properties');
      },
      listAccounts: function () { return load().accounts; },
      getBalances: function () { return []; },
      /** Every row of the account in the folder, also older than `dateFrom` (the upsert keeps it idempotent). */
      getTransactions: function (accountId) {
        return load().transactions.filter(function (t) { return t.accountId === accountId; });
      }
    });
  }

  ns.CsvImportProvider = { create: create };
})(globalThis.BankSync);
