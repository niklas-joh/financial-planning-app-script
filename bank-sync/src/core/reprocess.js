/**
 * @fileoverview *Re-process stored rows*: once `Config.MASK_ACCOUNT_NUMBERS` is off, rewrites the rows that
 * were stored while it was on. Enable Banking connections read every account's details again and re-fetch
 * each account's stored history; CSV connections re-import every file still in their Drive folder. It is
 * Sync all with `fullHistory`, limited to those providers. Transaction and account ids are built from the
 * source data, never from masked text, so the rows are updated in place: nothing is added, and a second run
 * changes nothing. Rows still holding a masked value afterwards are listed, with whether they can be
 * recovered (their source no longer has them) or only need another run.
 */
(function (ns) {
  /**
   * @typedef {Object} MaskedRow
   * @property {string} tab - Tab the row is on.
   * @property {string} connection - Connection key, or '' for rows from before labels existed.
   * @property {string} accountName - Account label, or ''.
   * @property {string} date - Booking (else value) date, or '' for accounts.
   * @property {string} id - Transaction id (account id on the account and balance tabs).
   * @property {Array<string>} columns - Columns that still hold a masked value.
   * @property {boolean} recoverable - True if the connection was not fully re-read in this run (run again);
   *   false if it was, or cannot be re-read, so the source no longer has the full value.
   */

  /**
   * @typedef {Object} ReprocessDeps
   * @property {import('./sync-all.js').SyncAllDeps} sync - What Sync all needs (all connections; filtered here).
   * @property {{readTable: function(string, Array<Object>): Array<Object>}} rawReader - Reads a raw tab back.
   */

  /**
   * Providers whose stored rows can be read again from their source.
   * @returns {Array<string>} Provider ids.
   */
  function rereadable() {
    return [ns.Config.ENABLE_BANKING.PROVIDER_ID, ns.Config.CSV.PROVIDER_ID];
  }

  /**
   * @param {import('./config.js').ConnectionConfig} connection - Connection.
   * @returns {string} Its provider id (connections stored before the field existed are Enable Banking).
   */
  function providerOf(connection) {
    return connection.provider || ns.Config.ENABLE_BANKING.PROVIDER_ID;
  }

  /**
   * Rows that still hold a masked value.
   * @param {function(Object): string} tabOf - Tab name of a row.
   * @param {Array<Object>} rows - Stored rows.
   * @param {function(string): boolean} done - Whether a connection was fully re-read in this run.
   * @returns {Array<MaskedRow>} One entry per row with at least one masked column; no amounts or text.
   */
  function findMasked(tabOf, rows, done) {
    return rows.reduce(function (out, row) {
      const columns = Object.keys(row).filter(function (k) { return ns.Redact.isMasked(row[k]); });
      if (!columns.length) return out;
      const connection = String(row.connection || '');
      out.push({
        tab: tabOf(row),
        connection: connection,
        accountName: String(row.accountName || ''),
        date: row.transactionId ? ns.TxnMerge.effectiveDate(row).slice(0, 10) : '',
        id: String(row.transactionId || row.accountId || ''),
        columns: columns,
        recoverable: !done(connection)
      });
      return out;
    }, []);
  }

  /**
   * Lists the rows on the transaction tabs and the raw tabs that still hold a masked value.
   * @param {ReprocessDeps} deps - Services.
   * @param {function(string): boolean} done - Whether a connection was fully re-read in this run.
   * @returns {Array<MaskedRow>} Masked rows, transaction tabs first.
   */
  function remaining(deps, done) {
    const tabs = ns.Config.TABS;
    const raw = [
      [tabs.RAW_TRANSACTIONS, ns.RawWriter.TRANSACTION_COLUMNS],
      [tabs.RAW_ACCOUNTS, ns.RawWriter.ACCOUNT_COLUMNS],
      [tabs.RAW_BALANCES, ns.RawWriter.BALANCE_COLUMNS]
    ];
    return raw.reduce(function (out, pair) {
      const name = pair[0];
      return out.concat(findMasked(function () { return name; }, deps.rawReader.readTable(name, pair[1]), done));
    }, findMasked(ns.TxnRouting.tabFor, deps.sync.store.load(), done));
  }

  /**
   * Re-reads every Enable Banking and CSV connection in full, then lists what is still masked.
   * @param {ReprocessDeps} deps - Services.
   * @returns {{sync: Object, remaining: Array<MaskedRow>}} The Sync all outcome and the rows still masked.
   * @throws {Error} ConfigError while `MASK_ACCOUNT_NUMBERS` is on (the run would only mask again).
   */
  function run(deps) {
    if (ns.Config.MASK_ACCOUNT_NUMBERS) {
      throw ns.Errors.config('MASK_ACCOUNT_NUMBERS is on in core/config.js, so re-processing would mask the rows again. ' +
        'Turn it off and push first.');
    }
    const ids = rereadable();
    const connections = deps.sync.connections.filter(function (c) { return ids.indexOf(providerOf(c)) !== -1; });
    const sync = ns.SyncAll.run(Object.assign({}, deps.sync, { connections: connections, fullHistory: true }));
    const finished = {};
    sync.results.forEach(function (r) {
      if (r.status === 'ok' && r.transactions.complete) finished[r.connection] = true;
    });
    const reread = {};
    connections.forEach(function (c) { reread[c.key] = true; });
    // A connection not re-read at all (another provider, or no label) cannot be recovered by running again.
    const done = function (key) { return Boolean(finished[key]) || !reread[key]; };
    return { sync: sync, remaining: remaining(deps, done) };
  }

  /**
   * Turns a run into the text shown to the user: the Sync all summary, then the rows still masked (tab,
   * connection, account, date, id and column names only; never amounts or text).
   * @param {{sync: Object, remaining: Array<MaskedRow>}} outcome - From `run`.
   * @returns {string} Multi-line summary.
   */
  function summarize(outcome) {
    const lines = [ns.SyncAll.summarize(outcome.sync), ''];
    const rows = outcome.remaining;
    if (!rows.length) {
      lines.push('Still masked: none. Every stored row holds full account numbers.');
      return lines.join('\n');
    }
    const lost = rows.filter(function (r) { return !r.recoverable; });
    const again = rows.length - lost.length;
    lines.push('Still masked: ' + rows.length + ' rows. ' +
      (lost.length ? lost.length + ' cannot be recovered (their source no longer has them)' : '') +
      (lost.length && again ? '; ' : '') +
      (again ? again + ' were not re-read in this run: run Re-process again' : '') + '.');
    const max = ns.Config.REPROCESS.MAX_LISTED;
    lost.concat(rows.filter(function (r) { return r.recoverable; })).slice(0, max).forEach(function (r) {
      lines.push('- ' + [r.tab, r.connection || '(no connection)', r.accountName, r.date, r.id].filter(Boolean).join(' | ') +
        ' (' + r.columns.join(', ') + ')' + (r.recoverable ? ' [run again]' : ''));
    });
    if (rows.length > max) lines.push('... and ' + (rows.length - max) + ' more.');
    return lines.join('\n');
  }

  ns.Reprocess = { run: run, summarize: summarize, remaining: remaining };
})(globalThis.BankSync);
