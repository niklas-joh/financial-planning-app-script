/**
 * @fileoverview Syncs every configured connection in turn: raw accounts and
 * balances, then transactions. One connection failing (e.g. a lapsed consent)
 * never stops the others, and the time budget is shared.
 */
(function (ns) {
  /**
   * @typedef {Object} SyncAllDeps
   * @property {Array<import('./config.js').ConnectionConfig>} connections - What to sync.
   * @property {function(import('./config.js').ConnectionConfig): boolean} isConnected - Whether a session exists.
   * @property {function(import('./config.js').ConnectionConfig): import('../providers/bank-provider.js').BankProvider} providerFor
   * @property {{load: function(): Array<Object>, save: function(Array<Object>): number}} store - Transactions store.
   * @property {{writeTable: function(string, Array<Object>, Array<Object>, Array<string>=): number}} rawWriter - Raw tab writer.
   * @property {function(): number} nowMs - Clock.
   * @property {function(): string} todayIso - Today's date in the configured time zone.
   */

  /**
   * @typedef {Object} ConnectionResult
   * @property {string} connection - Connection key.
   * @property {string} status - 'ok', 'not connected', 'skipped: time budget' or 'error'.
   * @property {string} [error] - Redacted message when status is 'error'.
   * @property {import('./txn-sync.js').TxnSyncResult} [transactions] - Counts when status is 'ok'.
   */

  /**
   * Syncs one connection.
   * @param {import('./config.js').ConnectionConfig} connection - Connection.
   * @param {SyncAllDeps} deps - Services.
   * @param {number} budgetMs - Time left for it.
   * @param {number} startedAt - When the whole run started (epoch ms).
   * @returns {{data: import('./raw-sync.js').RawData, transactions: Object}} What was read and the counts.
   */
  function syncOne(connection, deps, budgetMs, startedAt) {
    const provider = deps.providerFor(connection);
    const accounts = provider.listAccounts();
    const data = ns.RawSync.collect(provider, accounts, function () { return deps.nowMs() - startedAt > ns.Config.TRANSACTIONS.TIME_BUDGET_MS; });
    const transactions = ns.TxnSync.run({
      provider: provider, accounts: accounts, store: deps.store,
      nowMs: deps.nowMs, todayIso: deps.todayIso, budgetMs: budgetMs
    });
    return { data: data, transactions: transactions };
  }

  /**
   * Runs the sync for all connections.
   * @param {SyncAllDeps} deps - Injected services.
   * @returns {{results: Array<ConnectionResult>, raw: ?{accounts: number, balances: number, transactions: number}}} Per-connection
   *   outcome and the raw rows written (null if no connection could be read). No transaction content.
   */
  function run(deps) {
    const cfg = ns.Config.TRANSACTIONS;
    const startedAt = deps.nowMs();
    const collected = { accounts: [], balances: [], transactions: [] };
    const refreshed = [];

    const results = deps.connections.map(function (connection) {
      const result = { connection: connection.key, status: 'ok' };
      const left = cfg.TIME_BUDGET_MS - (deps.nowMs() - startedAt);
      if (!deps.isConnected(connection)) {
        result.status = 'not connected';
      } else if (left < cfg.MIN_CONNECTION_BUDGET_MS) {
        result.status = 'skipped: time budget';
      } else {
        try {
          const done = syncOne(connection, deps, left, startedAt);
          collected.accounts = collected.accounts.concat(done.data.accounts);
          collected.balances = collected.balances.concat(done.data.balances);
          collected.transactions = collected.transactions.concat(done.transactions.rawRows);
          result.transactions = Object.assign({}, done.transactions);
          delete result.transactions.rawRows;
          refreshed.push(connection.key);
        } catch (e) {
          if (e.name === 'TimeBudget') {
            result.status = 'skipped: time budget';
          } else {
            result.status = 'error';
            result.error = ns.Redact.text(e.name + ': ' + e.message);
          }
        }
      }
      return result;
    });

    return { results: results, raw: refreshed.length ? ns.RawSync.write(deps.rawWriter, collected, refreshed) : null };
  }

  /**
   * Turns a run into the text shown to the user. Counts only, no transaction content.
   * @param {{results: Array<ConnectionResult>, raw: ?{accounts: number, balances: number, transactions: number}}} outcome - From `run`.
   * @returns {string} Multi-line summary.
   */
  function summarize(outcome) {
    const lines = outcome.results.map(function (r) {
      if (r.status === 'ok') {
        const t = r.transactions;
        return r.connection + ': ' + t.added + ' added, ' + t.updated + ' updated, ' + t.removed + ' pending removed' +
          (t.complete ? '' : ' (time budget reached; run again to continue)');
      }
      return r.connection + ': ' + r.status + (r.error ? ' (' + r.error + ')' : '');
    });
    if (outcome.raw) lines.push('Raw tabs: ' + outcome.raw.accounts + ' accounts, ' + outcome.raw.balances + ' balances, ' +
      outcome.raw.transactions + ' transactions.');
    return lines.join('\n');
  }

  ns.SyncAll = { run: run, summarize: summarize };
})(globalThis.BankSync);
