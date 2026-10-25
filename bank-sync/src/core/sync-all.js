/**
 * @fileoverview Syncs every configured connection in turn: raw accounts and
 * balances, then transactions; then converts balances (latest rate) and the stored
 * transactions that have no EUR amount yet (booking-date rate) to EUR, fetching rates once for the whole run.
 * One connection failing (e.g. a lapsed consent) never stops the others, and the
 * time budget is shared.
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
   * @property {{convertTransactions: function(Array<Object>): Object, convertBalances: function(Array<Object>): Object}} [fx]
   *   FX service for this run; without it the EUR columns are left as they are.
   */

  /**
   * @typedef {Object} FxOutcome
   * @property {number} changed - Transactions whose EUR columns were set or corrected.
   * @property {number} missing - Transactions still without a rate.
   * @property {number} waiting - Transactions dated today whose rate is not published yet.
   * @property {number} balancesMissing - Fresh balances without a rate.
   * @property {Array<string>} failed - Redacted per-currency fetch errors.
   * @property {string} [error] - Redacted message if the FX step itself failed.
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
    const outOfTime = function () { return deps.nowMs() - startedAt > ns.Config.TRANSACTIONS.TIME_BUDGET_MS; };
    const accounts = provider.listAccounts({ outOfTime: outOfTime });
    const data = ns.RawSync.collect(provider, accounts, outOfTime);
    const transactions = ns.TxnSync.run({
      provider: provider, accounts: accounts, store: deps.store,
      nowMs: deps.nowMs, todayIso: deps.todayIso, budgetMs: budgetMs
    });
    return { data: data, transactions: transactions };
  }

  /**
   * Runs the sync for all connections.
   * @param {SyncAllDeps} deps - Injected services.
   * @returns {{results: Array<ConnectionResult>, raw: ?{accounts: number, balances: number, transactions: number}, fx: ?FxOutcome}}
   *   Per-connection outcome, the raw rows written (null if no connection could be read) and the EUR
   *   conversion counts (null without `deps.fx`). No transaction content.
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

    const fx = deps.fx ? { changed: 0, missing: 0, waiting: 0, balancesMissing: 0, failed: [] } : null;
    if (fx && refreshed.length) {
      guardFx(fx, function () {
        const converted = deps.fx.convertBalances(collected.balances);
        collected.balances = converted.rows;
        fx.balancesMissing = converted.missing;
        fx.failed = converted.failed;
      });
    }
    const raw = refreshed.length ? ns.RawSync.write(deps.rawWriter, collected, refreshed, deps.todayIso()) : null;
    if (fx) {
      guardFx(fx, function () {
        const converted = deps.fx.convertTransactions(deps.store.load());
        if (converted.changed) deps.store.save(converted.rows);
        fx.changed = converted.changed;
        fx.missing = converted.missing;
        fx.waiting = converted.waiting;
        fx.failed = converted.failed;
      });
    }
    return { results: results, raw: raw, fx: fx };
  }

  /**
   * Runs one FX step so that a failure (e.g. the Rates tab cannot be written) is reported
   * instead of stopping the sync; the EUR columns are then simply left as they were.
   * @param {FxOutcome} fx - Outcome to record the error on.
   * @param {function(): void} step - The step.
   */
  function guardFx(fx, step) {
    try {
      step();
    } catch (e) {
      fx.error = ns.Redact.text(e.name + ': ' + e.message);
    }
  }

  /**
   * Turns a run into the text shown to the user. Counts only, no transaction content.
   * @param {{results: Array<ConnectionResult>, raw: ?Object, fx: ?FxOutcome}} outcome - From `run`.
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
    const fx = outcome.fx;
    if (fx) {
      lines.push('EUR conversion: ' + fx.changed + ' transactions converted' +
        (fx.missing ? ', ' + fx.missing + ' without a rate' : '') +
        (fx.waiting ? ', ' + fx.waiting + " waiting for today's rate" : '') +
        (fx.balancesMissing ? ', ' + fx.balancesMissing + ' balances without a rate' : '') + '.');
      fx.failed.forEach(function (f) { lines.push('Rates unavailable for ' + f); });
      if (fx.error) lines.push('EUR conversion failed (' + fx.error + ')');
    }
    return lines.join('\n');
  }

  ns.SyncAll = { run: run, summarize: summarize };
})(globalThis.BankSync);
