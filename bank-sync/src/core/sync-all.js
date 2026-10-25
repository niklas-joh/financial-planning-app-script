/**
 * @fileoverview Syncs every configured connection in turn: raw accounts, balances and
 * (for providers that hold investments) holdings, then transactions; then adds the details
 * some providers give for other providers' rows (e.g. card merchants from a CSV report); then converts
 * balances and holdings (latest rate) and the stored
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
   * @property {number} holdingsMissing - Fresh holdings without a rate.
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
   * @returns {{data: import('./raw-sync.js').RawData, transactions: Object,
   *   enrichments: Array<import('./txn-merge.js').Enrichment>}} What was read and the counts.
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
    const enrichments = typeof provider.getEnrichments === 'function' ? provider.getEnrichments() : [];
    return { data: data, transactions: transactions, enrichments: enrichments };
  }

  /**
   * Applies the details read in this run to the stored rows (after every connection, so rows fetched
   * in the same run are there). A failure is reported, never thrown: the rest of the run still lands.
   * @param {Array<import('./txn-merge.js').Enrichment>} details - Details from all providers.
   * @param {SyncAllDeps} deps - Services.
   * @returns {?{read: number, enriched: number, error: (string|undefined)}} Counts, or null if none were read.
   */
  function applyEnrichments(details, deps) {
    if (!details.length) return null;
    const outcome = { read: details.length, enriched: 0 };
    try {
      const done = ns.TxnMerge.enrich(deps.store.load(), details);
      if (done.enriched) deps.store.save(done.rows);
      outcome.enriched = done.enriched;
    } catch (e) {
      outcome.error = ns.Redact.text(e.name + ': ' + e.message);
    }
    return outcome;
  }

  /**
   * Runs the sync for all connections.
   * @param {SyncAllDeps} deps - Injected services.
   * @returns {{results: Array<ConnectionResult>, raw: ?Object, details: ?Object, fx: ?FxOutcome}}
   *   Per-connection outcome, the row counts from `RawSync.write` (null if no connection could be read), the
   *   details applied (null if none were read) and the EUR conversion counts (null without `deps.fx`). No
   *   transaction content.
   */
  function run(deps) {
    const cfg = ns.Config.TRANSACTIONS;
    const startedAt = deps.nowMs();
    const collected = { accounts: [], balances: [], transactions: [], holdings: [], holdingConnections: [] };
    const refreshed = [];
    let enrichments = [];

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
          if (done.data.holdings) {
            collected.holdings = collected.holdings.concat(done.data.holdings);
            collected.holdingConnections.push(connection.key);
          }
          enrichments = enrichments.concat(done.enrichments);
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

    const fx = deps.fx ? { changed: 0, missing: 0, waiting: 0, balancesMissing: 0, holdingsMissing: 0, failed: [] } : null;
    if (fx && refreshed.length) {
      guardFx(fx, function () {
        const converted = deps.fx.convertBalances(collected.balances);
        collected.balances = converted.rows;
        fx.balancesMissing = converted.missing;
        fx.failed = converted.failed;
        if (!collected.holdings.length) return;
        const holdings = deps.fx.convertBalances(collected.holdings);
        collected.holdings = holdings.rows;
        fx.holdingsMissing = holdings.missing;
        fx.failed = holdings.failed;
      });
    }
    const raw = refreshed.length ? ns.RawSync.write(deps.rawWriter, collected, refreshed, deps.todayIso()) : null;
    const details = applyEnrichments(enrichments, deps);
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
    return { results: results, raw: raw, details: details, fx: fx };
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
   * @param {{results: Array<ConnectionResult>, raw: ?Object, details: ?Object, fx: ?FxOutcome}} outcome - From `run`.
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
    const raw = outcome.raw;
    if (raw) lines.push('Raw tabs: ' + raw.accounts + ' accounts, ' + raw.balances + ' balances, ' + raw.transactions + ' transactions' +
      (raw.holdings === undefined ? '' : ', ' + raw.holdings + ' holdings') +
      (raw.dailyValues === undefined ? '' : ', ' + raw.dailyValues + ' daily values') + '.');
    const details = outcome.details;
    if (details) {
      lines.push('Details from reports: ' + details.enriched + ' transactions updated (' + details.read + ' read).' +
        (details.error ? ' Failed (' + details.error + ')' : ''));
    }
    const fx = outcome.fx;
    if (fx) {
      lines.push('EUR conversion: ' + fx.changed + ' transactions converted' +
        (fx.missing ? ', ' + fx.missing + ' without a rate' : '') +
        (fx.waiting ? ', ' + fx.waiting + " waiting for today's rate" : '') +
        (fx.balancesMissing ? ', ' + fx.balancesMissing + ' balances without a rate' : '') +
        (fx.holdingsMissing ? ', ' + fx.holdingsMissing + ' holdings without a rate' : '') + '.');
      fx.failed.forEach(function (f) { lines.push('Rates unavailable for ' + f); });
      if (fx.error) lines.push('EUR conversion failed (' + fx.error + ')');
    }
    return lines.join('\n');
  }

  ns.SyncAll = { run: run, summarize: summarize };
})(globalThis.BankSync);
