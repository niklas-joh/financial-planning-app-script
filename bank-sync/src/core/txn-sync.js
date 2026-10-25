/**
 * @fileoverview Provider-agnostic transaction sync. Fetches oldest-to-newest in
 * date windows and saves once per account (also when a fetch fails or the time
 * budget ends the run), so nothing merged so far is lost: the next run restarts
 * each account from its latest booked date.
 */
(function (ns) {
  /**
   * @typedef {Object} TxnSyncDeps
   * @property {import('../providers/bank-provider.js').BankProvider} provider - Source.
   * @property {Array<import('../providers/bank-provider.js').Account>} accounts - The provider's accounts.
   * @property {number} [budgetMs] - Time allowed for this run (default `TIME_BUDGET_MS`).
   * @property {{load: function(): Array<Object>, save: function(Array<Object>): number}} store - Sheet store.
   * @property {function(): number} nowMs - Clock.
   * @property {function(): string} todayIso - Today's date (YYYY-MM-DD) in the configured time zone.
   */

  /**
   * @typedef {Object} TxnSyncResult
   * @property {number} added - New rows stored.
   * @property {number} updated - Existing rows changed.
   * @property {number} removed - Stale pending rows dropped.
   * @property {boolean} complete - False if the time budget ended the run early; run again.
   * @property {Array<import('../providers/bank-provider.js').Transaction>} rawRows - Everything this run
   *   fetched, labelled and with `apiFields`, for the raw tab. Never stored on `transactions`.
   */

  /**
   * Latest booked date stored for an account.
   * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Stored rows.
   * @param {string} accountId - Account id.
   * @returns {string} ISO date, or '' if none.
   */
  function latestBooked(rows, accountId) {
    return rows.reduce(function (latest, t) {
      const date = t.accountId === accountId && t.status === 'booked' ? t.bookingDate : '';
      return date > latest ? date : latest;
    }, '');
  }

  /**
   * @param {import('../providers/bank-provider.js').Transaction} t - Transaction with `apiFields`.
   * @returns {import('../providers/bank-provider.js').Transaction} Copy without `apiFields`.
   */
  function withoutApiFields(t) {
    const copy = Object.assign({}, t);
    delete copy.apiFields;
    return copy;
  }

  /**
   * Syncs every account of the provider.
   * @param {TxnSyncDeps} deps - Injected services.
   * @returns {TxnSyncResult} Counts, plus `rawRows` for the raw tab.
   * @throws {Error} Any provider or store error, e.g. AuthError when consent lapsed.
   */
  function run(deps) {
    const cfg = ns.Config.TRANSACTIONS;
    const budgetMs = deps.budgetMs === undefined ? cfg.TIME_BUDGET_MS : deps.budgetMs;
    const labelsFor = ns.AccountLabels.lookup(deps.accounts);
    const startedAt = deps.nowMs();
    const today = deps.todayIso();
    const result = { added: 0, updated: 0, removed: 0, complete: true, rawRows: [] };
    let rows = deps.store.load().map(function (t) { return ns.AccountLabels.stamp(t, labelsFor); });

    deps.accounts.forEach(function (account) {
      const start = ns.DateRange.startFor(latestBooked(rows, account.accountId), today, cfg);
      let dirty = false;
      try {
        ns.DateRange.planWindows(start, today, cfg.WINDOW_DAYS).forEach(function (window) {
          if (!result.complete || deps.nowMs() - startedAt > budgetMs) {
            result.complete = false;
            return;
          }
          const incoming = deps.provider
            .getTransactions(account.accountId, { dateFrom: window.from, dateTo: window.to })
            .map(function (t) { return ns.AccountLabels.stamp(t, labelsFor); });
          result.rawRows = result.rawRows.concat(incoming);
          const merged = ns.TxnMerge.merge(rows, incoming.map(withoutApiFields), [{
            accountId: account.accountId, from: window.from, to: window.to, includesUndated: window.to === today
          }]);
          rows = merged.rows;
          dirty = true;
          result.added += merged.added;
          result.updated += merged.updated;
          result.removed += merged.removed;
        });
      } finally {
        if (dirty) deps.store.save(rows);
      }
    });
    return result;
  }

  ns.TxnSync = { run: run, latestBooked: latestBooked };
})(globalThis.BankSync);
