/**
 * @fileoverview Provider-agnostic transaction sync. Fetches oldest-to-newest in
 * date windows and saves after each window, so a run that hits the time budget
 * loses nothing: the next run restarts each account from its latest booked date.
 */
(function (ns) {
  /**
   * @typedef {Object} TxnSyncDeps
   * @property {import('../providers/bank-provider.js').BankProvider} provider - Source.
   * @property {{load: function(): Array<Object>, save: function(Array<Object>): number}} store - Sheet store.
   * @property {function(): number} nowMs - Clock.
   */

  /**
   * @typedef {Object} TxnSyncResult
   * @property {number} added - New rows stored.
   * @property {number} updated - Existing rows changed.
   * @property {number} removed - Stale pending rows dropped.
   * @property {boolean} complete - False if the time budget ended the run early; run again.
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
   * Syncs every account of the provider.
   * @param {TxnSyncDeps} deps - Injected services.
   * @returns {TxnSyncResult} Counts only (no transaction content).
   * @throws {Error} Any provider or store error, e.g. AuthError when consent lapsed.
   */
  function run(deps) {
    const cfg = ns.Config.TRANSACTIONS;
    const startedAt = deps.nowMs();
    const today = ns.DateRange.isoDate(startedAt);
    const result = { added: 0, updated: 0, removed: 0, complete: true };
    let rows = deps.store.load();

    deps.provider.listAccounts().forEach(function (account) {
      const start = ns.DateRange.startFor(latestBooked(rows, account.accountId), today, cfg);
      ns.DateRange.planWindows(start, today, cfg.WINDOW_DAYS).forEach(function (window) {
        if (!result.complete || deps.nowMs() - startedAt > cfg.TIME_BUDGET_MS) {
          result.complete = false;
          return;
        }
        const incoming = deps.provider.getTransactions(account.accountId, { dateFrom: window.from, dateTo: window.to });
        const merged = ns.TxnMerge.merge(rows, incoming,
          [{ accountId: account.accountId, from: window.from, to: window.to }]);
        rows = merged.rows;
        deps.store.save(rows);
        result.added += merged.added;
        result.updated += merged.updated;
        result.removed += merged.removed;
      });
    });
    return result;
  }

  ns.TxnSync = { run: run, latestBooked: latestBooked };
})(globalThis.BankSync);
