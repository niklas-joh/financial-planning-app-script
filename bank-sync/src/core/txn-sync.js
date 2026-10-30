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
   * @property {boolean} [fullHistory] - Re-read each account from its oldest stored row instead of the latest booked
   *   date minus the overlap, so every stored row is fetched again (*Re-process stored rows*). A window the
   *   provider cannot serve (e.g. older than the bank keeps) is counted in `unavailable` and skipped instead of
   *   stopping the account; consent errors still stop it.
   */

  /**
   * @typedef {Object} TxnSyncResult
   * @property {number} added - New rows stored.
   * @property {number} updated - Existing rows changed.
   * @property {number} removed - Stale pending rows dropped.
   * @property {boolean} complete - False if the time budget ended the run early; run again.
   * @property {number} [unavailable] - With `fullHistory`: windows the provider could not serve (rows in them stay
   *   as they were).
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
   * Oldest stored date for an account, booked or pending.
   * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Stored rows.
   * @param {string} accountId - Account id.
   * @returns {string} ISO date, or '' if none.
   */
  function earliestStored(rows, accountId) {
    return rows.reduce(function (earliest, t) {
      const date = t.accountId === accountId ? ns.TxnMerge.effectiveDate(t).slice(0, 10) : '';
      return date && (!earliest || date < earliest) ? date : earliest;
    }, '');
  }

  /**
   * Whether an error means the consent or login is the problem (so re-reading other windows cannot help).
   * @param {Error} e - Error from the provider.
   * @returns {boolean} True for AuthError and HTTP 401/403.
   */
  function isAuthError(e) {
    return e.name === 'AuthError' || e.status === 401 || e.status === 403;
  }

  /**
   * First day to fetch for an account.
   * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Stored rows.
   * @param {string} accountId - Account id.
   * @param {string} today - ISO date of today.
   * @param {boolean} fullHistory - Start at the oldest stored row (when there is one).
   * @returns {string} ISO date.
   */
  function startOf(rows, accountId, today, fullHistory) {
    const oldest = fullHistory ? earliestStored(rows, accountId) : '';
    return oldest && oldest < today ? oldest : ns.DateRange.startFor(latestBooked(rows, accountId), today, ns.Config.TRANSACTIONS);
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
    if (deps.fullHistory) result.unavailable = 0;
    let rows = deps.store.load().map(function (t) { return ns.AccountLabels.stamp(t, labelsFor); });

    deps.accounts.forEach(function (account) {
      const start = startOf(rows, account.accountId, today, Boolean(deps.fullHistory));
      let dirty = false;
      try {
        ns.DateRange.planWindows(start, today, deps.provider.maxWindowDays || cfg.WINDOW_DAYS).forEach(function (window) {
          if (!result.complete || deps.nowMs() - startedAt > budgetMs) {
            result.complete = false;
            return;
          }
          let fetched;
          try {
            fetched = deps.provider.getTransactions(account.accountId, { dateFrom: window.from, dateTo: window.to });
          } catch (e) {
            if (!deps.fullHistory || isAuthError(e)) throw e;
            result.unavailable++;
            return;
          }
          const incoming = fetched.map(function (t) { return ns.AccountLabels.stamp(t, labelsFor); });
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

  ns.TxnSync = { run: run, latestBooked: latestBooked, earliestStored: earliestStored };
})(globalThis.BankSync);
