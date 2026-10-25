/**
 * @fileoverview Pure idempotent merge of freshly fetched transactions into the
 * stored set. Key: provider + accountId + transactionId.
 */
(function (ns) {
  const FX_FIELDS = ['amountEur', 'fxRate', 'fxRateDate'];

  /**
   * @typedef {Object} FetchedRange
   * @property {string} accountId - Account that was fetched.
   * @property {string} from - First day covered, inclusive (ISO date).
   * @property {string} to - Last day covered, inclusive (ISO date).
   * @property {boolean} [includesUndated] - True if the fetch also covers pending rows
   *   that have no date (the window ending today). Otherwise they are left alone.
   */

  /**
   * @typedef {Object} MergeResult
   * @property {Array<import('../providers/bank-provider.js').Transaction>} rows - New stored set.
   * @property {number} added - Rows not seen before.
   * @property {number} updated - Existing rows whose content changed.
   * @property {number} removed - Stale pending rows dropped.
   */

  /**
   * Builds the identity key of a transaction.
   * @param {import('../providers/bank-provider.js').Transaction} t - Transaction.
   * @returns {string} provider|accountId|transactionId.
   */
  function rowKey(t) {
    return [t.provider, t.accountId, t.transactionId].join('|');
  }

  /**
   * The date a transaction is placed at for range checks and sorting.
   * @param {import('../providers/bank-provider.js').Transaction} t - Transaction.
   * @returns {string} Booking date, else value date, else ''.
   */
  function effectiveDate(t) {
    return t.bookingDate || t.valueDate || '';
  }

  /**
   * Whether a stored pending row falls inside a range that was just re-fetched.
   * Undated pending rows count as inside only for a range flagged `includesUndated`.
   * @param {import('../providers/bank-provider.js').Transaction} t - Stored row.
   * @param {FetchedRange} range - Fetched range.
   * @returns {boolean} True if the row should have been returned again.
   */
  function coveredBy(t, range) {
    if (t.accountId !== range.accountId) return false;
    const date = effectiveDate(t);
    return date ? date >= range.from && date <= range.to : Boolean(range.includesUndated);
  }

  /**
   * Combines an incoming transaction with the stored one it replaces. A booked
   * row is never downgraded to pending, and derived FX columns survive as long
   * as the amount, currency and date the rate depends on are unchanged; otherwise
   * they are left empty so the row is converted again.
   * @param {import('../providers/bank-provider.js').Transaction} previous - Stored row.
   * @param {import('../providers/bank-provider.js').Transaction} incoming - Fetched row.
   * @returns {import('../providers/bank-provider.js').Transaction} Row to store.
   */
  function combine(previous, incoming) {
    if (previous.status === 'booked' && incoming.status === 'pending') return previous;
    const next = Object.assign({}, incoming);
    const sameMoney = previous.amount === incoming.amount && previous.currency === incoming.currency &&
      effectiveDate(previous) === effectiveDate(incoming);
    FX_FIELDS.forEach(function (field) {
      const empty = next[field] === null || next[field] === undefined || next[field] === '';
      if (sameMoney && empty) next[field] = previous[field];
    });
    return next;
  }

  /**
   * Whether two rows differ in any field. Compares values field by field, so the order the keys
   * were set in (a provider's row vs. a row read back from the sheet) does not count as a change.
   * @param {Object} a - First row.
   * @param {Object} b - Second row.
   * @returns {boolean} True if any field differs.
   */
  function changed(a, b) {
    const keys = Object.keys(a).concat(Object.keys(b));
    return keys.some(function (k) { return JSON.stringify(a[k]) !== JSON.stringify(b[k]); });
  }

  /**
   * Orders rows newest first, with the key as a stable tiebreaker.
   * @param {Array<import('../providers/bank-provider.js').Transaction>} rows - Rows.
   * @returns {Array<import('../providers/bank-provider.js').Transaction>} Sorted copy.
   */
  function sortRows(rows) {
    return rows.slice().sort(function (a, b) {
      const da = effectiveDate(a);
      const db = effectiveDate(b);
      if (da !== db) return da < db ? 1 : -1;
      const ka = rowKey(a);
      const kb = rowKey(b);
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    });
  }

  /**
   * Upserts incoming transactions into the stored set. Pending rows inside a
   * fetched range that were not returned again are dropped: they either
   * booked under a new id or were cancelled.
   * @param {Array<import('../providers/bank-provider.js').Transaction>} existing - Stored rows.
   * @param {Array<import('../providers/bank-provider.js').Transaction>} incoming - Fetched rows.
   * @param {Array<FetchedRange>} fetched - Ranges the incoming rows are complete for.
   * @returns {MergeResult} New set and counts. Inputs are not mutated.
   */
  function merge(existing, incoming, fetched) {
    const incomingKeys = {};
    incoming.forEach(function (t) { incomingKeys[rowKey(t)] = true; });

    const byKey = {};
    let removed = 0;
    existing.forEach(function (t) {
      const stale = t.status === 'pending' && !incomingKeys[rowKey(t)] &&
        fetched.some(function (range) { return coveredBy(t, range); });
      if (stale) removed++;
      else byKey[rowKey(t)] = t;
    });

    let added = 0;
    let updated = 0;
    incoming.forEach(function (t) {
      const key = rowKey(t);
      const previous = byKey[key];
      const next = previous ? combine(previous, t) : t;
      if (!previous) added++;
      else if (changed(previous, next)) updated++;
      byKey[key] = next;
    });

    return {
      rows: sortRows(Object.keys(byKey).map(function (key) { return byKey[key]; })),
      added: added,
      updated: updated,
      removed: removed
    };
  }

  ns.TxnMerge = { merge: merge, rowKey: rowKey, effectiveDate: effectiveDate };
})(globalThis.BankSync);
