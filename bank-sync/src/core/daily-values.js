/**
 * @fileoverview Pure, provider-agnostic daily value per account, from the balances tagged with a
 * `Balance.valueKind`: 'investments' is the investments value; 'total' (when the provider reports
 * one) is the account value, and cash is then total minus investments, so the total always matches
 * the provider's own figure; without a 'total', cash is the sum of the 'cash' balances and total is
 * cash plus investments. In the account currency and in EUR from the balances' own conversion, so
 * there is no second FX path. Kept as one row per connection, account and day: a later run on the
 * same day replaces that day's row.
 */
(function (ns) {
  const KINDS = ['cash', 'investments', 'total'];

  /**
   * @param {number} n - Amount.
   * @returns {number} Rounded to cents.
   */
  function cents(n) {
    return Math.round(n * 100) / 100;
  }

  /**
   * Builds one daily value per account from fresh, labelled and EUR-converted balances (rules in the
   * file overview). Balances without a kind (and accounts with none) are ignored. If any counted balance
   * has no EUR amount, the EUR columns of that account stay empty rather than show a partial sum.
   * @param {Array<Object>} balances - Balance rows (with labels, `amountEur`, `fxRate`, `fxRateDate`).
   * @param {string} date - ISO date of the run.
   * @returns {Array<import('../providers/bank-provider.js').DailyValue>} Rows, in input order.
   */
  function fromBalances(balances, date) {
    const byAccount = {};
    const order = [];
    balances.forEach(function (b) {
      if (KINDS.indexOf(b.valueKind) === -1) return;
      const key = b.connection + '|' + b.accountId;
      if (!byAccount[key]) {
        order.push(key);
        byAccount[key] = { first: b, hasTotal: false, cash: 0, investments: 0, total: 0, cashEur: 0, investmentsEur: 0, totalEur: 0, eurComplete: true };
      }
      const acc = byAccount[key];
      if (b.valueKind === 'total') acc.hasTotal = true;
      acc[b.valueKind] += b.amount;
      if (typeof b.amountEur === 'number') acc[b.valueKind + 'Eur'] += b.amountEur;
      else acc.eurComplete = false;
    });
    return order.map(function (key) {
      const acc = byAccount[key];
      const b = acc.first;
      const eur = function (n) { return acc.eurComplete ? cents(n) : null; };
      const total = acc.hasTotal ? acc.total : acc.cash + acc.investments;
      const totalEur = acc.hasTotal ? acc.totalEur : acc.cashEur + acc.investmentsEur;
      return {
        date: date,
        provider: b.provider,
        bank: b.bank || '',
        connection: b.connection || '',
        holderName: b.holderName || '',
        accountName: b.accountName || '',
        accountId: b.accountId,
        currency: b.currency,
        cash: cents(cents(total) - cents(acc.investments)),
        investments: cents(acc.investments),
        total: cents(total),
        cashEur: eur(cents(totalEur) - cents(acc.investmentsEur)),
        investmentsEur: eur(acc.investmentsEur),
        totalEur: eur(totalEur),
        fxRate: acc.eurComplete && typeof b.fxRate === 'number' ? b.fxRate : null,
        fxRateDate: acc.eurComplete ? b.fxRateDate || '' : '',
        fetchedAt: b.fetchedAt || ''
      };
    });
  }

  /**
   * Day of a stored row. Cells Sheets turned into dates come back as timestamps, so only the date part counts.
   * @param {Object} row - Daily value row.
   * @returns {string} YYYY-MM-DD.
   */
  function dayOf(row) {
    return String(row.date || '').slice(0, 10);
  }

  /**
   * How the `daily_values` tab grows: one row per connection, account and day (the fresh one wins),
   * never aged out, newest day first.
   * @returns {import('../sheets/raw-writer.js').Accumulate} Policy.
   */
  function accumulate() {
    const keyOf = function (r) { return r.connection && r.accountId && dayOf(r) ? [r.connection, r.accountId, dayOf(r)].join('|') : ''; };
    return {
      keyOf: keyOf,
      keep: function () { return true; },
      compare: function (a, b) {
        if (dayOf(a) !== dayOf(b)) return dayOf(a) < dayOf(b) ? 1 : -1;
        const ka = keyOf(a);
        const kb = keyOf(b);
        return ka < kb ? -1 : ka > kb ? 1 : 0;
      }
    };
  }

  ns.DailyValues = { fromBalances: fromBalances, accumulate: accumulate };
})(globalThis.BankSync);
