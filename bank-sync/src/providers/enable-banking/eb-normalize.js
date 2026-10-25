/**
 * @fileoverview Pure mapping from Enable Banking payloads to normalized models.
 */
(function (ns) {
  const PROVIDER_ID = ns.Config.ENABLE_BANKING.PROVIDER_ID;

  /**
   * Maps an Enable Banking AccountResource to an Account. The API's `name` is the
   * account holder's name (`holderName`); `details` is the label set by the user or
   * bank (`accountName`). `bank` and `connection` are stamped later by the provider.
   * @param {Object} raw - AccountResource (needs `uid` and `currency`).
   * @param {string} [detailsAt] - ISO date-time the details were read, '' if unknown.
   * @returns {import('../bank-provider.js').Account} Normalized account; IBAN masked, every API
   *   field kept (flattened) in `apiFields`.
   * @throws {Error} If `uid` is missing.
   */
  function normalizeAccount(raw, detailsAt) {
    if (!raw || !raw.uid) throw ns.Errors.create('DataError', 'Account is missing uid');
    return {
      provider: PROVIDER_ID,
      accountId: raw.uid,
      bank: '',
      connection: '',
      holderName: raw.name || '',
      accountName: raw.details || '',
      currency: raw.currency || '',
      ibanMasked: ns.Redact.maskIban(raw.account_id && raw.account_id.iban),
      accountType: raw.cash_account_type || '',
      product: raw.product || '',
      detailsAt: detailsAt || '',
      apiFields: ns.Flatten.flat(raw)
    };
  }

  /**
   * Maps an Enable Banking BalanceResource to a Balance.
   * @param {Object} raw - BalanceResource (`balance_amount.amount` is a decimal string).
   * @param {string} accountId - Account the balance was read from.
   * @param {string} fetchedAt - ISO date-time of the read.
   * @returns {import('../bank-provider.js').Balance} Normalized balance.
   * @throws {Error} If the amount is not a finite number.
   */
  function normalizeBalance(raw, accountId, fetchedAt) {
    const amount = Number(raw && raw.balance_amount && raw.balance_amount.amount);
    if (!isFinite(amount)) throw ns.Errors.create('DataError', 'Balance amount is not numeric');
    return {
      provider: PROVIDER_ID,
      accountId: accountId,
      balanceType: raw.balance_type || '',
      description: raw.name || '',
      amount: amount,
      currency: raw.balance_amount.currency || '',
      referenceDate: raw.reference_date || '',
      lastChange: raw.last_change_date_time || '',
      fetchedAt: fetchedAt,
      apiFields: ns.Flatten.flat(raw)
    };
  }

  /** Enable Banking transaction status to normalized status; others are skipped. */
  const STATUS_MAP = { BOOK: 'booked', PDNG: 'pending', HOLD: 'pending' };

  /**
   * Signed amount: the API reports a positive amount plus a credit/debit indicator.
   * @param {Object} raw - Enable Banking transaction.
   * @returns {number} Negative for debits (DBIT), positive for credits (CRDT).
   * @throws {Error} DataError if the amount or indicator is unusable.
   */
  function signedAmount(raw) {
    const text = raw.transaction_amount && raw.transaction_amount.amount;
    const magnitude = text === '' || text === null || text === undefined ? NaN : Math.abs(Number(text));
    if (!isFinite(magnitude)) throw ns.Errors.create('DataError', 'Transaction amount is not numeric');
    if (raw.credit_debit_indicator === 'DBIT') return -magnitude;
    if (raw.credit_debit_indicator === 'CRDT') return magnitude;
    throw ns.Errors.create('DataError', 'Transaction has no credit/debit indicator');
  }

  /**
   * The other party: who was paid for a debit, who paid for a credit.
   * @param {Object} raw - Enable Banking transaction.
   * @returns {string} Party name, or ''.
   */
  function counterpartyName(raw) {
    const party = raw.credit_debit_indicator === 'DBIT' ? raw.creditor : raw.debtor;
    return (party && party.name) || '';
  }

  /**
   * Joined free-text description.
   * @param {Object} raw - Enable Banking transaction.
   * @returns {string} Remittance lines joined by a space, else the note, else ''.
   */
  function descriptionOf(raw) {
    const lines = (raw.remittance_information || []).filter(Boolean);
    return lines.length ? lines.join(' ') : (raw.note || '');
  }

  /**
   * Maps one transaction without an id (the id is assigned by the batch mapper).
   * @param {Object} raw - Enable Banking transaction.
   * @param {string} accountId - Account the transaction belongs to.
   * @param {string} status - Normalized status.
   * @returns {import('../bank-provider.js').Transaction} Transaction with an empty id.
   */
  function toTransaction(raw, accountId, status) {
    return {
      provider: PROVIDER_ID,
      bank: '',
      connection: '',
      holderName: '',
      accountName: '',
      ibanMasked: '',
      accountId: accountId,
      transactionId: '',
      bookingDate: raw.booking_date || '',
      valueDate: raw.value_date || '',
      status: status,
      amount: signedAmount(raw),
      currency: raw.transaction_amount.currency || '',
      amountEur: null,
      fxRate: null,
      fxRateDate: '',
      counterparty: counterpartyName(raw),
      description: descriptionOf(raw),
      apiFields: ns.Flatten.flat(raw)
    };
  }

  /**
   * Content fingerprint for rows the bank gave no id. Status is left out so a
   * pending row keeps its id when it books unchanged.
   * @param {import('../bank-provider.js').Transaction} t - Transaction.
   * @returns {string} Fingerprint text.
   */
  function fingerprint(t) {
    return [t.bookingDate || t.valueDate, t.amount, t.currency, t.counterparty, t.description].join('|');
  }

  /**
   * Maps all raw transactions of one fetch. The id is `entry_reference` (documented
   * as unique and immutable per account), else `transaction_id`, else a content hash
   * with an occurrence counter so genuinely identical rows stay distinct.
   * Cancelled, rejected, scheduled and unknown-status rows are skipped. Every API field is
   * kept (flattened, account numbers masked) in `apiFields`; callers strip it before storing.
   * @param {Array<Object>} rawList - Enable Banking transactions (all pages of one fetch).
   * @param {string} accountId - Account the transactions belong to.
   * @returns {{transactions: Array<import('../bank-provider.js').Transaction>, skipped: number}}
   *   Normalized rows and the count of skipped ones.
   * @throws {Error} DataError for an unusable amount or indicator.
   */
  function normalizeTransactions(rawList, accountId) {
    const seen = {};
    const transactions = [];
    let skipped = 0;
    rawList.forEach(function (raw) {
      const status = STATUS_MAP[raw.status || 'BOOK'];
      if (!status) {
        skipped++;
        return;
      }
      const t = toTransaction(raw, accountId, status);
      let id = raw.entry_reference || raw.transaction_id;
      if (!id) {
        const print = fingerprint(t);
        seen[print] = (seen[print] || 0) + 1;
        id = 'h' + ns.Hash.fnv1a(print) + '-' + seen[print];
      }
      t.transactionId = String(id);
      transactions.push(t);
    });
    return { transactions: transactions, skipped: skipped };
  }

  ns.EbNormalize = {
    normalizeAccount: normalizeAccount,
    normalizeBalance: normalizeBalance,
    normalizeTransactions: normalizeTransactions
  };
})(globalThis.BankSync);
