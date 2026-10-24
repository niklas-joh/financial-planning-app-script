/**
 * @fileoverview Pure mapping from Enable Banking payloads to normalized models.
 */
(function (ns) {
  const PROVIDER_ID = ns.Config.ENABLE_BANKING.PROVIDER_ID;

  /**
   * Maps an Enable Banking AccountResource to an Account.
   * @param {Object} raw - AccountResource (needs `uid` and `currency`).
   * @returns {import('../bank-provider.js').Account} Normalized account; IBAN masked.
   * @throws {Error} If `uid` is missing.
   */
  function normalizeAccount(raw) {
    if (!raw || !raw.uid) throw ns.Errors.create('DataError', 'Account is missing uid');
    return {
      provider: PROVIDER_ID,
      accountId: raw.uid,
      name: raw.name || raw.product || '',
      currency: raw.currency || '',
      ibanMasked: ns.Redact.maskIban(raw.account_id && raw.account_id.iban),
      accountType: raw.cash_account_type || '',
      product: raw.product || ''
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
      name: raw.name || '',
      amount: amount,
      currency: raw.balance_amount.currency || '',
      referenceDate: raw.reference_date || '',
      lastChange: raw.last_change_date_time || '',
      fetchedAt: fetchedAt
    };
  }

  ns.EbNormalize = { normalizeAccount: normalizeAccount, normalizeBalance: normalizeBalance };
})(globalThis.BankSync);
