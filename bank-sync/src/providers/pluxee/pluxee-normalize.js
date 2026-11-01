/**
 * @fileoverview Pure mapping from the Pluxee portal's card transactions response
 * (`{card, transactions}`) to the normalized models: the card as one account, and its approved
 * payments and top-ups as transactions. Amounts come as `{value, exponent, currency}` in minor
 * units (-6787 with exponent 2 is -67.87). Declined payments never moved money and are left out.
 * Every row is a card row (`classification` = 'card'), like Trading 212 card spending.
 */
(function (ns) {
  const PROVIDER_ID = ns.Config.PLUXEE.PROVIDER_ID;
  /** Statuses that never moved money; such rows are dropped. */
  const SKIPPED = ['DECLINED'];
  /** The status of a settled movement; anything else (a status not seen yet) is kept as pending. */
  const BOOKED = 'APPROVED';

  /**
   * @param {*} value - Text from the API.
   * @returns {string} Trimmed, with runs of spaces collapsed ('Shop   Town ' becomes 'Shop Town').
   */
  function clean(value) {
    return value === null || value === undefined ? '' : String(value).replace(/\s+/g, ' ').trim();
  }

  /**
   * Maps the response's card to the one account of a connection.
   * @param {Object} card - `card` of the response (cardId, maskedPan, name...).
   * @param {string} cardId - The configured card id (used when the response names none).
   * @param {import('../../core/config.js').ConnectionConfig} connection - Connection.
   * @returns {import('../bank-provider.js').Account} Account; `accountId` is the card id.
   */
  function account(card, cardId, connection) {
    const c = card || {};
    const name = clean(c.name || c.productName || c.label);
    return {
      provider: PROVIDER_ID,
      accountId: clean(c.cardId) || cardId,
      bank: connection.bank.aspspName,
      connection: connection.key,
      holderName: '',
      accountName: name || connection.bank.aspspName + ' card',
      currency: 'EUR',
      ibanMasked: ns.Redact.maskAccount(clean(c.maskedPan)),
      accountType: 'CARD',
      product: name,
      detailsAt: '',
      apiFields: ns.Flatten.flat(c)
    };
  }

  /**
   * @param {Object} money - `{value, exponent, currency}`.
   * @returns {{amount: number, currency: string}} Amount in major units.
   * @throws {Error} DataError when the value or exponent is not a whole number.
   */
  function amountOf(money) {
    const m = money || {};
    const value = Number(m.value);
    const exponent = m.exponent === undefined || m.exponent === null ? 2 : Number(m.exponent);
    if (m.value === null || m.value === undefined || m.value === '' || !Number.isInteger(value)) {
      throw ns.Errors.create('DataError', 'Pluxee transaction amount is not numeric');
    }
    if (!Number.isInteger(exponent) || exponent < 0) throw ns.Errors.create('DataError', 'Pluxee amount exponent is not usable');
    return { amount: Number((value / Math.pow(10, exponent)).toFixed(exponent)), currency: clean(m.currency) || 'EUR' };
  }

  /**
   * Local date of a Pluxee timestamp. The portal sends local times without a zone
   * ('2026-10-25T09:42:04.000'); their date is taken as is. A timestamp with a zone goes
   * through `toLocalDate`.
   * @param {*} value - Timestamp.
   * @param {function(string): string} toLocalDate - ISO date-time to YYYY-MM-DD in the Bank Sync time zone.
   * @returns {string} YYYY-MM-DD.
   * @throws {Error} DataError when missing or invalid.
   */
  function localDate(value, toLocalDate) {
    const text = clean(value);
    if (!/^\d{4}-\d{2}-\d{2}/.test(text) || isNaN(Date.parse(text.slice(0, 10)))) {
      throw ns.Errors.create('DataError', 'Pluxee transaction has no valid date');
    }
    return /(Z|[+-]\d{2}:?\d{2})$/.test(text) ? toLocalDate(text) : text.slice(0, 10);
  }

  /**
   * Maps one transaction.
   * @param {Object} item - {id, description, date, status, type, amount, merchantName, merchantCity, ...}.
   * @param {string} accountId - Card id.
   * @param {function(string): string} toLocalDate - Date-time to local date.
   * @returns {?import('../bank-provider.js').Transaction} Transaction, or null for a declined one.
   * @throws {Error} DataError when the id, amount or date is unusable.
   */
  function transaction(item, accountId, toLocalDate) {
    const status = clean(item.status).toUpperCase();
    if (SKIPPED.indexOf(status) !== -1) return null;
    if (!clean(item.id)) throw ns.Errors.create('DataError', 'Pluxee transaction has no id');
    const money = amountOf(item.amount);
    const date = localDate(item.date, toLocalDate);
    const description = clean(item.description);
    return {
      provider: PROVIDER_ID,
      accountId: accountId,
      transactionId: clean(item.id),
      bookingDate: date,
      valueDate: date,
      status: status === BOOKED ? 'booked' : 'pending',
      amount: money.amount,
      currency: money.currency,
      amountEur: null,
      fxRate: null,
      fxRateDate: '',
      counterparty: clean(item.merchantName) || description,
      description: description,
      category: clean(item.type),
      classification: 'card',
      counterpartyAccount: '',
      counterpartyBank: '',
      reference: '',
      apiFields: ns.Flatten.flat(item)
    };
  }

  /**
   * Maps every transaction of a response, declined ones left out.
   * @param {Array<Object>} items - `transactions` of the response.
   * @param {string} accountId - Card id.
   * @param {function(string): string} toLocalDate - Date-time to local date.
   * @returns {Array<import('../bank-provider.js').Transaction>} Transactions.
   */
  function transactions(items, accountId, toLocalDate) {
    return (items || []).map(function (i) { return transaction(i, accountId, toLocalDate); }).filter(Boolean);
  }

  ns.PluxeeNormalize = { account: account, amountOf: amountOf, localDate: localDate, transaction: transaction, transactions: transactions };
})(globalThis.BankSync);
