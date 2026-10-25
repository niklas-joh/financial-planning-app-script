/**
 * @fileoverview Pure mapping from Trading 212 public API payloads to the normalized
 * models: one account per API key, its cash and investment balances, holdings, and
 * cash movements and dividends as transactions. Amounts come in the account's primary
 * currency; instrument prices in the instrument's currency.
 */
(function (ns) {
  const PROVIDER_ID = ns.Config.TRADING212.PROVIDER_ID;

  /**
   * Summary amounts that become balances; `kind` puts them in the daily value. The daily value takes
   * Trading 212's own total and investments value (which already contains the cash in pies), and cash
   * as the rest, so amounts the summary does not break down (seen live: a small gap between the total
   * and available cash plus investments) still count. The cash rows are kept as detail.
   */
  const SUMMARY_BALANCES = [
    { path: ['cash', 'availableToTrade'], description: 'Cash available to trade', kind: '' },
    { path: ['cash', 'inPies'], description: 'Cash in pies, not yet invested (part of the investments value)', kind: '' },
    { path: ['cash', 'reservedForOrders'], description: 'Cash reserved for pending orders', kind: '' },
    { path: ['investments', 'currentValue'], description: 'Current value of investments, including cash in pies', kind: 'investments' },
    { path: ['totalValue'], description: 'Total account value', kind: 'total' }
  ];

  /** Cash movement types with a known direction; others (TRANSFER, new types) keep the API's sign. */
  const OUTFLOWS = ['WITHDRAW', 'FEE'];
  const INFLOWS = ['DEPOSIT', 'INTEREST_ON_FREE_CASH', 'LENDING_INTEREST'];

  /**
   * Reads a number that may be absent.
   * @param {*} value - API value.
   * @param {string} what - Field name for the error.
   * @returns {?number} The number, or null when absent.
   * @throws {Error} DataError if present but not a finite number.
   */
  function optionalNumber(value, what) {
    if (value === null || value === undefined || value === '') return null;
    const n = Number(value);
    if (!isFinite(n)) throw ns.Errors.create('DataError', what + ' is not numeric');
    return n;
  }

  /**
   * Reads a number that must be there.
   * @param {*} value - API value.
   * @param {string} what - Field name for the error.
   * @returns {number} The number.
   * @throws {Error} DataError if absent or not numeric.
   */
  function requiredNumber(value, what) {
    const n = optionalNumber(value, what);
    if (n === null) throw ns.Errors.create('DataError', what + ' is missing');
    return n;
  }

  /**
   * @param {Object} obj - Object.
   * @param {Array<string>} path - Keys.
   * @returns {*} The nested value, or undefined.
   */
  function dig(obj, path) {
    return path.reduce(function (o, key) { return o === null || o === undefined ? undefined : o[key]; }, obj);
  }

  /**
   * Maps the account summary to the one account behind an API key.
   * @param {Object} summary - `GET /equity/account/summary` (needs `id` and `currency`).
   * @param {import('../../core/config.js').ConnectionConfig} connection - Connection the key belongs to.
   * @returns {import('../bank-provider.js').Account} Account; `accountId` is the Trading 212 account number.
   * @throws {Error} DataError if `id` or `currency` is missing.
   */
  function account(summary, connection) {
    if (!summary || summary.id === undefined || summary.id === null || !summary.currency) {
      throw ns.Errors.create('DataError', 'Account summary is missing id or currency');
    }
    return {
      provider: PROVIDER_ID,
      accountId: String(summary.id),
      bank: connection.bank.aspspName,
      connection: connection.key,
      holderName: '',
      accountName: connection.bank.aspspName + ' ' + summary.currency,
      currency: summary.currency,
      ibanMasked: '',
      accountType: '',
      product: '',
      detailsAt: '',
      apiFields: { id: summary.id, currency: summary.currency }
    };
  }

  /**
   * Maps the summary's cash and investment figures to balances. Absent figures are skipped.
   * @param {Object} summary - Account summary.
   * @param {string} accountId - Account id.
   * @param {string} fetchedAt - ISO date-time of the read.
   * @returns {Array<import('../bank-provider.js').Balance>} Balances; `balanceType` is the API field path.
   * @throws {Error} DataError if a figure is not numeric.
   */
  function balances(summary, accountId, fetchedAt) {
    return SUMMARY_BALANCES.map(function (spec) {
      const type = spec.path.join('.');
      const amount = optionalNumber(dig(summary, spec.path), type);
      return amount === null ? null : {
        provider: PROVIDER_ID,
        accountId: accountId,
        balanceType: type,
        description: spec.description,
        valueKind: spec.kind,
        amount: amount,
        currency: summary.currency,
        referenceDate: '',
        lastChange: '',
        fetchedAt: fetchedAt
      };
    }).filter(Boolean);
  }

  /**
   * Maps open positions to holdings.
   * @param {Array<Object>} positions - `GET /equity/positions`.
   * @param {string} accountId - Account id.
   * @param {string} currency - Account currency (used when a position names none).
   * @param {string} asOf - ISO date-time of the read.
   * @returns {Array<import('../bank-provider.js').Holding>} Holdings with every API field.
   * @throws {Error} DataError if a position has no ticker, quantity or value.
   */
  function holdings(positions, accountId, currency, asOf) {
    return (positions || []).map(function (p) {
      const instrument = p.instrument || {};
      const wallet = p.walletImpact || {};
      if (!instrument.ticker) throw ns.Errors.create('DataError', 'Position is missing its ticker');
      return {
        provider: PROVIDER_ID,
        accountId: accountId,
        ticker: instrument.ticker,
        isin: instrument.isin || '',
        name: instrument.name || '',
        instrumentCurrency: instrument.currency || '',
        quantity: requiredNumber(p.quantity, 'Position quantity'),
        averagePrice: optionalNumber(p.averagePricePaid, 'Average price'),
        currentPrice: optionalNumber(p.currentPrice, 'Current price'),
        amount: requiredNumber(wallet.currentValue, 'Position value'),
        currency: wallet.currency || currency,
        totalCost: optionalNumber(wallet.totalCost, 'Position cost'),
        unrealizedPnl: optionalNumber(wallet.unrealizedProfitLoss, 'Unrealized profit/loss'),
        fxImpact: optionalNumber(wallet.fxImpact, 'FX impact'),
        openedAt: p.createdAt || '',
        asOf: asOf,
        apiFields: ns.Flatten.flat(p)
      };
    });
  }

  /**
   * Shared shape of a booked Trading 212 transaction.
   * @param {Object} fields - accountId, transactionId, date, amount, currency, counterparty, description, raw.
   * @returns {import('../bank-provider.js').Transaction} Transaction (labels are stamped later).
   */
  function booked(fields) {
    return {
      provider: PROVIDER_ID,
      accountId: fields.accountId,
      transactionId: fields.transactionId,
      bookingDate: fields.date,
      valueDate: fields.date,
      status: 'booked',
      amount: fields.amount,
      currency: fields.currency,
      amountEur: null,
      fxRate: null,
      fxRateDate: '',
      counterparty: fields.counterparty,
      description: fields.description,
      apiFields: ns.Flatten.flat(fields.raw)
    };
  }

  /**
   * Local calendar date of an API timestamp.
   * @param {string} dateTime - ISO date-time.
   * @param {function(string): string} toLocalDate - ISO date-time to YYYY-MM-DD in the Bank Sync time zone.
   * @returns {string} YYYY-MM-DD.
   * @throws {Error} DataError if the timestamp is missing or invalid.
   */
  function localDate(dateTime, toLocalDate) {
    if (!dateTime || isNaN(Date.parse(dateTime))) throw ns.Errors.create('DataError', 'Item has no valid date');
    return toLocalDate(dateTime);
  }

  /**
   * Id of an item that has no reference (not expected; keeps the upsert stable anyway).
   * @param {Array<*>} parts - Identifying fields.
   * @returns {string} 'h:' plus a content hash.
   */
  function fallbackId(parts) {
    return 'h:' + ns.Hash.fnv1a(parts.join('|'));
  }

  /**
   * Maps a cash movement (`GET /equity/history/transactions` item) to a transaction. Withdrawals and
   * fees are negative, deposits and interest positive; TRANSFER and unknown types keep the API's sign.
   * @param {Object} item - {type, amount, currency, dateTime, reference}.
   * @param {string} accountId - Account id.
   * @param {function(string): string} toLocalDate - Date-time to local date.
   * @returns {import('../bank-provider.js').Transaction} Booked transaction; `description` is the type.
   * @throws {Error} DataError if amount or date is unusable.
   */
  function cashMovement(item, accountId, toLocalDate) {
    const raw = requiredNumber(item.amount, 'Transaction amount');
    const type = String(item.type || '');
    const amount = OUTFLOWS.indexOf(type) !== -1 ? -Math.abs(raw) : INFLOWS.indexOf(type) !== -1 ? Math.abs(raw) : raw;
    return booked({
      accountId: accountId,
      transactionId: item.reference ? String(item.reference) : fallbackId([type, item.dateTime, raw]),
      date: localDate(item.dateTime, toLocalDate),
      amount: amount,
      currency: item.currency || '',
      counterparty: '',
      description: type,
      raw: item
    });
  }

  /**
   * Maps a paid dividend (`GET /equity/history/dividends` item) to a transaction. Its id is prefixed
   * so it can never collide with a cash movement's reference.
   * @param {Object} item - {amount (account currency), currency, paidOn, reference, ticker, instrument, type, ...}.
   * @param {string} accountId - Account id.
   * @param {function(string): string} toLocalDate - Date-time to local date.
   * @returns {import('../bank-provider.js').Transaction} Booked, positive transaction.
   * @throws {Error} DataError if amount or date is unusable.
   */
  function dividend(item, accountId, toLocalDate) {
    const amount = Math.abs(requiredNumber(item.amount, 'Dividend amount'));
    const instrument = item.instrument || {};
    return booked({
      accountId: accountId,
      transactionId: 'div:' + (item.reference ? String(item.reference) : fallbackId([item.ticker, item.paidOn, amount])),
      date: localDate(item.paidOn, toLocalDate),
      amount: amount,
      currency: item.currency || '',
      counterparty: instrument.name || item.ticker || '',
      description: 'Dividend' + (item.type ? ' ' + item.type : ''),
      raw: item
    });
  }

  ns.T212Normalize = {
    account: account, balances: balances, holdings: holdings, cashMovement: cashMovement, dividend: dividend
  };
})(globalThis.BankSync);
