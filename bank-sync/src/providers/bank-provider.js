/**
 * @fileoverview Common provider contract and normalized data models.
 * Providers are independent: adding one never requires editing another.
 */
(function (ns) {
  /**
   * @typedef {Object} Account
   * @property {string} provider - Provider id, e.g. 'enablebanking'.
   * @property {string} accountId - Provider-scoped account identifier.
   * @property {string} bank - Bank name, e.g. 'Revolut'. Set by the provider.
   * @property {string} connection - Connection key (bank plus holder). Set by the provider.
   * @property {string} holderName - Account holder's name as the API reports it (Enable Banking `name`).
   * @property {string} accountName - Name set by the user or bank (Enable Banking `details`), or ''.
   * @property {string} currency - ISO 4217 code.
   * @property {string} ibanMasked - IBAN reduced to last four characters, or ''.
   * @property {string} accountType - Provider account type (e.g. ISO cash account type).
   * @property {string} product - Provider product name, or ''.
   * @property {string} detailsAt - ISO date-time the full details were last read, or ''.
   * @property {Object<string,*>} [apiFields] - Every API field, flattened and IBAN-masked (raw tab).
   */

  /**
   * @typedef {Object} Balance
   * @property {string} provider - Provider id.
   * @property {string} accountId - Account the balance belongs to.
   * @property {string} balanceType - Provider balance type code (e.g. CLBD).
   * @property {string} description - What the balance type means (Enable Banking balance `name`).
   * @property {number} amount - Amount in the account currency.
   * @property {string} currency - ISO 4217 code.
   * @property {string} referenceDate - ISO date, or ''.
   * @property {string} lastChange - ISO date-time, or ''.
   * @property {string} fetchedAt - ISO date-time the balance was read.
   * @property {string} [valueKind] - 'cash' or 'investments' when the balance is part of the account's
   *   daily value (see `core/daily-values.js`); '' or missing otherwise.
   * @property {Object<string,*>} [apiFields] - Every API field, flattened (raw tab).
   */

  /**
   * @typedef {Object} Transaction
   * @property {string} provider
   * @property {string} bank - Label, filled from the account list at sync time.
   * @property {string} connection - Label, as above.
   * @property {string} holderName - Label, as above.
   * @property {string} accountName - Label, as above.
   * @property {string} ibanMasked - Label, as above.
   * @property {string} accountId
   * @property {string} transactionId
   * @property {string} bookingDate
   * @property {string} valueDate
   * @property {string} status - 'pending' or 'booked'.
   * @property {number} amount
   * @property {string} currency
   * @property {?number} amountEur
   * @property {?number} fxRate
   * @property {string} fxRateDate
   * @property {string} counterparty
   * @property {string} description
   * @property {string} [category] - Kind of movement or spending category (e.g. a card's merchant category), or ''.
   * @property {string} [counterpartyAccount] - The other party's account, masked to its last four characters, or ''.
   * @property {string} [counterpartyBank] - The other party's bank, or ''.
   * @property {string} [reference] - Payment reference (e.g. end-to-end id), masked like account numbers, or ''.
   * @property {Object<string,*>} [apiFields] - Every API field, flattened (raw tab only; never stored on `transactions`).
   */

  /**
   * @typedef {Object} Holding
   * @property {string} provider - Provider id.
   * @property {string} accountId - Account the position is held in.
   * @property {string} ticker - Provider's instrument id (e.g. 'AAPL_US_EQ').
   * @property {string} isin - ISIN, or ''.
   * @property {string} name - Instrument name.
   * @property {string} instrumentCurrency - Currency the instrument trades in (prices are in it).
   * @property {number} quantity - Shares held.
   * @property {number} averagePrice - Average price paid per share, instrument currency.
   * @property {number} currentPrice - Current price per share, instrument currency.
   * @property {number} amount - Current market value in `currency` (the account currency); named like a
   *   balance amount so the same EUR conversion applies.
   * @property {string} currency - Currency of `amount`, `totalCost`, `unrealizedPnl` and `fxImpact`.
   * @property {?number} totalCost - What the position cost.
   * @property {?number} unrealizedPnl - `amount` minus `totalCost`.
   * @property {?number} fxImpact - Part of the value change caused by exchange rates, or null.
   * @property {string} openedAt - ISO date-time the position was opened, or ''.
   * @property {string} asOf - ISO date-time the position was read.
   * @property {?number} [amountEur] - Set by the FX step.
   * @property {?number} [fxRate] - Set by the FX step.
   * @property {string} [fxRateDate] - Set by the FX step.
   * @property {Object<string,*>} [apiFields] - Every API field, flattened (holdings tab).
   */

  /**
   * @typedef {Object} DailyValue
   * @property {string} date - ISO date of the run (Bank Sync time zone).
   * @property {string} provider
   * @property {string} bank - Label.
   * @property {string} connection - Label.
   * @property {string} holderName - Label.
   * @property {string} accountName - Label.
   * @property {string} accountId
   * @property {string} currency - Account currency.
   * @property {number} cash - Sum of the account's 'cash' balances.
   * @property {number} investments - Sum of its 'investments' balances.
   * @property {number} total - cash plus investments.
   * @property {?number} cashEur
   * @property {?number} investmentsEur
   * @property {?number} totalEur
   * @property {?number} fxRate
   * @property {string} fxRateDate
   * @property {string} fetchedAt - When the balances were read.
   */

  /**
   * @typedef {Object} BankProvider
   * @property {function(Object): Object} authorize - Starts user authorization.
   * @property {function(Object=): Array<Account>} listAccounts - Optional `{outOfTime: function(): boolean}`.
   * @property {function(string): Array<Balance>} getBalances
   * @property {function(string, Object): Array<Transaction>} getTransactions - `(accountId, {dateFrom, dateTo})`.
   * @property {number} [maxWindowDays] - Longest date range one `getTransactions` call takes (default
   *   `Config.TRANSACTIONS.WINDOW_DAYS`); Infinity for cursor-paged APIs that read the whole range at once.
   * @property {function(string): Array<Holding>} [getHoldings] - Open positions of an account, for providers
   *   that hold investments. Optional: banks do not implement it.
   * @property {function(): Array<import('../core/txn-merge.js').Enrichment>} [getEnrichments] - Details for rows
   *   another provider stores (e.g. card merchants from a CSV report). Optional.
   */

  const REQUIRED_METHODS = ['authorize', 'listAccounts', 'getBalances', 'getTransactions'];

  /**
   * Verifies an object implements the provider contract.
   * @param {Object} candidate - Possible provider.
   * @returns {Object} The same object, if valid.
   * @throws {Error} ConfigError naming the missing methods.
   */
  function assertBankProvider(candidate) {
    const missing = REQUIRED_METHODS.filter(function (name) {
      return !candidate || typeof candidate[name] !== 'function';
    });
    if (missing.length) {
      throw ns.Errors.config('Provider is missing methods: ' + missing.join(', '));
    }
    return candidate;
  }

  ns.BankProvider = { REQUIRED_METHODS: REQUIRED_METHODS, assertBankProvider: assertBankProvider };
})(globalThis.BankSync);
