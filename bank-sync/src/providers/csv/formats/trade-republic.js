/**
 * @fileoverview Trade Republic timeline export, written by this repo's own job (`bank-sync/trade-republic/`,
 * which reads the timeline with the community library pytr), so the layout is ours and versioned with the
 * code: UTF-8, `;`, header `TR event id;Timestamp;Date;Account;Event type;Type;Title;Subtitle;Amount;
 * Currency;ISIN;Shares;Fees;Taxes;Classification`; dates `YYYY-MM-DD` (Europe/Brussels day); numbers with a
 * `.` decimal point and no thousands separator; empty cells blank. One file holds the whole history and is
 * overwritten by every run, so re-reading it changes nothing: each row carries Trade Republic's own event id.
 * `Amount` is signed from the cash side (a card payment or a purchase -, a deposit, sale or dividend +).
 * `Classification` is set by the job from the event type: 'card' and 'transfer' rows go to `transactions`,
 * 'investment' rows to `investment_transactions`; '' (a type the job does not know) goes to `transactions`
 * so it is seen.
 */
(function (ns) {
  const NUMBER = { decimal: '.', thousands: '' };
  const DATE = 'YYYY-MM-DD';
  const CLASSIFICATIONS = ['card', 'transfer', 'investment', ''];

  /**
   * @param {string} text - Cell text.
   * @param {string} what - Column, for the error.
   * @returns {string} The text, squeezed.
   * @throws {Error} DataError if it is empty.
   */
  function required(text, what) {
    const value = ns.CsvParse.squeeze(text);
    if (!value) throw ns.Errors.create('DataError', 'A Trade Republic row has no ' + what);
    return value;
  }

  /** @type {import('../csv-formats.js').CsvFormat} */
  const FORMAT = {
    id: 'trade-republic',
    kind: 'transactions',
    bankName: 'Trade Republic',
    separator: ';',
    encodings: ['UTF-8'],
    columns: ['TR event id', 'Timestamp', 'Date', 'Account', 'Event type', 'Type', 'Title', 'Subtitle', 'Amount',
      'Currency', 'ISIN', 'Shares', 'Fees', 'Taxes', 'Classification'],
    row: function (r) {
      const S = ns.CsvParse;
      const id = required(r['TR event id'], 'event id');
      const account = required(r.Account, 'account');
      const eventType = required(r['Event type'], 'event type');
      const currency = required(r.Currency, 'currency').toUpperCase();
      if (!/^[A-Z]{3}$/.test(currency)) throw ns.Errors.create('DataError', 'A Trade Republic row has no currency code');
      const amount = S.number(r.Amount, NUMBER);
      if (amount === null) throw ns.Errors.create('DataError', 'A Trade Republic row has no amount');
      const classification = S.squeeze(r.Classification).toLowerCase();
      if (CLASSIFICATIONS.indexOf(classification) === -1) {
        throw ns.Errors.create('DataError', 'A Trade Republic row has an unknown classification');
      }
      // Read so that a malformed number skips the file; the values stay on the raw tab.
      S.number(r.Shares, NUMBER);
      S.number(r.Fees, NUMBER);
      S.number(r.Taxes, NUMBER);
      const subtitle = S.squeeze(r.Subtitle);
      return {
        id: id,
        account: { number: '', name: account, currency: currency },
        date: S.date(r.Date, DATE),
        valueDate: '',
        amount: amount,
        currency: currency,
        // Merchant (card), sender or recipient (transfer), or security (trade, dividend).
        counterparty: S.squeeze(r.Title),
        description: subtitle || S.squeeze(r.Type) || eventType,
        category: S.code(eventType),
        reference: S.squeeze(r.ISIN),
        classification: classification,
        fingerprint: []
      };
    }
  };

  ns.CsvTradeRepublic = { FORMAT: FORMAT };
})(globalThis.BankSync);
