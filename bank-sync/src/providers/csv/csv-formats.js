/**
 * @fileoverview The CSV exports Bank Sync can read, one entry per source, each written from a
 * real export's header and layout (never guessed). A format names the columns its header must
 * contain, how its numbers and dates are written, and either where the account comes from and how
 * a record becomes a transaction (`kind: 'transactions'`), or how a record adds details to rows
 * another provider already stores (`kind: 'details'`). Adding a source is adding an entry here;
 * files matching no entry are skipped and reported.
 */
(function (ns) {
  /**
   * @typedef {Object} CsvRow
   * @property {string} date - Booking date (ISO).
   * @property {string} valueDate - Value date (ISO), or ''.
   * @property {number} amount - Signed: money out is negative.
   * @property {string} counterparty - Other party, or '' (account numbers are masked later).
   * @property {string} description - Text for the transactions tab (account numbers are masked later).
   * @property {string} [id] - The file's own transaction id, when it has one.
   * @property {Array<*>} fingerprint - Identifying values for a content-hash id (when `id` is empty).
   */

  /**
   * @typedef {Object} CsvDetail
   * @property {string} id - Transaction id of the row the detail is for (as the other provider stores it).
   * @property {string} date - Date of the record (ISO), for the last-read record only.
   * @property {string} counterparty - Other party, or ''.
   * @property {string} category - Category, or ''.
   */

  /**
   * @typedef {Object} CsvFormat
   * @property {string} id - Source id, e.g. 'bil'; also the account id prefix.
   * @property {string} kind - 'transactions' (rows of its own) or 'details' (details for another provider's rows).
   * @property {string} [enriches] - For 'details': the provider whose rows the details are for.
   * @property {string} bankName - Bank label on the tabs.
   * @property {string} separator - Field separator.
   * @property {Array<string>} encodings - Charsets to try, in order (a UTF-8 decode with invalid bytes is skipped).
   * @property {Array<string>} columns - Column names the header row must contain.
   * @property {function(Array<Array<string>>): {number: string, name: string, currency: string}} [account] -
   *   For 'transactions': the account the file belongs to, from the lines above the header.
   * @property {function(Object<string,string>): ?(CsvRow|CsvDetail)} row - One record to a row ('transactions') or
   *   a detail, or null when the record carries none ('details').
   */

  /** Amounts look like `-824.50`, `8,828.03` or `-1,500`. */
  const BIL_NUMBER = { decimal: '.', thousands: ',' };
  const BIL_DATE = 'DD/MM/YYYY';

  /**
   * @param {string} text - Text.
   * @returns {string} Trimmed, with runs of spaces collapsed.
   */
  function squeeze(text) {
    return String(text).replace(/\s+/g, ' ').trim();
  }

  /**
   * Other party of a BIL movement, from the message: the beneficiary of a payment, who ordered a
   * credit, or the merchant of a card payment (the text before its date).
   * @param {string} name - The movement type (`Name` column), e.g. 'Card payment'.
   * @param {string} message - The `Message` column, squeezed.
   * @returns {string} Counterparty, or ''.
   */
  function bilCounterparty(name, message) {
    const marked = /Beneficiary: (.+?)(?: Beneficiary account:| At:|$)/.exec(message) ||
      /By order of: (.+?)(?: Paying bank:| End-to-end|$)/.exec(message);
    if (marked) return marked[1].trim();
    if (name === 'Card payment') return message.split(/ \d{2}\/\d{2}\/\d{4}/)[0].trim();
    return '';
  }

  /** @type {CsvFormat} BILnet "List of account movements" export. */
  const BIL = {
    id: 'bil',
    kind: 'transactions',
    bankName: 'BIL',
    separator: ';',
    encodings: ['UTF-8', 'windows-1252'],
    columns: ['Date', 'Value date', 'Name', 'Message', 'Amount'],
    account: function (preamble) {
      const line = preamble.map(function (row) { return squeeze(row.join(' ')); })
        .map(function (text) { return /^([A-Z]{2}\d{2}[A-Z0-9]{11,30}) ?(.*)$/.exec(text); })
        .filter(Boolean)[0];
      if (!line) throw ns.Errors.create('DataError', 'BIL file has no account line above the header');
      // The export has no currency column; the owner's BIL accounts are EUR (confirmed 2026-10-25).
      return { number: line[1], name: line[2], currency: 'EUR' };
    },
    row: function (r) {
      const message = squeeze(r.Message);
      const name = squeeze(r.Name);
      const date = ns.CsvParse.date(r.Date, BIL_DATE);
      const valueDate = r['Value date'].trim() ? ns.CsvParse.date(r['Value date'], BIL_DATE) : '';
      const amount = ns.CsvParse.number(r.Amount, BIL_NUMBER);
      if (amount === null) throw ns.Errors.create('DataError', 'A BIL row has no amount');
      return {
        date: date,
        valueDate: valueDate,
        amount: amount,
        counterparty: bilCounterparty(name, message),
        description: name + (message ? ': ' + message : ''),
        fingerprint: [date, valueDate, name, message, amount]
      };
    }
  };

  /**
   * Trading 212 history report (app export or `POST /equity/history/exports`). Its `ID` is the API's
   * `reference`, so it names rows the Trading 212 provider already stores, and every record gives its
   * row what the API lacks: `category` = the merchant category, else the action as a code (`CARD_DEBIT`,
   * `SPENDING_CASHBACK`, `DEPOSIT`, ...), so a card payment is told apart from a withdrawal and cashback
   * from a deposit; `counterparty` = the merchant name, else the notes (e.g. `Bank Transfer`). Columns
   * that are empty for the exported period may be missing (merchant columns without card payments), so
   * only the ones every report has are required.
   * @type {CsvFormat}
   */
  const T212_REPORT = {
    id: 'trading212-report',
    kind: 'details',
    enriches: ns.Config.TRADING212.PROVIDER_ID,
    bankName: ns.Config.TRADING212.BANK_NAME,
    separator: ',',
    encodings: ['UTF-8'],
    columns: ['Action', 'Time (UTC)', 'ID', 'Total'],
    row: function (r) {
      const id = squeeze(r.ID);
      if (!id) throw ns.Errors.create('DataError', 'A Trading 212 report row has no ID');
      // Time is 'YYYY-MM-DD HH:MM:SS+00:00' (UTC); the date only feeds the last-read record.
      const date = ns.CsvParse.date(String(r['Time (UTC)']).trim().slice(0, 10), 'YYYY-MM-DD');
      return {
        id: id,
        date: date,
        counterparty: squeeze(r['Merchant name'] || '') || squeeze(r.Notes || ''),
        category: squeeze(r['Merchant category'] || '') || code(r.Action)
      };
    }
  };

  /**
   * @param {string} text - Label, e.g. 'Spending cashback'.
   * @returns {string} Upper-case code, e.g. 'SPENDING_CASHBACK'.
   */
  function code(text) {
    return String(text || '').trim().toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  }

  const ALL = [BIL, T212_REPORT];

  ns.CsvFormats = { ALL: ALL, bilCounterparty: bilCounterparty };
})(globalThis.BankSync);
