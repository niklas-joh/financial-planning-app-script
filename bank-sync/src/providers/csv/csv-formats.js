/**
 * @fileoverview The CSV exports Bank Sync can read. Each source has its own file under `formats/`,
 * written from a real export's header and layout (never guessed), so bank-specific logic never
 * mixes; this file holds the shared shapes and the list. A format names the columns its header must
 * contain, how its numbers and dates are written, and either where the account comes from (the lines
 * above the header, or each row when one file holds several accounts) and how a record becomes a
 * transaction (`kind: 'transactions'`), or how a record adds details to rows another
 * provider already stores (`kind: 'details'`). Every format writes to the same tabs, so totals stay
 * one sum. Adding a source is a new file plus an entry in `ALL`; files matching no entry are skipped
 * and reported.
 */
(function (ns) {
  /**
   * @typedef {Object} CsvAccountInfo
   * @property {string} number - Account number (hashed into the account id as found, so the id does not depend on masking; on the
   *   tabs through `Redact.maskAccount`), or '' when the
   *   export has none.
   * @property {string} name - Account name; hashed into the account id when there is no number.
   * @property {string} currency - Account currency.
   */

  /**
   * @typedef {Object} CsvRow
   * @property {string} date - Booking date (ISO).
   * @property {string} valueDate - Value date (ISO), or ''.
   * @property {number} amount - Signed: money out is negative.
   * @property {string} [currency] - The row's currency, when the file states it per row (else the account's).
   * @property {CsvAccountInfo} [account] - The row's account, for a format without a file-level `account`.
   * @property {string} counterparty - Other party, or '' (account numbers are masked later when masking is on).
   * @property {string} description - Text for the transactions tab (account numbers are masked later when masking is on).
   * @property {string} [category] - Kind of movement or spending category, or ''.
   * @property {string} [counterpartyAccount] - The other party's account (masked later when masking is on), or ''.
   * @property {string} [counterpartyBank] - The other party's bank, or ''.
   * @property {string} [reference] - Payment reference (e.g. end-to-end id), or ''.
   * @property {Object<string, string>} [extra] - Further parsed parts, kept on the raw tab only.
   * @property {string} [id] - The file's own transaction id, when it has one.
   * @property {Array<*>} fingerprint - Identifying values for a content-hash id (when `id` is empty).
   */

  /**
   * @typedef {Object} CsvDetail
   * @property {string} id - Transaction id of the row the detail is for (as the other provider stores it).
   * @property {string} date - Date of the record (ISO), for the last-read record only.
   * @property {string} counterparty - Other party, or ''.
   * @property {string} category - Category, or ''.
   * @property {string} [classification] - 'card', 'transfer' or 'investment' (which tab the row belongs on), or ''.
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
   * @property {function(Array<Array<string>>): CsvAccountInfo} [account] - For 'transactions': the account the
   *   file belongs to, from the lines above the header. Without it, every row names its own `account`.
   * @property {function(Object<string,string>): ?(CsvRow|CsvDetail)} row - One record to a row ('transactions') or
   *   a detail, or null when the record carries none ('details').
   */

  /** @type {Array<CsvFormat>} */
  const ALL = [ns.CsvBil.FORMAT, ns.CsvT212Report.FORMAT, ns.CsvAvanza.FORMAT];

  ns.CsvFormats = { ALL: ALL };
})(globalThis.BankSync);
