/**
 * @fileoverview Pure: one CSV file to the normalized account and transactions, or (for a details
 * format) to details for rows another provider stores. Finds the file's
 * format by its header, masks account and card numbers in every text that is stored, and gives
 * each row a stable id: the file's own id when it has one, else the date, a content hash and an
 * occurrence counter (so identical rows on the same day stay distinct). Exports cover whole days, so the same
 * or an overlapping file always yields the same ids and re-importing it changes nothing.
 */
(function (ns) {
  const IBAN_IN_TEXT = /\b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b/g;
  const LONG_NUMBER = /\b\d{12,19}\b/g;

  /**
   * Masks IBANs and long digit runs (card, account and personal id numbers) inside free text
   * down to their last four characters.
   * @param {string} text - Text.
   * @returns {string} Masked text.
   */
  function maskNumbers(text) {
    return String(text).replace(IBAN_IN_TEXT, ns.Redact.maskIban).replace(LONG_NUMBER, ns.Redact.maskIban);
  }

  /**
   * Finds the format of a file by trying each format's charsets and header.
   * @param {function(string): string} decode - The file's content in a given charset.
   * @param {Array<import('./csv-formats.js').CsvFormat>} [formats] - Candidates (default all).
   * @returns {?{format: import('./csv-formats.js').CsvFormat, parsed: Object}} The format and the
   *   parsed table, or null if no format's header is in the file.
   * @throws {Error} DataError from parsing a file whose header matched.
   */
  function detect(decode, formats) {
    const tried = {};
    const candidates = formats || ns.CsvFormats.ALL;
    for (let i = 0; i < candidates.length; i++) {
      const format = candidates[i];
      for (let j = 0; j < format.encodings.length; j++) {
        const charset = format.encodings[j];
        const text = tried[charset] === undefined ? (tried[charset] = decode(charset)) : tried[charset];
        if (/^utf-?8$/i.test(charset) && text.indexOf('\uFFFD') !== -1) continue;
        const parsed = ns.CsvParse.table(ns.CsvParse.rows(text, format.separator), format.columns);
        if (parsed) return { format: format, parsed: parsed };
      }
    }
    return null;
  }

  /**
   * @param {Object<string,string>} record - CSV record.
   * @param {Object<string,string>} [extra] - Parts the format parsed out of it, added as `parsed.<name>`.
   * @returns {Object<string,*>} The record (and parts) as flat `apiFields`, trimmed and masked.
   */
  function apiFields(record, extra) {
    const masked = {};
    Object.keys(record).forEach(function (k) { masked[k] = maskNumbers(String(record[k]).trim()); });
    Object.keys(extra || {}).forEach(function (k) { masked['parsed.' + k] = maskNumbers(String(extra[k]).trim()); });
    return ns.Flatten.flat(masked);
  }

  /**
   * Runs the format over every record.
   * @param {{format: import('./csv-formats.js').CsvFormat, parsed: Object}} found - From `detect`.
   * @returns {Array<?Object>} The format's result per record (null where it has none).
   * @throws {Error} DataError naming the record (never its content) that cannot be read.
   */
  function rowsOf(found) {
    return found.parsed.records.map(function (record, i) {
      try {
        return found.format.row(record);
      } catch (e) {
        throw ns.Errors.create('DataError', 'Record ' + (i + 1) + ': ' + e.message);
      }
    });
  }

  /**
   * Turns a detected details file into details for another provider's rows.
   * @param {{format: import('./csv-formats.js').CsvFormat, parsed: Object}} found - From `detect` (kind 'details').
   * @returns {Array<import('../../core/txn-merge.js').Enrichment>} One per record that carries details, with its
   *   `date` (for the last-read record); texts masked.
   * @throws {Error} DataError naming the record that cannot be read.
   */
  function details(found) {
    return rowsOf(found).filter(Boolean).map(function (d) {
      return {
        provider: found.format.enriches,
        transactionId: d.id,
        counterparty: maskNumbers(d.counterparty),
        category: d.category,
        kind: d.kind || '',
        date: d.date
      };
    });
  }

  /**
   * Turns a detected file into the account and its transactions.
   * @param {{format: import('./csv-formats.js').CsvFormat, parsed: Object}} found - From `detect`.
   * @param {import('../../core/config.js').ConnectionConfig} connection - CSV connection (folder).
   * @returns {{account: import('../bank-provider.js').Account,
   *   transactions: Array<import('../bank-provider.js').Transaction>}} Normalized rows, with `apiFields`.
   * @throws {Error} DataError naming the row (never its content) that cannot be read.
   */
  function file(found, connection) {
    const format = found.format;
    const info = format.account(found.parsed.preamble);
    const accountId = format.id + '-' + ns.Hash.fnv1a(info.number);
    const account = {
      provider: ns.Config.CSV.PROVIDER_ID,
      accountId: accountId,
      bank: format.bankName,
      connection: connection.key,
      holderName: '',
      accountName: info.name,
      currency: info.currency,
      ibanMasked: ns.Redact.maskIban(info.number),
      accountType: '',
      product: info.name,
      detailsAt: ''
    };
    const seen = {};
    const records = found.parsed.records;
    const transactions = rowsOf(found).map(function (row, i) {
      let id = row.id;
      if (!id) {
        const print = JSON.stringify(row.fingerprint);
        seen[print] = (seen[print] || 0) + 1;
        // The date in the id keeps a 32-bit hash collision to rows of the same day.
        id = 'h' + row.date.replace(/-/g, '') + '-' + ns.Hash.fnv1a(print) + '-' + seen[print];
      }
      return {
        provider: account.provider,
        accountId: accountId,
        transactionId: String(id),
        bookingDate: row.date,
        valueDate: row.valueDate,
        status: 'booked',
        amount: row.amount,
        currency: info.currency,
        amountEur: null,
        fxRate: null,
        fxRateDate: '',
        counterparty: maskNumbers(row.counterparty),
        description: maskNumbers(row.description),
        category: row.category || '',
        counterpartyAccount: maskNumbers(row.counterpartyAccount || ''),
        counterpartyBank: row.counterpartyBank || '',
        reference: maskNumbers(row.reference || ''),
        apiFields: apiFields(records[i], row.extra)
      };
    });
    return { account: account, transactions: transactions };
  }

  ns.CsvNormalize = { detect: detect, file: file, details: details, maskNumbers: maskNumbers };
})(globalThis.BankSync);
