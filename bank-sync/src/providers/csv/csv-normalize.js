/**
 * @fileoverview Pure: one CSV file to the normalized accounts and transactions, or (for a details
 * format) to details for rows another provider stores. Finds the file's
 * format by its header, masks account and card numbers in every text that is stored (when
 * `MASK_ACCOUNT_NUMBERS` is on; see `Redact.maskAccount`), and gives
 * each row a stable id: the file's own id when it has one, else the date, a content hash and an
 * occurrence counter (so identical rows on the same day stay distinct). Exports cover whole days, so the same
 * or an overlapping file always yields the same ids and re-importing it changes nothing.
 */
(function (ns) {
  const IBAN_IN_TEXT = /\b[A-Z]{2}\d{2}[A-Z0-9]{11,30}\b/g;
  const LONG_NUMBER = /\b\d{12,19}\b/g;

  /**
   * Passes IBANs and long digit runs (card, account and personal id numbers) inside free text
   * through `Redact.maskAccount`: cut to their last four characters when masking is on, else unchanged.
   * @param {string} text - Text.
   * @returns {string} Text, masked when masking is on.
   */
  function maskNumbers(text) {
    return String(text).replace(IBAN_IN_TEXT, ns.Redact.maskAccount).replace(LONG_NUMBER, ns.Redact.maskAccount);
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
   * @returns {Object<string,*>} The record (and parts) as flat `apiFields`, trimmed and masked (when masking is on).
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
   *   `date` (for the last-read record); texts masked when masking is on.
   * @throws {Error} DataError naming the record that cannot be read.
   */
  function details(found) {
    return rowsOf(found).filter(Boolean).map(function (d) {
      return {
        provider: found.format.enriches,
        transactionId: d.id,
        counterparty: maskNumbers(d.counterparty),
        category: d.category,
        classification: d.classification || '',
        date: d.date
      };
    });
  }

  /**
   * @param {Array<string>} values - Values (at least one).
   * @returns {string} The value that occurs most often; on a tie, the one seen first.
   */
  function mostCommon(values) {
    const counts = {};
    values.forEach(function (v) { counts[v] = (counts[v] || 0) + 1; });
    return values.reduce(function (best, v) { return counts[v] > counts[best] ? v : best; }, values[0]);
  }

  /**
   * The normalized account for one account of a file.
   * @param {import('./csv-formats.js').CsvFormat} format - The file's format.
   * @param {import('./csv-formats.js').CsvAccountInfo} info - The account as the file names it.
   * @param {import('../../core/config.js').ConnectionConfig} connection - CSV connection (folder).
   * @returns {import('../bank-provider.js').Account} Account; its id is the source plus a hash of the
   *   number, or of the name when the export has no number.
   */
  function accountOf(format, info, connection) {
    return {
      provider: ns.Config.CSV.PROVIDER_ID,
      accountId: format.id + '-' + ns.Hash.fnv1a(info.number || info.name),
      bank: format.bankName,
      connection: connection.key,
      holderName: '',
      accountName: info.name,
      currency: info.currency,
      ibanMasked: ns.Redact.maskAccount(info.number),
      accountType: '',
      product: info.name,
      detailsAt: ''
    };
  }

  /**
   * Turns a detected file into its accounts and their transactions. The account comes from the lines
   * above the header (one per file, e.g. BIL) or, for a format without them, from each row (several per
   * file, e.g. Avanza).
   * @param {{format: import('./csv-formats.js').CsvFormat, parsed: Object}} found - From `detect`.
   * @param {import('../../core/config.js').ConnectionConfig} connection - CSV connection (folder).
   * @returns {{accounts: Array<import('../bank-provider.js').Account>,
   *   transactions: Array<import('../bank-provider.js').Transaction>}} Normalized rows, with `apiFields`.
   * @throws {Error} DataError naming the row (never its content) that cannot be read.
   */
  function file(found, connection) {
    const format = found.format;
    const fileInfo = format.account ? format.account(found.parsed.preamble) : null;
    const accounts = {};
    const add = function (info) {
      const account = accountOf(format, info, connection);
      return accounts[account.accountId] || (accounts[account.accountId] = account);
    };
    if (fileInfo) add(fileInfo);
    const seen = {};
    const records = found.parsed.records;
    const transactions = rowsOf(found).map(function (row, i) {
      const account = add(fileInfo || row.account);
      let id = row.id;
      if (!id) {
        const print = JSON.stringify(row.fingerprint);
        // Identical rows are counted per account (one account per file: the same as per file).
        const occurrence = account.accountId + '|' + print;
        seen[occurrence] = (seen[occurrence] || 0) + 1;
        // The date in the id keeps a 32-bit hash collision to rows of the same day.
        id = 'h' + row.date.replace(/-/g, '') + '-' + ns.Hash.fnv1a(print) + '-' + seen[occurrence];
      }
      return {
        provider: account.provider,
        accountId: account.accountId,
        transactionId: String(id),
        bookingDate: row.date,
        valueDate: row.valueDate,
        status: 'booked',
        amount: row.amount,
        currency: row.currency || account.currency,
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
    if (!fileInfo) {
      // An account named per row takes the currency most of its rows have, so one foreign-currency row
      // (e.g. a withholding tax) does not label the whole account; rows keep their own currency.
      Object.keys(accounts).forEach(function (accountId) {
        accounts[accountId].currency = mostCommon(transactions
          .filter(function (t) { return t.accountId === accountId; })
          .map(function (t) { return t.currency; }));
      });
    }
    return {
      accounts: Object.keys(accounts).map(function (id) { return accounts[id]; }),
      transactions: transactions
    };
  }

  ns.CsvNormalize = { detect: detect, file: file, details: details, maskNumbers: maskNumbers };
})(globalThis.BankSync);
