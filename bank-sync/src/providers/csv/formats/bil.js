/**
 * @fileoverview BIL (BILnet "List of account movements") CSV export, written from the owner's
 * export (2026-10-25). `;`-separated; title lines, then `<IBAN> <product>`, then the header
 * `Date;Value date;Name;Message;Amount`; dates `DD/MM/YYYY`; signed amounts with `.` decimals and
 * `,` thousands. The `Message` is BIL's own text with labelled parts ("Standing order Remittance
 * Info: ... Beneficiary: ... Beneficiary account: ... At: ..."), split here on those labels only, so
 * a colon inside the remittance text (e.g. "DATE CISS: 19.08.2026") stays part of it.
 */
(function (ns) {
  const NUMBER = { decimal: '.', thousands: ',' };
  const DATE = 'DD/MM/YYYY';
  /**
   * BIL's labels in the order they appear, per kind of message: a payment names the beneficiary, a credit
   * who ordered it; the two never mix.
   */
  const ORDERS = [
    ['Remittance Info', 'Beneficiary', 'Beneficiary account', 'At', 'End-to-end Identification', 'Payment initiation reference'],
    ['Remittance Info', 'By order of', 'Paying bank', 'End-to-end Identification', 'Payment initiation reference']
  ];
  /** Every label, longest first where one starts another ("Beneficiary account" before "Beneficiary"), for the regex. */
  const LABELS = ORDERS[0].concat(ORDERS[1]).filter(function (l, i, all) { return all.indexOf(l) === i; })
    .sort(function (a, b) { return b.length - a.length; });
  const CARD_PAYMENT = 'Card payment';
  /** Columns a card payment has no value for. */
  const NO_PARTIES = Object.freeze({ counterpartyAccount: '', counterpartyBank: '', reference: '', extra: Object.freeze({}) });

  /**
   * The label marks that form a valid sequence in one order, read from the end: a mark counts only if its
   * label comes earlier in the order than the label after it.
   * @param {Array<{label: string, start: number, end: number}>} found - Every label-looking mark, in text order.
   * @param {Array<string>} order - One of `ORDERS`.
   * @returns {Array<{label: string, start: number, end: number}>} The accepted marks, in text order.
   */
  function chain(found, order) {
    const marks = [];
    let next = order.length;
    for (let i = found.length - 1; i >= 0; i--) {
      const rank = order.indexOf(found[i].label);
      if (rank !== -1 && rank < next) {
        marks.unshift(found[i]);
        next = rank;
      }
    }
    return marks;
  }

  /**
   * Splits a BIL message on its labels. Only the remittance is free text, and it comes first, so a label
   * typed inside it (e.g. "Meet At: the club") sits before the real labels. The marks are matched from the
   * end against each kind's label order (`chain`), and the kind that accepts the most wins; a mark that is
   * not accepted stays part of the text it is in.
   * @param {string} message - The `Message` column, squeezed.
   * @returns {{lead: string, fields: Object<string, string>}} The text before the first label (the kind
   *   of movement, e.g. 'Standing order'; the whole message when there is no label) and each label's text.
   */
  function parts(message) {
    const re = new RegExp('(^|\\s)(' + LABELS.join('|') + '):\\s*', 'g');
    const found = [];
    let m;
    while ((m = re.exec(message)) !== null) {
      found.push({ label: m[2], start: m.index + m[1].length, end: re.lastIndex });
    }
    const marks = ORDERS.map(function (order) { return chain(found, order); })
      .reduce(function (best, c) { return c.length > best.length ? c : best; });
    const fields = {};
    marks.forEach(function (mark, i) {
      fields[mark.label] = message.slice(mark.end, i + 1 < marks.length ? marks[i + 1].start : message.length).trim();
    });
    return { lead: message.slice(0, marks.length ? marks[0].start : message.length).trim(), fields: fields };
  }

  /**
   * @param {string} message - A card payment's message: merchant, country, purchase date and time, card.
   * @returns {string} The merchant text before the purchase date.
   */
  function merchant(message) {
    return message.split(/ \d{2}\/\d{2}\/\d{4}/)[0].trim();
  }

  /**
   * Maps a parsed message to the transaction columns. Card payments: counterparty = merchant, the whole
   * message as description. Others: description = the remittance text, counterparty = the beneficiary or
   * the payer, their account and bank, and the end-to-end id (else the payment reference) as reference;
   * category = the kind of movement as a code (STANDING_ORDER, INSTANT_PAYMENT, CREDIT, ...).
   * @param {string} name - The `Name` column (BIL's movement type, e.g. 'Payment order').
   * @param {string} message - The `Message` column, squeezed.
   * @returns {{counterparty: string, description: string, category: string, counterpartyAccount: string,
   *   counterpartyBank: string, reference: string, extra: Object<string, string>}} Columns, and every
   *   labelled part for the raw tab.
   */
  function describe(name, message) {
    const S = ns.CsvParse;
    if (name === CARD_PAYMENT) {
      return Object.assign({ counterparty: merchant(message), description: message, category: S.code(name) }, NO_PARTIES);
    }
    const p = parts(message);
    const f = p.fields;
    const labelled = Object.keys(f).length > 0;
    return {
      counterparty: f.Beneficiary || f['By order of'] || '',
      description: f['Remittance Info'] || (labelled ? p.lead : message) || name,
      category: S.code(labelled && p.lead ? p.lead : name),
      counterpartyAccount: f['Beneficiary account'] || '',
      counterpartyBank: f.At || f['Paying bank'] || '',
      reference: f['End-to-end Identification'] || f['Payment initiation reference'] || '',
      extra: f
    };
  }

  /** @type {import('../csv-formats.js').CsvFormat} */
  const FORMAT = {
    id: 'bil',
    kind: 'transactions',
    bankName: 'BIL',
    separator: ';',
    encodings: ['UTF-8', 'windows-1252'],
    columns: ['Date', 'Value date', 'Name', 'Message', 'Amount'],
    account: function (preamble) {
      const line = preamble.map(function (row) { return ns.CsvParse.squeeze(row.join(' ')); })
        .map(function (text) { return /^([A-Z]{2}\d{2}[A-Z0-9]{11,30}) ?(.*)$/.exec(text); })
        .filter(Boolean)[0];
      if (!line) throw ns.Errors.create('DataError', 'BIL file has no account line above the header');
      // The export has no currency column; the owner's BIL accounts are EUR (confirmed 2026-10-25).
      return { number: line[1], name: line[2], currency: 'EUR' };
    },
    row: function (r) {
      const S = ns.CsvParse;
      const message = S.squeeze(r.Message);
      const name = S.squeeze(r.Name);
      const date = S.date(r.Date, DATE);
      const valueDate = S.squeeze(r['Value date']) ? S.date(r['Value date'], DATE) : '';
      const amount = S.number(r.Amount, NUMBER);
      if (amount === null) throw ns.Errors.create('DataError', 'A BIL row has no amount');
      // The fingerprint (and so the id) uses the raw columns only, so how the message is split can change
      // without changing ids.
      const row = describe(name, message);
      row.date = date;
      row.valueDate = valueDate;
      row.amount = amount;
      row.fingerprint = [date, valueDate, name, message, amount];
      return row;
    }
  };

  // `parts` and `describe` are exported for the tests only; the import uses `FORMAT`.
  ns.CsvBil = { FORMAT: FORMAT, parts: parts, describe: describe };
})(globalThis.BankSync);
