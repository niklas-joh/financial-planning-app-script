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
  /** BIL's labels, longest first where one starts another ("Beneficiary account" before "Beneficiary"). */
  const LABELS = ['Remittance Info', 'Beneficiary account', 'Beneficiary', 'At', 'By order of', 'Paying bank',
    'End-to-end Identification', 'Payment initiation reference'];
  const CARD_PAYMENT = 'Card payment';

  /**
   * Splits a BIL message on its labels.
   * @param {string} message - The `Message` column, squeezed.
   * @returns {{lead: string, fields: Object<string, string>}} The text before the first label (the kind
   *   of movement, e.g. 'Standing order'; the whole message when there is no label) and each label's
   *   text (the first occurrence of a label wins).
   */
  function parts(message) {
    const re = new RegExp('(^|\\s)(' + LABELS.join('|') + '):\\s*', 'g');
    const marks = [];
    let m;
    while ((m = re.exec(message)) !== null) {
      marks.push({ label: m[2], start: m.index + m[1].length, end: re.lastIndex });
    }
    const fields = {};
    marks.forEach(function (mark, i) {
      const value = message.slice(mark.end, i + 1 < marks.length ? marks[i + 1].start : message.length).trim();
      if (!(mark.label in fields)) fields[mark.label] = value;
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
      return { counterparty: merchant(message), description: message, category: S.code(name),
        counterpartyAccount: '', counterpartyBank: '', reference: '', extra: {} };
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
      return Object.assign({ date: date, valueDate: valueDate, amount: amount, fingerprint: [date, valueDate, name, message, amount] },
        describe(name, message));
    }
  };

  ns.CsvBil = { FORMAT: FORMAT, parts: parts, describe: describe };
})(globalThis.BankSync);
