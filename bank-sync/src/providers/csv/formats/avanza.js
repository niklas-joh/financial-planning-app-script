/**
 * @fileoverview Avanza transaction export (website: Min ekonomi > Transaktioner > Exportera transaktioner),
 * written from the owner's export (2026-10-26): UTF-8 with BOM, `;`, LF, header
 * `Datum;Konto;Typ av transaktion;Värdepapper/beskrivning;Antal;Kurs;Belopp;Transaktionsvaluta;Courtage;
 * Valutakurs;Instrumentvaluta;ISIN;Resultat`; dates `YYYY-MM-DD`; decimal comma, no thousands separator;
 * empty cells blank. One file holds every account, so the account comes from each row's `Konto` (the
 * account's name: the export has no account number). `Belopp` is signed from the cash side (a sale or a
 * dividend +, a purchase, fee, tax or reinvested dividend -), so every type, also one not seen yet, is
 * read the same way: amount = `Belopp`, description = the type as written, category = the type as a code.
 * An older export names the fee column `Courtage (SEK)`, so only the columns every version has are
 * required; the rest are kept on the raw tab.
 */
(function (ns) {
  const NUMBER = { decimal: ',', thousands: '' };
  const DATE = 'YYYY-MM-DD';

  /**
   * @param {string} text - Cell text.
   * @param {string} what - Column, for the error.
   * @returns {string} The text, squeezed.
   * @throws {Error} DataError if it is empty.
   */
  function required(text, what) {
    const value = ns.CsvParse.squeeze(text);
    if (!value) throw ns.Errors.create('DataError', 'An Avanza row has no ' + what);
    return value;
  }

  /** @type {import('../csv-formats.js').CsvFormat} */
  const FORMAT = {
    id: 'avanza',
    kind: 'transactions',
    bankName: 'Avanza',
    separator: ';',
    encodings: ['UTF-8'],
    columns: ['Datum', 'Konto', 'Typ av transaktion', 'Värdepapper/beskrivning', 'Antal', 'Kurs', 'Belopp',
      'Transaktionsvaluta', 'ISIN'],
    row: function (r) {
      const S = ns.CsvParse;
      const date = S.date(r.Datum, DATE);
      const account = required(r.Konto, 'account');
      const type = required(r['Typ av transaktion'], 'type');
      const currency = required(r.Transaktionsvaluta, 'currency').toUpperCase();
      if (!/^[A-Z]{3}$/.test(currency)) throw ns.Errors.create('DataError', 'An Avanza row has no currency code');
      const security = S.squeeze(r['Värdepapper/beskrivning']);
      const isin = S.squeeze(r.ISIN);
      // A blank amount would be a movement without cash (e.g. a split); none was in the sample.
      const amount = S.number(r.Belopp, NUMBER) || 0;
      // Read so that a malformed number skips the file; the values stay on the raw tab.
      const quantity = S.number(r.Antal, NUMBER);
      const price = S.number(r.Kurs, NUMBER);
      return {
        account: { number: '', name: account, currency: currency },
        date: date,
        valueDate: '',
        amount: amount,
        currency: currency,
        counterparty: security,
        description: type,
        category: S.code(type),
        reference: isin,
        // Columns Avanza may recompute or rename (fee, exchange rate, result) are left out, so ids stay stable.
        fingerprint: [date, account, type, security, quantity, price, amount, currency, isin]
      };
    }
  };

  ns.CsvAvanza = { FORMAT: FORMAT };
})(globalThis.BankSync);
