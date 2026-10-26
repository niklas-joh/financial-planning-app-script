/**
 * @fileoverview Pure CSV reading: splits text into rows (quoted fields, CRLF, BOM), finds the
 * header row below any preamble lines, parses numbers and dates in the layout a format declares,
 * and has the small text helpers the formats share. Nothing here guesses a layout; the format
 * (`csv-formats.js`) states it.
 */
(function (ns) {
  /**
   * @param {string} message - What is wrong (row numbers only, never cell content).
   * @returns {Error} DataError.
   */
  function dataError(message) {
    return ns.Errors.create('DataError', message);
  }

  /**
   * Splits CSV text into rows of fields. A field that starts with a double quote runs to the
   * closing quote (separators and line breaks inside it are kept, `""` is one quote).
   * @param {string} text - File content.
   * @param {string} separator - Field separator, e.g. ';'.
   * @returns {Array<Array<string>>} Rows as read (not trimmed); a final line break adds no row.
   * @throws {Error} DataError if a quoted field is never closed.
   */
  function rows(text, separator) {
    const s = String(text).replace(/^\uFEFF/, '');
    const out = [];
    let row = [];
    let field = '';
    let quoted = false;
    let i = 0;
    while (i < s.length) {
      const c = s[i];
      if (quoted) {
        if (c !== '"') field += c;
        else if (s[i + 1] === '"') { field += '"'; i++; }
        else quoted = false;
      } else if (c === '"' && field === '') {
        quoted = true;
      } else if (c === separator) {
        row.push(field);
        field = '';
      } else if (c === '\r' || c === '\n') {
        row.push(field);
        out.push(row);
        row = [];
        field = '';
        if (c === '\r' && s[i + 1] === '\n') i++;
      } else {
        field += c;
      }
      i++;
    }
    if (quoted) throw dataError('A quoted field is not closed (row ' + (out.length + 1) + ')');
    if (field !== '' || row.length) {
      row.push(field);
      out.push(row);
    }
    return out;
  }

  /**
   * @param {Array<string>} row - Row.
   * @returns {boolean} True if every field is blank.
   */
  function blank(row) {
    return row.every(function (cell) { return cell.trim() === ''; });
  }

  /**
   * Splits rows into the preamble, the header and the records. The header is the first row
   * that contains every required column name (trimmed); trailing empty fields are ignored.
   * @param {Array<Array<string>>} all - Rows from `rows`.
   * @param {Array<string>} required - Column names the header must contain.
   * @returns {?{preamble: Array<Array<string>>, header: Array<string>, records: Array<Object<string,string>>}}
   *   Records keyed by column name (values untrimmed, blank rows dropped), or null if no row is a
   *   header.
   * @throws {Error} DataError if a record has more non-empty fields than the header has columns.
   */
  function table(all, required) {
    const at = all.findIndex(function (row) {
      const names = row.map(function (cell) { return cell.trim(); });
      return required.every(function (name) { return names.indexOf(name) !== -1; });
    });
    if (at === -1) return null;
    const header = all[at].map(function (cell) { return cell.trim(); });
    while (header.length && header[header.length - 1] === '') header.pop();
    const records = [];
    all.slice(at + 1).forEach(function (row, i) {
      if (blank(row)) return;
      if (row.slice(header.length).some(function (cell) { return cell.trim() !== ''; })) {
        throw dataError('Row ' + (at + i + 2) + ' has more fields than the header');
      }
      const record = {};
      header.forEach(function (name, j) { record[name] = row[j] === undefined ? '' : row[j]; });
      records.push(record);
    });
    return { preamble: all.slice(0, at), header: header, records: records };
  }

  /**
   * Parses a number written with the given decimal mark and thousands separator. Spaces
   * (including no-break spaces) and a Unicode minus are accepted; thousands groups must have
   * three digits, so a misplaced separator is an error rather than a wrong value.
   * @param {string} text - Cell text.
   * @param {{decimal: string, thousands: string}} style - E.g. `{decimal: '.', thousands: ','}`.
   * @returns {?number} The number, or null for a blank cell.
   * @throws {Error} DataError if the text is not a number in that style.
   */
  function number(text, style) {
    const t = String(text).replace(/[\s\u00A0\u202F]/g, '').replace(/\u2212/g, '-');
    if (!t) return null;
    const parts = t.split(style.decimal);
    const whole = parts[0];
    const grouped = style.thousands && whole.indexOf(style.thousands) !== -1;
    const groupOk = !grouped || new RegExp('^[+-]?\\d{1,3}(' + escape(style.thousands) + '\\d{3})+$').test(whole);
    const digits = grouped ? whole.split(style.thousands).join('') : whole;
    const fraction = parts.length > 1 ? parts[1] : null;
    // Each part must be digits only, so a '.' in a decimal-comma file is an error, not a decimal point.
    if (parts.length > 2 || !groupOk || !/^[+-]?\d+$/.test(digits) || (fraction !== null && !/^\d+$/.test(fraction))) {
      throw dataError('Not a number in the expected format');
    }
    return Number(digits + (fraction !== null ? '.' + fraction : ''));
  }

  /**
   * @param {string} s - Literal text.
   * @returns {string} The text with regular-expression characters escaped.
   */
  function escape(s) {
    return s.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  }

  /**
   * Parses a date in a fixed layout made of `DD`, `MM`, `YYYY` and literal characters.
   * @param {string} text - Cell text.
   * @param {string} format - E.g. 'DD/MM/YYYY'.
   * @returns {string} ISO date (YYYY-MM-DD).
   * @throws {Error} DataError if the text does not match the layout or is not a real date.
   */
  function date(text, format) {
    const order = [];
    const pattern = escape(format).replace(/YYYY|MM|DD/g, function (token) {
      order.push(token);
      return token === 'YYYY' ? '(\\d{4})' : '(\\d{1,2})';
    });
    const m = new RegExp('^' + pattern + '$').exec(String(text).trim());
    if (!m) throw dataError('Not a date in the format ' + format);
    const part = {};
    order.forEach(function (token, i) { part[token] = Number(m[i + 1]); });
    const d = new Date(Date.UTC(part.YYYY, part.MM - 1, part.DD));
    if (d.getUTCFullYear() !== part.YYYY || d.getUTCMonth() !== part.MM - 1 || d.getUTCDate() !== part.DD) {
      throw dataError('Not a real date');
    }
    return d.toISOString().slice(0, 10);
  }

  /**
   * @param {*} text - Text.
   * @returns {string} Trimmed, with runs of white space collapsed to one space.
   */
  function squeeze(text) {
    return String(text === undefined || text === null ? '' : text).replace(/\s+/g, ' ').trim();
  }

  /**
   * @param {*} text - Label, e.g. 'Spending cashback' or 'Återinv. utdelning'.
   * @returns {string} Upper-case code, accents dropped, e.g. 'SPENDING_CASHBACK', 'ATERINV_UTDELNING'.
   */
  function code(text) {
    return squeeze(text).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase()
      .replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  }

  ns.CsvParse = { rows: rows, table: table, number: number, date: date, squeeze: squeeze, code: code };
})(globalThis.BankSync);
