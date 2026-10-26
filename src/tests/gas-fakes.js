/**
 * Financial Planning Tools - In-memory Apps Script fakes for tests.
 *
 * A small, behaviour-faithful stand-in for SpreadsheetApp/HtmlService so the
 * suites can exercise the real FinancialPlanner modules instead of copies of
 * them. Values written through ranges are stored per sheet and can be read
 * back; formatting calls are recorded and chain like the real API.
 *
 * Only used by the Node test runner (scripts/run-tests.mjs). tests/** is
 * excluded by .claspignore, so this file is never pushed to Apps Script.
 *
 * Usage inside a suite:
 *   const gas = FakeGas.install(global);
 *   const sheet = gas.spreadsheet.insertSheet('Transactions');
 *   sheet.getRange(1, 1, 2, 2).setValues([['a', 'b'], [1, 2]]);
 *   gas.ui.alerts; // every SpreadsheetApp.getUi().alert(...) call
 */
var FakeGas = (function() {
  const FORMAT_PREFIXES = ['set', 'clear', 'merge', 'break', 'apply', 'auto', 'trim', 'randomize', 'shift', 'sort', 'protect', 'activate', 'remove'];

  function columnLetterToNumber(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  function columnToLetter(column) {
    let letter = '';
    while (column > 0) {
      const t = (column - 1) % 26;
      letter = String.fromCharCode(65 + t) + letter;
      column = (column - t - 1) / 26;
    }
    return letter;
  }

  /** Parses "B3", "A1:C4", "A:A" and "R2C3" into {row, col, numRows, numCols}. */
  function parseA1(a1, sheet) {
    let m = /^R(\d+)C(\d+)$/i.exec(a1);
    if (m) return { row: +m[1], col: +m[2], numRows: 1, numCols: 1 };
    const [start, end] = a1.replace(/^.*!/, '').split(':');
    const cell = (s) => {
      const c = /^([A-Z]+)?(\d+)?$/i.exec(s);
      if (!c) throw new Error('FakeGas: unsupported A1 notation "' + a1 + '"');
      return { col: c[1] ? columnLetterToNumber(c[1]) : null, row: c[2] ? +c[2] : null };
    };
    const s = cell(start);
    const e = end ? cell(end) : s;
    const row = s.row || 1;
    const col = s.col || 1;
    const lastRow = e.row || Math.max(sheet.getMaxRows(), row);
    const lastCol = e.col || Math.max(sheet.getMaxColumns(), col);
    return { row, col, numRows: lastRow - row + 1, numCols: lastCol - col + 1 };
  }

  /**
   * Wraps an object so any unimplemented setter-style call (setFontWeight,
   * clearFormat, merge, …) is recorded and returns the object for chaining,
   * while unknown getters throw so a missing fake is noticed immediately.
   */
  function chainable(target, label, calls) {
    return new Proxy(target, {
      get(obj, prop) {
        if (prop in obj || typeof prop === 'symbol') return obj[prop];
        const name = String(prop);
        if (name === 'then' || name === 'toJSON') return undefined;
        if (FORMAT_PREFIXES.some((p) => name.startsWith(p))) {
          return function() {
            calls.push({ target: label, method: name, args: Array.from(arguments) });
            return this;
          };
        }
        throw new Error('FakeGas: ' + label + '.' + name + '() is not implemented. Add it to src/tests/gas-fakes.js.');
      },
    });
  }

  function makeRange(sheet, row, col, numRows, numCols) {
    const range = {
      getSheet: () => sheet,
      getRow: () => row,
      getColumn: () => col,
      getNumRows: () => numRows,
      getNumColumns: () => numCols,
      getLastRow: () => row + numRows - 1,
      getLastColumn: () => col + numCols - 1,
      getA1Notation: () => {
        const s = columnToLetter(col) + row;
        return numRows === 1 && numCols === 1 ? s : s + ':' + columnToLetter(col + numCols - 1) + (row + numRows - 1);
      },
      getValues: () => {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const line = [];
          for (let c = 0; c < numCols; c++) line.push(sheet._get(row + r, col + c));
          out.push(line);
        }
        return out;
      },
      getValue: () => sheet._get(row, col),
      getDisplayValue: () => String(sheet._get(row, col)),
      getDisplayValues: () => range.getValues().map((r) => r.map(String)),
      setValues: function(values) {
        if (values.length !== numRows || values.some((r) => r.length !== numCols)) {
          throw new Error('The number of rows/columns in the data does not match the range (' + numRows + 'x' + numCols + ').');
        }
        values.forEach((line, r) => line.forEach((v, c) => sheet._set(row + r, col + c, v)));
        return this;
      },
      setValue: function(value) {
        for (let r = 0; r < numRows; r++) for (let c = 0; c < numCols; c++) sheet._set(row + r, col + c, value);
        return this;
      },
      getFormula: () => sheet._formulas[row + ',' + col] || '',
      getFormulas: () => {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const line = [];
          for (let c = 0; c < numCols; c++) line.push(sheet._formulas[(row + r) + ',' + (col + c)] || '');
          out.push(line);
        }
        return out;
      },
      setFormula: function(formula) {
        for (let r = 0; r < numRows; r++) {
          for (let c = 0; c < numCols; c++) {
            sheet._formulas[(row + r) + ',' + (col + c)] = formula;
            sheet._set(row + r, col + c, formula);
          }
        }
        return this;
      },
      setFormulas: function(formulas) {
        formulas.forEach((line, r) => line.forEach((f, c) => {
          sheet._formulas[(row + r) + ',' + (col + c)] = f;
          sheet._set(row + r, col + c, f);
        }));
        return this;
      },
      clearContent: function() {
        for (let r = 0; r < numRows; r++) for (let c = 0; c < numCols; c++) sheet._set(row + r, col + c, '');
        return this;
      },
      clear: function() {
        return this.clearContent();
      },
      insertCheckboxes: function() {
        for (let r = 0; r < numRows; r++) for (let c = 0; c < numCols; c++) sheet._set(row + r, col + c, false);
        sheet._calls.push({ target: 'Range', method: 'insertCheckboxes', args: [range.getA1Notation()] });
        return this;
      },
      setDataValidation: function(rule) {
        for (let r = 0; r < numRows; r++) for (let c = 0; c < numCols; c++) sheet._validations[(row + r) + ',' + (col + c)] = rule;
        return this;
      },
      clearDataValidations: function() {
        for (let r = 0; r < numRows; r++) for (let c = 0; c < numCols; c++) delete sheet._validations[(row + r) + ',' + (col + c)];
        return this;
      },
      getDataValidation: () => sheet._validations[row + ',' + col] || null,
      setBackground: function(color) {
        for (let r = 0; r < numRows; r++) for (let c = 0; c < numCols; c++) sheet._backgrounds[(row + r) + ',' + (col + c)] = color;
        return this;
      },
      getBackground: () => sheet._backgrounds[row + ',' + col] || '#ffffff',
      setNote: function(note) {
        sheet._notes[row + ',' + col] = note;
        return this;
      },
      getNote: () => sheet._notes[row + ',' + col] || '',
      offset: (dr, dc, nr, nc) => sheet.getRange(row + dr, col + dc, nr || numRows, nc || numCols),
    };
    return chainable(range, 'Range', sheet._calls);
  }

  function makeChartBuilder(sheet) {
    const spec = { ranges: [], options: {}, type: null, position: null };
    const builder = {
      addRange: function(r) { spec.ranges.push(r); return this; },
      setChartType: function(t) { spec.type = t; return this; },
      setOption: function(k, v) { spec.options[k] = v; return this; },
      setPosition: function(r, c, ox, oy) { spec.position = [r, c, ox, oy]; return this; },
      asPieChart: function() { spec.type = 'PIE'; return this; },
      asColumnChart: function() { spec.type = 'COLUMN'; return this; },
      asLineChart: function() { spec.type = 'LINE'; return this; },
      asBarChart: function() { spec.type = 'BAR'; return this; },
      build: () => ({ spec, getOptions: () => ({ get: (k) => spec.options[k] }) }),
    };
    return chainable(builder, 'EmbeddedChartBuilder', sheet._calls);
  }

  function makeSheet(spreadsheet, name) {
    const cells = new Map();
    let self = null; // the chainable proxy handed out to callers
    const sheet = {
      _calls: [],
      _formulas: {},
      _validations: {},
      _backgrounds: {},
      _notes: {},
      _charts: [],
      _hidden: false,
      _frozenRows: 0,
      _frozenColumns: 0,
      _name: name,
      _get: (r, c) => (cells.has(r + ',' + c) ? cells.get(r + ',' + c) : ''),
      _set: (r, c, v) => {
        if (v === '' || v === null || v === undefined) cells.delete(r + ',' + c);
        else cells.set(r + ',' + c, v);
      },
      getName: () => sheet._name,
      setName: function(n) { sheet._name = n; return this; },
      getParent: () => spreadsheet,
      getSheetId: () => spreadsheet._sheets.indexOf(sheet) + 1,
      getIndex: () => spreadsheet._sheets.indexOf(sheet) + 1,
      getLastRow: () => {
        let last = 0;
        for (const k of cells.keys()) last = Math.max(last, +k.split(',')[0]);
        return last;
      },
      getLastColumn: () => {
        let last = 0;
        for (const k of cells.keys()) last = Math.max(last, +k.split(',')[1]);
        return last;
      },
      getMaxRows: () => Math.max(1000, sheet.getLastRow()),
      getMaxColumns: () => Math.max(26, sheet.getLastColumn()),
      getRange: function(a, b, c, d) {
        if (typeof a === 'string') {
          const p = parseA1(a, sheet);
          return makeRange(self, p.row, p.col, p.numRows, p.numCols);
        }
        return makeRange(self, a, b, c || 1, d || 1);
      },
      getDataRange: () => makeRange(self, 1, 1, Math.max(1, sheet.getLastRow()), Math.max(1, sheet.getLastColumn())),
      appendRow: function(values) {
        const r = sheet.getLastRow() + 1;
        values.forEach((v, i) => sheet._set(r, i + 1, v));
        return this;
      },
      clear: function() {
        cells.clear();
        sheet._formulas = {};
        sheet._validations = {};
        sheet._backgrounds = {};
        // Like the real API, clear() leaves embedded charts in place.
        return this;
      },
      clearContents: function() {
        cells.clear();
        sheet._formulas = {};
        return this;
      },
      deleteRows: function(start, howMany) {
        const last = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        for (let r = start; r <= last; r++) {
          for (let c = 1; c <= lastCol; c++) sheet._set(r, c, sheet._get(r + howMany, c));
        }
        return this;
      },
      newChart: () => makeChartBuilder(self),
      insertChart: function(chart) { sheet._charts.push(chart); return this; },
      getCharts: () => sheet._charts.slice(),
      removeChart: function(chart) { sheet._charts = sheet._charts.filter((c) => c !== chart); return this; },
      setFrozenRows: function(n) { sheet._frozenRows = n; return this; },
      getFrozenRows: () => sheet._frozenRows,
      setFrozenColumns: function(n) { sheet._frozenColumns = n; return this; },
      getFrozenColumns: () => sheet._frozenColumns,
      hideSheet: function() { sheet._hidden = true; return this; },
      showSheet: function() { sheet._hidden = false; return this; },
      isSheetHidden: () => sheet._hidden,
      activate: function() { spreadsheet._active = self; return this; },
      getConditionalFormatRules: () => sheet._conditionalRules || [],
      setConditionalFormatRules: function(rules) { sheet._conditionalRules = rules; return this; },
      /** Test helper: all values as a 2D array (no trailing empty rows). */
      _values: () => (sheet.getLastRow() ? sheet.getDataRange().getValues() : []),
    };
    self = chainable(sheet, 'Sheet', sheet._calls);
    return self;
  }

  function makeSpreadsheet(ui) {
    const ss = {
      _sheets: [],
      _active: null,
      toasts: [],
      getId: () => 'fake-spreadsheet-id',
      getName: () => 'Fake Spreadsheet',
      getUrl: () => 'https://docs.google.com/spreadsheets/d/fake-spreadsheet-id/edit',
      getSpreadsheetLocale: () => 'en_GB',
      getSpreadsheetTimeZone: () => 'Europe/Brussels',
      getSheets: () => ss._sheets.slice(),
      getSheetByName: (n) => ss._sheets.find((s) => s.getName() === n) || null,
      insertSheet: function(n, index) {
        if (ss.getSheetByName(n)) throw new Error('A sheet with the name "' + n + '" already exists.');
        const sheet = makeSheet(ss, n);
        if (typeof index === 'number') ss._sheets.splice(index, 0, sheet);
        else ss._sheets.push(sheet);
        return sheet;
      },
      deleteSheet: function(sheet) { ss._sheets = ss._sheets.filter((s) => s !== sheet); },
      getActiveSheet: () => ss._active || ss._sheets[0] || null,
      setActiveSheet: function(sheet) { ss._active = sheet; return sheet; },
      toast: function(message, title, timeout) { ss.toasts.push({ message, title, timeout }); },
      getUi: () => ui,
    };
    return chainable(ss, 'Spreadsheet', []);
  }

  function makeUi() {
    const ui = {
      Button: { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO', CLOSE: 'CLOSE' },
      ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', YES_NO_CANCEL: 'YES_NO_CANCEL' },
      alerts: [],
      prompts: [],
      sidebars: [],
      dialogs: [],
      menus: [],
      /** Responses returned by the next alert() calls (default: OK). */
      alertResponses: [],
      /** Responses for the next prompt() calls: { button, text }. */
      promptResponses: [],
      alert: function(a, b, c) {
        // alert(prompt) | alert(prompt, buttons) | alert(title, prompt, buttons)
        const entry = arguments.length >= 3
          ? { title: a, message: b, buttons: c }
          : arguments.length === 2 && typeof b === 'string' && !(b in ui.ButtonSet)
            ? { title: a, message: b, buttons: undefined }
            : { title: undefined, message: a, buttons: b };
        ui.alerts.push(entry);
        return ui.alertResponses.length ? ui.alertResponses.shift() : ui.Button.OK;
      },
      prompt: function(a, b, c) {
        const entry = arguments.length >= 3 ? { title: a, message: b, buttons: c } : { title: undefined, message: a, buttons: b };
        ui.prompts.push(entry);
        const r = ui.promptResponses.length ? ui.promptResponses.shift() : { button: ui.Button.CANCEL, text: '' };
        return { getSelectedButton: () => r.button, getResponseText: () => r.text };
      },
      showSidebar: function(html) { ui.sidebars.push(html); },
      showModalDialog: function(html, title) { ui.dialogs.push({ html, title }); },
      showModelessDialog: function(html, title) { ui.dialogs.push({ html, title, modeless: true }); },
      createMenu: function(title) {
        const menu = { title, items: [] };
        ui.menus.push(menu);
        const builder = {
          addItem: function(label, fn) { menu.items.push({ label, fn }); return this; },
          addSeparator: function() { menu.items.push({ separator: true }); return this; },
          addSubMenu: function(sub) { menu.items.push({ submenu: sub }); return this; },
          addToUi: function() { menu.added = true; },
          _menu: menu,
        };
        return builder;
      },
    };
    return ui;
  }

  function makeHtmlOutput(content) {
    const out = {
      _content: content || '',
      _title: '',
      _width: null,
      _height: null,
      getContent: () => out._content,
      setContent: function(c) { out._content = c; return this; },
      append: function(c) { out._content += c; return this; },
      setTitle: function(t) { out._title = t; return this; },
      getTitle: () => out._title,
      setWidth: function(w) { out._width = w; return this; },
      getWidth: () => out._width,
      setHeight: function(h) { out._height = h; return this; },
      getHeight: () => out._height,
      setSandboxMode: function() { return this; },
      setXFrameOptionsMode: function() { return this; },
      addMetaTag: function() { return this; },
    };
    return out;
  }

  function makeValidationBuilder() {
    const rule = { values: null, showDropdown: null, allowInvalid: false, helpText: '', checkbox: false };
    return {
      requireValueInList: function(values, showDropdown) { rule.values = values.slice(); rule.showDropdown = showDropdown; return this; },
      requireValueInRange: function(range, showDropdown) { rule.range = range; rule.showDropdown = showDropdown; return this; },
      requireCheckbox: function() { rule.checkbox = true; return this; },
      setAllowInvalid: function(v) { rule.allowInvalid = v; return this; },
      setHelpText: function(t) { rule.helpText = t; return this; },
      build: () => Object.assign({ getCriteriaValues: () => [rule.values, rule.showDropdown] }, rule),
    };
  }

  function makeConditionalRuleBuilder() {
    const rule = {};
    const builder = {
      build: () => rule,
    };
    return new Proxy(builder, {
      get(obj, prop) {
        if (prop in obj) return obj[prop];
        return function() { rule[prop] = Array.from(arguments); return this; };
      },
    });
  }

  /** Enum whose unknown members throw, like referencing them in Apps Script yields undefined and later fails. */
  function strictEnum(label, members) {
    const values = {};
    members.forEach((m) => { values[m] = m; });
    return new Proxy(values, {
      get(obj, prop) {
        if (prop in obj || typeof prop === 'symbol') return obj[prop];
        throw new Error('FakeGas: ' + label + '.' + String(prop) + ' does not exist in Apps Script.');
      },
    });
  }

  /**
   * Installs fresh fakes on the given global object and returns handles to
   * inspect them. Call once per suite (each suite runs in its own context).
   */
  function install(g) {
    const ui = makeUi();
    const spreadsheet = makeSpreadsheet(ui);
    const htmlFiles = {};
    g.SpreadsheetApp = {
      getActiveSpreadsheet: () => spreadsheet,
      getActive: () => spreadsheet,
      getUi: () => ui,
      flush: () => {},
      newDataValidation: makeValidationBuilder,
      newConditionalFormatRule: makeConditionalRuleBuilder,
      BorderStyle: strictEnum('SpreadsheetApp.BorderStyle', ['DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE']),
      WrapStrategy: strictEnum('SpreadsheetApp.WrapStrategy', ['WRAP', 'CLIP', 'OVERFLOW']),
    };
    g.HtmlService = {
      createHtmlOutput: (content) => makeHtmlOutput(content),
      createHtmlOutputFromFile: (name) => {
        const out = makeHtmlOutput(htmlFiles[name] || '<!-- ' + name + ' -->');
        out._file = name;
        return out;
      },
      SandboxMode: { IFRAME: 'IFRAME' },
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' },
    };
    g.Charts = {
      ChartType: strictEnum('Charts.ChartType', ['AREA', 'BAR', 'BUBBLE', 'CANDLESTICK', 'COLUMN', 'COMBO', 'HISTOGRAM', 'LINE', 'ORG', 'PIE', 'RADAR', 'SCATTER', 'SPARKLINE', 'STEPPED_AREA', 'TABLE', 'TIMELINE', 'TREEMAP', 'WATERFALL']),
    };
    return { spreadsheet, ui, htmlFiles, columnToLetter };
  }

  return { install, columnToLetter, parseA1 };
})();
