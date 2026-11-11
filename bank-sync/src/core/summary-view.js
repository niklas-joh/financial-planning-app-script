/**
 * @fileoverview Turns a summary into the HTML of the summary window: a coloured headline, a table of counts,
 * label and value rows, and a list of things to do. Pure (no Apps Script services); every text is escaped.
 */
(function (ns) {
  /**
   * @typedef {Object} View
   * @property {{tone: string, text: string}} [headline] - One line on top; tone 'ok', 'warn' or 'error'.
   * @property {{title: string, columns: Array<string>, rows: Array<{name: string, counts: (Array<number>|undefined),
   *   note: (string|undefined), error: (string|undefined)}>}} [table] - Counts per row (first column is the name);
   *   a row with `error` shows that text instead of counts.
   * @property {Array<{text: string, error: (boolean|undefined)}>} [lines] - "Label: value" lines, shown as rows.
   * @property {?{title: string, items: Array<{name: string, action: string}>, more: number}} [list] - Things to do.
   * @property {Array<string>} [extra] - More plain lines, shown below.
   */

  const TONES = {
    ok: 'background:#e6f4ea;color:#137333',
    warn: 'background:#fef7e0;color:#7a4f01',
    error: 'background:#fce8e6;color:#c5221f'
  };
  const MUTED = 'color:#5f6368';
  const RULE = 'border-top:1px solid #e8eaed';
  const WIDTH = 560;
  const MAX_HEIGHT = 560;
  /** Longest text before the first ": " that is still shown as a label. */
  const MAX_LABEL = 44;
  /** Characters that fit on one line of the window, for the height estimate. */
  const LINE_CHARS = 70;

  /**
   * Escapes text for safe inclusion in HTML.
   * @param {*} s - Raw text.
   * @returns {string} Escaped text.
   */
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /**
   * @param {string} text - Text.
   * @returns {number} How many window lines it takes, roughly.
   */
  function wraps(text) {
    return Math.max(1, Math.ceil(String(text).length / LINE_CHARS));
  }

  /**
   * One line as a table row: "Label: value" becomes two cells; anything else (and a failure, whose text may
   * hold a colon anywhere) spans both; '' is a gap.
   * @param {{text: string, error: (boolean|undefined)}} line - Line.
   * @returns {string} `<tr>`.
   */
  function lineRow(line) {
    const colour = line.error ? 'color:#c5221f;' : '';
    if (!line.text.trim()) return '<tr><td colspan="2" style="height:10px"></td></tr>';
    const parts = /^([^:]+): (.+)$/.exec(line.text);
    if (!parts || line.error || parts[1].length > MAX_LABEL || line.text.indexOf('- ') === 0) {
      return '<tr><td colspan="2" style="padding:3px 0;' + colour + '">' + escapeHtml(line.text) + '</td></tr>';
    }
    // A lone closing full stop reads oddly in a table cell; sentences inside the value keep theirs.
    const value = parts[2].indexOf('. ') < 0 ? parts[2].replace(/\.$/, '') : parts[2];
    return '<tr><td style="padding:3px 12px 3px 0;vertical-align:top;width:34%;' + MUTED + '">' +
      escapeHtml(parts[1]) + '</td><td style="padding:3px 0;vertical-align:top">' + escapeHtml(value) + '</td></tr>';
  }

  /**
   * @param {Array<{text: string, error: (boolean|undefined)}>} lines - Lines.
   * @returns {string} A table of them, or '' without lines.
   */
  function linesTable(lines) {
    if (!lines.length) return '';
    return '<table style="width:100%;border-collapse:collapse;margin-bottom:14px">' + lines.map(lineRow).join('') + '</table>';
  }

  /**
   * @param {View['table']} table - Counts per row.
   * @returns {string} Title and table. Rows with only zeros are greyed, so the ones that changed stand out.
   */
  function countsTable(table) {
    const right = 'text-align:right;padding:4px 0 4px 12px;';
    const head = '<tr style="font-size:12px;' + MUTED + '">' + table.columns.map(function (c, i) {
      return '<td style="' + (i ? right : 'padding:4px 0;') + '">' + escapeHtml(c) + '</td>';
    }).join('') + '</tr>';
    const rows = table.rows.map(function (r) {
      const name = '<td style="padding:4px 0;' + (r.error ? '' : 'white-space:nowrap;') + '">' + escapeHtml(r.name);
      if (r.error) {
        return '<tr style="' + RULE + ';color:#c5221f">' + name + '</td><td colspan="' + (table.columns.length - 1) +
          '" style="' + right + '">' + escapeHtml(r.error) + '</td></tr>';
      }
      const quiet = r.counts.every(function (n) { return !n; }) && !r.note;
      return '<tr style="' + RULE + (quiet ? ';color:#80868b' : '') + '">' + name +
        (r.note ? '<div style="font-size:12px;color:#7a4f01;white-space:normal">' + escapeHtml(r.note) + '</div>' : '') + '</td>' +
        r.counts.map(function (n) {
          return '<td style="' + right + 'vertical-align:top;' + (n ? 'font-weight:bold;color:#202124' : 'color:#80868b') + '">' +
            escapeHtml(n) + '</td>';
        }).join('') + '</tr>';
    }).join('');
    return '<div style="font-weight:bold;margin-bottom:2px">' + escapeHtml(table.title) + '</div>' +
      '<table style="width:100%;border-collapse:collapse;margin-bottom:14px">' + head + rows + '</table>';
  }

  /**
   * @param {NonNullable<View['list']>} list - Things to do.
   * @returns {string} A yellow box, one item per line.
   */
  function todoBox(list) {
    return '<div style="' + TONES.warn + ';border-radius:8px;padding:10px 12px;margin-bottom:14px">' +
      '<div style="font-weight:bold;margin-bottom:4px">' + escapeHtml(list.title) + '</div>' +
      '<table style="width:100%;border-collapse:collapse">' + list.items.map(function (item) {
        return '<tr><td style="padding:2px 0">' + escapeHtml(item.name) + '</td><td style="padding:2px 0 2px 12px;text-align:right">' +
          escapeHtml(item.action) + '</td></tr>';
      }).join('') + '</table>' + (list.more ? '<div style="padding-top:2px">and ' + escapeHtml(list.more) + ' more</div>' : '') + '</div>';
  }

  /**
   * A summary that is only text (one line each), as a view.
   * @param {string} text - Multi-line summary.
   * @returns {View} View.
   */
  function fromText(text) {
    return { extra: String(text).split('\n') };
  }

  /**
   * @param {View} view - What to show.
   * @returns {{html: string, width: number, height: number}} The window's content and size; the height fits the
   *   content up to a cap, after which the window scrolls.
   */
  function render(view) {
    const lines = view.lines || [];
    const extra = (view.extra || []).map(function (text) { return { text: text }; });
    const html = '<div style="font:13px/1.45 Roboto,Arial,sans-serif;color:#202124">' +
      (view.headline ? '<div style="' + (TONES[view.headline.tone] || TONES.ok) + ';border-radius:8px;padding:8px 12px;' +
        'margin-bottom:14px;font-weight:bold">' + escapeHtml(view.headline.text) + '</div>' : '') +
      (view.table ? countsTable(view.table) : '') +
      linesTable(lines) +
      (view.list ? todoBox(view.list) : '') +
      linesTable(extra) + '</div>';
    const units = (view.headline ? 2.5 : 0) +
      (view.table ? 3 + view.table.rows.reduce(function (n, r) { return n + 1.2 + (r.note || r.error ? 1 : 0); }, 0) : 0) +
      lines.concat(extra).reduce(function (n, l) { return n + wraps(l.text); }, 0) + (lines.length && extra.length ? 1 : 0) +
      (view.list ? 3.5 + view.list.items.length + (view.list.more ? 1 : 0) : 0);
    return { html: html, width: WIDTH, height: Math.min(MAX_HEIGHT, Math.round(50 + units * 22)) };
  }

  ns.SummaryView = { escapeHtml: escapeHtml, fromText: fromText, render: render };
})(globalThis.BankSync);
