/**
 * @fileoverview The `account_owners` tab: one row per account with its owner, which the owner edits by hand.
 * Read by header name, so columns can be moved or added (e.g. a notes column); new accounts are only ever
 * appended below the last row, so the code never rewrites what was typed. One batch call each way.
 */
(function (ns) {
  /** Columns of a new tab: what is needed to recognise an account, then its key. */
  const HEADER = ['owner', 'bank', 'accountName', 'ibanMasked', 'connection', 'provider', 'accountId'];

  /**
   * @param {Object} spreadsheet - Spreadsheet-like (getSheetByName, insertSheet).
   * @returns {import('../core/account-owners.js').OwnersTab} The tab.
   */
  function create(spreadsheet) {
    const tabName = ns.Config.TABS.ACCOUNT_OWNERS;

    /**
     * @param {Object} sheet - The tab.
     * @returns {Array<string>} Its header row, trimmed.
     * @throws {Error} ConfigError when the header has no `accountId` or `owner` column.
     */
    const headerOf = function (sheet) {
      const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
      if (header.indexOf('accountId') === -1 || header.indexOf('owner') === -1) {
        throw ns.Errors.config('The first row of ' + tabName + ' needs an "owner" and an "accountId" column');
      }
      return header;
    };

    return {
      load: function () {
        const sheet = spreadsheet.getSheetByName(tabName);
        if (!sheet || sheet.getLastRow() < 1) return [];
        const header = headerOf(sheet);
        if (sheet.getLastRow() < 2) return [];
        return sheet.getRange(2, 1, sheet.getLastRow() - 1, header.length).getValues().map(function (cells) {
          const row = {};
          header.forEach(function (h, i) { if (h) row[h] = String(cells[i]).trim(); });
          return row;
        });
      },
      append: function (rows) {
        if (!rows.length) return;
        const sheet = spreadsheet.getSheetByName(tabName) || spreadsheet.insertSheet(tabName);
        const isNew = sheet.getLastRow() < 1;
        const header = isNew ? HEADER : headerOf(sheet);
        const body = rows.map(function (r) { return header.map(function (h) { return r[h] === undefined ? '' : r[h]; }); });
        const matrix = isNew ? [header].concat(body) : body;
        const range = sheet.getRange(isNew ? 1 : sheet.getLastRow() + 1, 1, matrix.length, header.length);
        range.setNumberFormat('@');
        range.setValues(matrix);
        if (isNew) sheet.setFrozenRows(1);
      }
    };
  }

  ns.AccountOwnersTab = { create: create, HEADER: HEADER };
})(globalThis.BankSync);
