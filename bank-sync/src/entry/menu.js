/**
 * @fileoverview Spreadsheet menu handlers. The global functions at the bottom
 * are the only names Apps Script needs at top level.
 */
(function (ns) {
  /**
   * Escapes text for safe inclusion in HTML.
   * @param {string} s - Raw text.
   * @returns {string} Escaped text.
   */
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /**
   * Runs an action, showing any failure as a redacted alert.
   * @param {function(): void} action - Menu action.
   */
  function guarded(action) {
    try {
      action();
    } catch (e) {
      SpreadsheetApp.getUi().alert('Bank Sync: ' + ns.Redact.text(e.name + ': ' + e.message));
    }
  }

  /**
   * Lets the user pick one of several connections.
   * @param {Array<import('../core/config.js').ConnectionConfig>} connections - Candidates.
   * @param {string} title - Prompt title.
   * @returns {?import('../core/config.js').ConnectionConfig} The choice, or null if cancelled.
   * @throws {Error} InputError if the answer is not a listed number.
   */
  function pick(connections, title) {
    if (connections.length === 1) return connections[0];
    const ui = SpreadsheetApp.getUi();
    const list = connections.map(function (c, i) { return (i + 1) + '. ' + ns.Connections.label(c); }).join('\n');
    const answer = ui.prompt(title, 'Enter a number:\n' + list, ui.ButtonSet.OK_CANCEL);
    if (answer.getSelectedButton() !== ui.Button.OK) return null;
    return ns.Connections.choose(connections, answer.getResponseText());
  }

  /** Starts authorization for a chosen connection and shows the login link. */
  function connect() {
    guarded(function () {
      const connection = pick(ns.Connections.all(), 'Connect a bank');
      if (!connection) return;
      const result = ns.Container.providerFor(connection).authorize();
      const html = HtmlService.createHtmlOutput(
        '<p>Open the link, log in to ' + escapeHtml(connection.bank.aspspName) + ' as ' + escapeHtml(connection.holder) +
        ', and approve access.</p>' +
        '<p><a href="' + escapeHtml(result.url) + '" target="_blank">Open bank login</a></p>')
        .setWidth(360).setHeight(140);
      SpreadsheetApp.getUi().showModalDialog(html, 'Connect ' + ns.Connections.label(connection));
    });
  }

  /**
   * Runs an action under the script lock so syncs never overlap.
   * @param {function(): void} action - Work to run while holding the lock.
   * @throws {Error} LockError if another sync holds the lock.
   */
  function withLock(action) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(ns.Config.TRANSACTIONS.LOCK_WAIT_MS)) throw ns.Errors.create('LockError', 'Another sync is running');
    try {
      action();
    } finally {
      lock.releaseLock();
    }
  }

  /** Syncs accounts, balances and transactions for every connected login. */
  function syncAll() {
    guarded(function () {
      withLock(function () {
        const wired = ns.Container.services();
        const formatDate = ns.Gas.dateFormatter(ns.Config.TIMEZONE);
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const outcome = ns.SyncAll.run({
          connections: ns.Connections.all(),
          isConnected: function (c) { return Boolean(wired.store.loadSession(c.key)); },
          providerFor: function (c) { return ns.Container.providerFor(c, wired); },
          store: ns.TransactionsStore.create(spreadsheet, formatDate),
          rawWriter: ns.RawWriter.create(spreadsheet),
          nowMs: wired.nowMs,
          todayIso: function () { return formatDate(new Date()); }
        });
        SpreadsheetApp.getUi().alert(ns.SyncAll.summarize(outcome));
      });
    });
  }

  /** Re-reads account names from the bank for every connected login (one call per account). */
  function refreshAccountNames() {
    guarded(function () {
      withLock(function () {
        const wired = ns.Container.services();
        const lines = ns.Connections.all().map(function (c) {
          if (!wired.store.loadSession(c.key)) return c.key + ': not connected';
          try {
            return c.key + ': ' + wired.auth.refreshAccounts(c) + ' accounts refreshed';
          } catch (e) {
            return c.key + ': ' + ns.Redact.text(e.name + ': ' + e.message);
          }
        });
        SpreadsheetApp.getUi().alert(lines.join('\n') + '\nRun "Sync all" to update the sheet.');
      });
    });
  }

  /** Shows non-secret setup status: which settings exist and when each consent lapses. */
  function status() {
    guarded(function () {
      const s = ns.Container.services();
      const keys = ns.Config.PROPERTY_KEYS;
      const yesNo = function (key) { return s.props.getProperty(key) ? 'set' : 'MISSING'; };
      const connections = ns.Connections.all().map(function (c) {
        const session = s.store.loadSession(c.key);
        const consent = session
          ? session.accessValidUntil + ' (' + Math.floor((Date.parse(session.accessValidUntil) - s.nowMs()) / 86400000) + ' days left)'
          : 'not connected';
        return ns.Connections.label(c) + ': ' + consent;
      });
      SpreadsheetApp.getUi().alert([
        'EB_APP_ID: ' + yesNo(keys.EB_APP_ID),
        'EB_PRIVATE_KEY: ' + yesNo(keys.EB_PRIVATE_KEY),
        'EB_REDIRECT_URL: ' + yesNo(keys.EB_REDIRECT_URL)
      ].concat(connections).join('\n'));
    });
  }

  ns.Menu = { connect: connect, syncAll: syncAll, refreshAccountNames: refreshAccountNames, status: status };
})(globalThis.BankSync);

/** Adds the Bank Sync menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bank Sync')
    .addItem('Connect a bank...', 'menuConnect')
    .addItem('Sync all', 'menuSyncAll')
    .addSeparator()
    .addItem('Refresh account names', 'menuRefreshAccountNames')
    .addItem('Show status', 'menuStatus')
    .addToUi();
}

/** Menu entry: connect a bank. */
function menuConnect() { BankSync.Menu.connect(); }

/** Menu entry: sync everything. */
function menuSyncAll() { BankSync.Menu.syncAll(); }

/** Menu entry: refresh account names. */
function menuRefreshAccountNames() { BankSync.Menu.refreshAccountNames(); }

/** Menu entry: show setup status. */
function menuStatus() { BankSync.Menu.status(); }
