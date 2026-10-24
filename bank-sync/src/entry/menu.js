/**
 * @fileoverview Spreadsheet menu handlers. The global functions at the bottom
 * are the only names Apps Script needs at top level.
 */
(function (ns) {
  const BANK = ns.Config.BANKS.REVOLUT;

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

  /** Starts Revolut authorization and shows the login link. */
  function authorize() {
    guarded(function () {
      const result = ns.Container.providerFor(BANK).authorize();
      const html = HtmlService.createHtmlOutput(
        '<p>Open the link, log in to ' + escapeHtml(BANK.aspspName) + ', and approve access.</p>' +
        '<p><a href="' + escapeHtml(result.url) + '" target="_blank">Open bank login</a></p>')
        .setWidth(360).setHeight(140);
      SpreadsheetApp.getUi().showModalDialog(html, 'Authorize ' + BANK.aspspName);
    });
  }

  /** Refreshes the raw accounts and balances tabs, serialised by a script lock. */
  function refreshRaw() {
    guarded(function () {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) throw ns.Errors.create('LockError', 'Another sync is running');
      try {
        const writer = ns.RawWriter.create(SpreadsheetApp.getActiveSpreadsheet());
        const counts = ns.RawSync.run(ns.Container.providerFor(BANK), writer);
        SpreadsheetApp.getUi().alert('Wrote ' + counts.accounts + ' accounts and ' + counts.balances + ' balances.');
      } finally {
        lock.releaseLock();
      }
    });
  }

  /** Syncs transactions into the transactions tab, serialised by a script lock. */
  function syncTransactions() {
    guarded(function () {
      const lock = LockService.getScriptLock();
      if (!lock.tryLock(30000)) throw ns.Errors.create('LockError', 'Another sync is running');
      try {
        const store = ns.TransactionsStore.create(
          SpreadsheetApp.getActiveSpreadsheet(), ns.Gas.dateFormatter(ns.Config.TIMEZONE));
        const result = ns.TxnSync.run({
          provider: ns.Container.providerFor(BANK), store: store, nowMs: function () { return Date.now(); }
        });
        SpreadsheetApp.getUi().alert('Transactions: ' + result.added + ' added, ' + result.updated +
          ' updated, ' + result.removed + ' pending removed.' +
          (result.complete ? '' : '\nTime budget reached; run again to continue.'));
      } finally {
        lock.releaseLock();
      }
    });
  }

  /** Shows non-secret setup status: which settings exist and when consent lapses. */
  function status() {
    guarded(function () {
      const s = ns.Container.services();
      const keys = ns.Config.PROPERTY_KEYS;
      const yesNo = function (key) { return s.props.getProperty(key) ? 'set' : 'MISSING'; };
      const session = s.store.loadSession(BANK.key);
      const consent = session
        ? session.accessValidUntil + ' (' + Math.floor((Date.parse(session.accessValidUntil) - s.nowMs()) / 86400000) + ' days left)'
        : 'not authorized';
      SpreadsheetApp.getUi().alert([
        'EB_APP_ID: ' + yesNo(keys.EB_APP_ID),
        'EB_PRIVATE_KEY: ' + yesNo(keys.EB_PRIVATE_KEY),
        'EB_REDIRECT_URL: ' + yesNo(keys.EB_REDIRECT_URL),
        BANK.aspspName + ' consent: ' + consent
      ].join('\n'));
    });
  }

  ns.Menu = { authorize: authorize, refreshRaw: refreshRaw, syncTransactions: syncTransactions, status: status };
})(globalThis.BankSync);

/** Adds the Bank Sync menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bank Sync')
    .addItem('Authorize Revolut', 'menuAuthorize')
    .addItem('Refresh accounts & balances', 'menuRefreshRaw')
    .addItem('Sync transactions', 'menuSyncTransactions')
    .addItem('Show status', 'menuStatus')
    .addToUi();
}

/** Menu entry: start authorization. */
function menuAuthorize() { BankSync.Menu.authorize(); }

/** Menu entry: refresh raw tabs. */
function menuRefreshRaw() { BankSync.Menu.refreshRaw(); }

/** Menu entry: sync transactions. */
function menuSyncTransactions() { BankSync.Menu.syncTransactions(); }

/** Menu entry: show setup status. */
function menuStatus() { BankSync.Menu.status(); }
