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
   * Asks a question and returns the trimmed answer.
   * @param {string} title - Prompt title.
   * @param {string} question - Prompt text.
   * @returns {?string} The answer, or null if cancelled.
   */
  function ask(title, question) {
    const ui = SpreadsheetApp.getUi();
    const answer = ui.prompt(title, question, ui.ButtonSet.OK_CANCEL);
    return answer.getSelectedButton() === ui.Button.OK ? answer.getResponseText().trim() : null;
  }

  /**
   * Lets the user pick a bank from Enable Banking's live list for a country.
   * @param {Object} client - EbClient.
   * @returns {?{bankName: string, country: string}} The bank, or null if cancelled or not resolved.
   */
  function pickBank(client) {
    const ui = SpreadsheetApp.getUi();
    const country = ask('Connect a bank', 'Country code of the bank (e.g. LU, SE, DE):');
    if (!country) return null;
    if (!/^[A-Za-z]{2}$/.test(country)) throw ns.Errors.create('InputError', 'Use a two-letter country code');
    const names = client.listAspsps(country.toUpperCase()).aspsps.map(function (a) { return a.name; });
    const query = ask('Connect a bank', names.length + ' banks in ' + country.toUpperCase() + '. Type (part of) the bank name:');
    if (!query) return null;
    const found = ns.Connections.matchBanks(names, query);
    if (!found.bank) {
      ui.alert(found.matches.length
        ? 'Several banks match. Run Connect again and type more of the name:\n- ' + found.matches.join('\n- ')
        : 'No bank in ' + country.toUpperCase() + ' matches "' + query + '".');
      return null;
    }
    return { bankName: found.bank, country: country.toUpperCase() };
  }

  /** Picks a bank from Enable Banking's list, asks whose login it is, and shows the login link. */
  function connect() {
    guarded(function () {
      const wired = ns.Container.services();
      const bank = pickBank(wired.client);
      if (!bank) return;
      const holder = ask('Connect ' + bank.bankName, 'Whose login is this? A short name, e.g. niklas:');
      if (!holder) return;
      const connection = ns.Connections.build({ bankName: bank.bankName, country: bank.country, holder: holder });
      let result;
      try {
        result = ns.Lock.run(LockService, ns.Config.TRANSACTIONS.LOCK_WAIT_MS, function () {
          return ns.Container.providerFor(connection, wired).authorize();
        });
      } catch (e) {
        const advice = ns.Connections.explainNotEnabled(e);
        if (advice) {
          SpreadsheetApp.getUi().alert(advice);
          return;
        }
        throw e;
      }
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
    ns.Lock.run(LockService, ns.Config.TRANSACTIONS.LOCK_WAIT_MS, action);
  }

  /** Syncs accounts, balances and transactions for every connected login. */
  function syncAll() {
    guarded(function () {
      withLock(function () {
        const wired = ns.Container.services();
        const formatDate = ns.Gas.dateFormatter(ns.Config.TIMEZONE);
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const outcome = ns.SyncAll.run({
          connections: wired.connections.all(),
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

  /** Shows non-secret setup status: which settings exist and when each consent lapses. */
  function status() {
    guarded(function () {
      const s = ns.Container.services();
      const keys = ns.Config.PROPERTY_KEYS;
      const yesNo = function (key) { return s.props.getProperty(key) ? 'set' : 'MISSING'; };
      const connections = s.connections.all().map(function (c) {
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
      ].concat(connections.length ? connections : ['No bank connected yet: use "Connect a bank..."']).join('\n'));
    });
  }

  ns.Menu = { connect: connect, syncAll: syncAll, status: status };
})(globalThis.BankSync);

/** Adds the Bank Sync menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bank Sync')
    .addItem('Connect a bank...', 'menuConnect')
    .addItem('Sync all', 'menuSyncAll')
    .addItem('Show status', 'menuStatus')
    .addToUi();
}

/** Menu entry: connect a bank. */
function menuConnect() { BankSync.Menu.connect(); }

/** Menu entry: sync everything. */
function menuSyncAll() { BankSync.Menu.syncAll(); }

/** Menu entry: show setup status. */
function menuStatus() { BankSync.Menu.status(); }
