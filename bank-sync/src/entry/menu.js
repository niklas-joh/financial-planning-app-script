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

  /** Syncs accounts, balances and transactions for every connected login, then fills the EUR columns. */
  function syncAll() {
    guarded(function () {
      withLock(function () {
        const wired = ns.Container.services();
        const lastRun = ns.LastRun.create(wired.props);
        const now = function () { return ns.Gas.dateTimeFormatter(ns.Config.TIMEZONE)(new Date()); };
        let outcome;
        try {
          outcome = ns.SyncAll.run(ns.Container.syncAllDeps(wired));
        } catch (e) {
          ns.AutoSync.recordManual(lastRun, { error: e }, now());
          throw e;
        }
        ns.AutoSync.recordManual(lastRun, { outcome: outcome }, now());
        SpreadsheetApp.getUi().alert(ns.SyncAll.summarize(outcome));
      });
    });
  }

  /**
   * Imports the CSV files in every holder's Drive folder now: Sync all limited to the CSV connections
   * (Sync all and the daily sync read the folders too). Leaves the Sync all last-run record alone.
   */
  function importCsv() {
    guarded(function () {
      withLock(function () {
        const wired = ns.Container.services();
        const deps = ns.Container.syncAllDeps(wired);
        deps.connections = deps.connections.filter(function (c) { return c.provider === ns.Config.CSV.PROVIDER_ID; });
        if (!deps.connections.length) {
          SpreadsheetApp.getUi().alert('No CSV folder set. Add ' + ns.Config.PROPERTY_KEYS.CSV_FOLDER_PREFIX +
            '<HOLDER> with the Drive folder id in Script Properties (see the README).');
          return;
        }
        const outcome = ns.SyncAll.run(deps);
        SpreadsheetApp.getUi().alert(ns.SyncAll.summarize(outcome) + '\n\n' + ns.CsvFolders.statusLines(wired.props).join('\n'));
      });
    });
  }

  /** Turns the daily sync on (one trigger, never duplicated). */
  function enableDailySync() {
    guarded(function () {
      const created = ns.AutoSync.enable(ns.Gas.triggers());
      SpreadsheetApp.getUi().alert((created ? 'Daily sync is on.' : 'Daily sync was already on.') +
        ' Sync all runs every day ' + ns.AutoSync.scheduleText() +
        ' and emails you before a bank consent lapses. Show status shows the last run.');
    });
  }

  /** Turns the daily sync off, including a pending follow-up run. */
  function disableDailySync() {
    guarded(function () {
      const removed = ns.AutoSync.disable(ns.Gas.triggers());
      SpreadsheetApp.getUi().alert(removed ? 'Daily sync is off. Consent warnings are no longer emailed.' : 'Daily sync was already off.');
    });
  }

  /** Shows non-secret setup status: each provider's settings and connections, daily sync and last run. */
  function status() {
    guarded(function () {
      const s = ns.Container.services();
      SpreadsheetApp.getUi().alert(ns.ProviderRegistry.statusLines(s)
        .concat(ns.AutoSync.statusLines(ns.Gas.triggers().list(), ns.LastRun.create(s.props).load()))
        .join('\n'));
    });
  }

  ns.Menu = { connect: connect, syncAll: syncAll, importCsv: importCsv, status: status, enableDailySync: enableDailySync, disableDailySync: disableDailySync };
})(globalThis.BankSync);

/** Adds the Bank Sync menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bank Sync')
    .addItem('Connect a bank...', 'menuConnect')
    .addItem('Sync all', 'menuSyncAll')
    .addItem('Import CSV files', 'menuImportCsv')
    .addItem('Show status', 'menuStatus')
    .addSeparator()
    .addItem('Enable daily sync', 'menuEnableDailySync')
    .addItem('Disable daily sync', 'menuDisableDailySync')
    .addToUi();
}

/** Menu entry: connect a bank. */
function menuConnect() { BankSync.Menu.connect(); }

/** Menu entry: sync everything. */
function menuSyncAll() { BankSync.Menu.syncAll(); }

/** Menu entry: import the CSV files in the Drive folders now. */
function menuImportCsv() { BankSync.Menu.importCsv(); }

/** Menu entry: show setup status. */
function menuStatus() { BankSync.Menu.status(); }

/** Menu entry: turn the daily sync on. */
function menuEnableDailySync() { BankSync.Menu.enableDailySync(); }

/** Menu entry: turn the daily sync off. */
function menuDisableDailySync() { BankSync.Menu.disableDailySync(); }
