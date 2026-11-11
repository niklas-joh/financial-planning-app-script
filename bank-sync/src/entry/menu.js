/**
 * @fileoverview Spreadsheet menu handlers. The global functions at the bottom
 * are the only names Apps Script needs at top level.
 */
(function (ns) {
  const escapeHtml = ns.SummaryView.escapeHtml;

  /**
   * Shows a summary in a modeless window: unlike `ui.alert`, it does not keep the run waiting until it
   * is closed, so the execution ends as soon as the work is done. Every text is escaped.
   * @param {string} title - Window title.
   * @param {string|import('../core/summary-view.js').View} summary - Text (one line each) or a view
   *   (counts and names only).
   */
  function showSummary(title, summary) {
    const page = ns.SummaryView.render(typeof summary === 'string' ? ns.SummaryView.fromText(summary) : summary);
    const html = HtmlService.createHtmlOutput(page.html).setWidth(page.width).setHeight(page.height);
    SpreadsheetApp.getUi().showModelessDialog(html, title);
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

  /** Starts a one-time Pluxee sign-in: asks whose card, then shows the sign-in link and what to do next. */
  function connectPluxee() {
    guarded(function () {
      const holder = ask('Connect Pluxee', 'Whose Pluxee card is this? A short name, e.g. niklas:');
      if (!holder) return;
      const started = ns.Container.pluxeeConnect().start(holder);
      const html = HtmlService.createHtmlOutput(
        '<p>Open the link and sign in to Pluxee as <b>' + escapeHtml(started.holder) + '</b> ' +
        '(captcha, then the emailed code).</p>' +
        '<p><a href="' + escapeHtml(started.url) + '" target="_blank">Open Pluxee sign-in</a></p>' +
        '<p>Your browser will try to open an address starting <code>app.pluxee.consumers.eva://</code> and fail ' +
        'to — that is expected. Copy that whole address from the address bar, then run ' +
        '<b>Bank Sync &gt; Finish Pluxee connection</b> and paste it.</p>')
        .setWidth(420).setHeight(240);
      SpreadsheetApp.getUi().showModalDialog(html, 'Connect Pluxee (' + started.holder + ')'); // title is plain text (holder is a slug)
    });
  }

  /** Finishes a Pluxee sign-in: takes the pasted redirect address, stores the refresh token. */
  function finishPluxee() {
    guarded(function () {
      const pasted = ask('Finish Pluxee connection',
        'Paste the address your browser landed on (it starts app.pluxee.consumers.eva://home/?code=...):');
      if (!pasted) return;
      const done = ns.Container.pluxeeConnect().finish(pasted);
      SpreadsheetApp.getUi().alert('Pluxee is connected for ' + done.holder + '. ' +
        'Make sure ' + ns.PluxeeCredentials.namesForHolder(done.holder).card +
        ' holds the card id (see the README), then run Sync all.');
    });
  }

  /**
   * Runs an action under the script lock so syncs never overlap. Show results after it returns.
   * @param {function(): *} action - Work to run while holding the lock.
   * @returns {*} What the action returns.
   * @throws {Error} LockError if another sync holds the lock.
   */
  function withLock(action) {
    return ns.Lock.run(LockService, ns.Config.TRANSACTIONS.LOCK_WAIT_MS, action);
  }

  /** Syncs accounts, balances and transactions for every connected login, then fills the EUR columns. */
  function syncAll() {
    guarded(function () {
      const summary = withLock(function () {
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
        return ns.SyncAll.view(outcome);
      });
      showSummary('Bank Sync: Sync all', summary);
    });
  }

  /**
   * Imports the CSV files in every holder's Drive folder now: Sync all limited to the CSV connections
   * (Sync all and the daily sync read the folders too). Leaves the Sync all last-run record alone.
   */
  function importCsv() {
    guarded(function () {
      const summary = withLock(function () {
        const wired = ns.Container.services();
        const deps = ns.Container.syncAllDeps(wired);
        deps.connections = deps.connections.filter(function (c) { return c.provider === ns.Config.CSV.PROVIDER_ID; });
        if (!deps.connections.length) {
          return 'No CSV folder set. Add ' + ns.Config.PROPERTY_KEYS.CSV_FOLDER_PREFIX +
            '<HOLDER> with the Drive folder id in Script Properties (see the README).';
        }
        const outcome = ns.SyncAll.run(deps);
        return ns.SyncAll.view(outcome, ns.CsvFolders.statusLines(wired.props));
      });
      showSummary('Bank Sync: Import CSV files', summary);
    });
  }

  /**
   * Re-reads every Enable Banking connection's stored history and every CSV folder, so rows stored while
   * `MASK_ACCOUNT_NUMBERS` was on are rewritten with full account numbers; lists the rows that stay masked.
   * Leaves the Sync all last-run record alone.
   */
  function reprocess() {
    guarded(function () {
      const summary = withLock(function () {
        const wired = ns.Container.services();
        const sync = ns.Container.syncAllDeps(wired);
        return ns.Reprocess.view(ns.Reprocess.run({ sync: sync, rawReader: sync.rawWriter }));
      });
      showSummary('Bank Sync: Re-process stored rows', summary);
    });
  }

  /**
   * Asks Trading 212 for a report now (as the monthly trigger does); it is saved to the exports folder a
   * few minutes later by a one-off trigger.
   */
  function requestT212Report() {
    guarded(function () {
      showSummary('Bank Sync: Trading 212 report', ns.T212Export.summarize(ns.T212Export.requestAndSchedule(ns.Container.exportDeps())));
    });
  }

  /** Turns the daily sync on, with the monthly Trading 212 report (one trigger each, never duplicated). */
  function enableDailySync() {
    guarded(function () {
      const created = ns.AutoSync.enable(ns.Gas.triggers());
      SpreadsheetApp.getUi().alert((created ? 'Daily sync is on.' : 'Daily sync was already on.') +
        ' Sync all runs every day ' + ns.AutoSync.scheduleText() +
        ' and emails you before a bank consent lapses. A Trading 212 report is requested ' + ns.AutoSync.reportScheduleText() +
        ' each month and saved to the exports folder. Show status shows the last run.');
    });
  }

  /** Turns the daily sync off, including a pending follow-up run and the monthly Trading 212 report. */
  function disableDailySync() {
    guarded(function () {
      const removed = ns.AutoSync.disable(ns.Gas.triggers());
      SpreadsheetApp.getUi().alert(removed
        ? 'Daily sync is off, and so is the monthly Trading 212 report. Consent warnings are no longer emailed.'
        : 'Daily sync was already off.');
    });
  }

  /** Shows non-secret setup status: each provider's settings and connections, daily sync and last run. */
  function status() {
    guarded(function () {
      const s = ns.Container.services();
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      showSummary('Bank Sync: Status', ns.ProviderRegistry.statusLines(s)
        .concat([ns.AccountOwners.statusLine(ns.AccountOwnersTab.create(spreadsheet)), ns.TxnRouting.statusLine(spreadsheet),
          ns.DataCoverage.statusLine(ns.DataCoverageTab.create(spreadsheet, ns.Gas.dateFormatter(ns.Config.TIMEZONE)))])
        .concat(ns.AutoSync.statusLines(ns.Gas.triggers().list(), ns.LastRun.create(s.props).load()))
        .join('\n'));
    });
  }

  ns.Menu = {
    connect: connect, connectPluxee: connectPluxee, finishPluxee: finishPluxee,
    syncAll: syncAll, importCsv: importCsv, reprocess: reprocess, requestT212Report: requestT212Report, status: status,
    enableDailySync: enableDailySync, disableDailySync: disableDailySync
  };
})(globalThis.BankSync);

/** Adds the Bank Sync menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bank Sync')
    .addItem('Connect a bank...', 'menuConnect')
    .addItem('Connect Pluxee...', 'menuConnectPluxee')
    .addItem('Finish Pluxee connection...', 'menuFinishPluxee')
    .addItem('Sync all', 'menuSyncAll')
    .addItem('Import CSV files', 'menuImportCsv')
    .addItem('Request Trading 212 report', 'menuRequestT212Report')
    .addItem('Show status', 'menuStatus')
    .addSeparator()
    .addItem('Re-process stored rows', 'menuReprocess')
    .addSeparator()
    .addItem('Enable daily sync', 'menuEnableDailySync')
    .addItem('Disable daily sync', 'menuDisableDailySync')
    .addToUi();
}

/** Menu entry: connect a bank. */
function menuConnect() { BankSync.Menu.connect(); }

/** Menu entry: start a one-time Pluxee sign-in. */
function menuConnectPluxee() { BankSync.Menu.connectPluxee(); }

/** Menu entry: finish a Pluxee sign-in by pasting the redirect address. */
function menuFinishPluxee() { BankSync.Menu.finishPluxee(); }

/** Menu entry: sync everything. */
function menuSyncAll() { BankSync.Menu.syncAll(); }

/** Menu entry: import the CSV files in the Drive folders now. */
function menuImportCsv() { BankSync.Menu.importCsv(); }

/** Menu entry: rewrite rows stored while account numbers were masked (re-fetch and re-import). */
function menuReprocess() { BankSync.Menu.reprocess(); }

/** Menu entry: ask Trading 212 for a report now; it is saved to the exports folder a few minutes later. */
function menuRequestT212Report() { BankSync.Menu.requestT212Report(); }

/** Menu entry: show setup status. */
function menuStatus() { BankSync.Menu.status(); }

/** Menu entry: turn the daily sync on. */
function menuEnableDailySync() { BankSync.Menu.enableDailySync(); }

/** Menu entry: turn the daily sync off. */
function menuDisableDailySync() { BankSync.Menu.disableDailySync(); }
