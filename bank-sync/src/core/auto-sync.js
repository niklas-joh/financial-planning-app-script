/**
 * @fileoverview The daily sync: turning its trigger on and off without duplicates, and running
 * Sync all from a trigger. A run that ends at the time budget (or finds the lock busy) schedules
 * one follow-up run, which continues where the last one stopped, because Sync all always resumes
 * from the stored state. After each triggered run the due consent warnings are emailed and the
 * outcome is recorded for Show status. Apps Script services come in as dependencies.
 */
(function (ns) {
  /**
   * @typedef {Object} TriggerService
   * @property {function(): Array<import('./schedule.js').TriggerInfo>} list - The project's triggers.
   * @property {function(string, number): void} createDaily - Daily trigger for a handler at an hour (script time zone).
   * @property {function(string, number, number): void} createMonthly - Monthly trigger for a handler on a day of
   *   the month at an hour (script time zone).
   * @property {function(string, number): void} createAfter - One-off trigger for a handler after some ms.
   * @property {function(Array<string>): void} remove - Deletes the triggers with these ids.
   */

  /**
   * @typedef {Object} AutoSyncDeps
   * @property {TriggerService} triggers - Trigger service.
   * @property {{load: function(): ?Object, save: function(Object): void}} lastRun - Last-run store.
   * @property {function(): Object} sync - Runs Sync all under the script lock and returns its outcome;
   *   throws LockError if the lock is busy.
   * @property {function(): Array<{connection: Object, validUntil: string}>} consents - Connected consents.
   * @property {function(string, string): void} send - Emails the owner (subject, body).
   * @property {Object} props - Script Properties (consent warnings record).
   * @property {function(): number} nowMs - Clock.
   * @property {function(): string} nowLabel - Local date and time for the record.
   */

  /**
   * @returns {string} When the daily run happens, e.g. "between 6:00 and 7:00 (Europe/Brussels)".
   */
  function scheduleText() {
    const hour = ns.Config.AUTOMATION.DAILY_HOUR;
    return 'between ' + hour + ':00 and ' + (hour + 1) + ':00 (' + ns.Config.TIMEZONE + ')';
  }

  /**
   * @returns {string} When the monthly Trading 212 report is requested, e.g. "on day 1 between 5:00 and 6:00".
   */
  function reportScheduleText() {
    const A = ns.Config.AUTOMATION;
    return 'on day ' + A.REPORT_DAY + ' between ' + A.REPORT_HOUR + ':00 and ' + (A.REPORT_HOUR + 1) + ':00';
  }

  /**
   * Turns the daily sync on, with the monthly Trading 212 report request; repairs duplicates if there are any.
   * @param {TriggerService} triggers - Trigger service.
   * @returns {boolean} True if the daily trigger was created, false if it already existed.
   */
  function enable(triggers) {
    const A = ns.Config.AUTOMATION;
    const list = triggers.list();
    const daily = ns.Schedule.plan(list, A.DAILY_HANDLER, true);
    const monthly = ns.Schedule.plan(list, A.REPORT_REQUEST_HANDLER, true);
    const duplicates = daily.remove.concat(monthly.remove);
    if (duplicates.length) triggers.remove(duplicates);
    if (daily.create) triggers.createDaily(A.DAILY_HANDLER, A.DAILY_HOUR);
    if (monthly.create) triggers.createMonthly(A.REPORT_REQUEST_HANDLER, A.REPORT_DAY, A.REPORT_HOUR);
    return daily.create;
  }

  /**
   * Turns the daily sync off, including a pending follow-up run, the monthly report request and a pending
   * report check.
   * @param {TriggerService} triggers - Trigger service.
   * @returns {number} Triggers deleted.
   */
  function disable(triggers) {
    const A = ns.Config.AUTOMATION;
    const list = triggers.list();
    const ids = [A.DAILY_HANDLER, A.CONTINUE_HANDLER, A.REPORT_REQUEST_HANDLER, A.REPORT_FETCH_HANDLER]
      .reduce(function (all, handler) { return all.concat(ns.Schedule.plan(list, handler, false).remove); }, []);
    if (ids.length) triggers.remove(ids);
    return ids.length;
  }

  /**
   * Runs Sync all from a trigger, emails due consent warnings (also when the sync failed) and
   * records the outcome.
   * @param {AutoSyncDeps} deps - Services.
   * @param {string} source - 'daily' or 'follow-up'.
   * @returns {Object} The stored LastRun record.
   * @throws {Error} Whatever stopped the run, except a busy lock (that schedules a follow-up),
   *   after recording it; Apps Script then emails its own failure notice too.
   */
  function runTriggered(deps, source) {
    const A = ns.Config.AUTOMATION;
    // One-off triggers stay listed after they fire, so clear old follow-ups before deciding anew.
    const stale = ns.Schedule.plan(deps.triggers.list(), A.CONTINUE_HANDLER, false).remove;
    if (stale.length) deps.triggers.remove(stale);
    const chain = source === 'daily' ? 0 : previousChain(deps.lastRun) + 1;
    const followUp = function (unfinished) {
      const go = ns.Schedule.shouldContinue(unfinished, chain, A.MAX_CONTINUATIONS);
      if (go) deps.triggers.createAfter(A.CONTINUE_HANDLER, A.CONTINUE_AFTER_MS);
      return go;
    };

    let outcome;
    let failure;
    try {
      outcome = deps.sync();
    } catch (e) {
      failure = e;
    }
    // Consent expiry does not depend on the sync working, so warn either way.
    const warnings = sendWarnings(deps);
    const meta = {
      source: source, at: deps.nowLabel(), chain: chain,
      warningsSent: warnings.sent, warningsError: warnings.error
    };
    if (failure) {
      const busy = failure.name === 'LockError';
      const failed = ns.LastRun.fromError(failure, Object.assign(meta, { followUp: followUp(busy) }));
      deps.lastRun.save(failed);
      if (!busy) throw failure;
      return failed;
    }
    const record = ns.LastRun.fromOutcome(outcome, Object.assign(meta, { followUp: followUp(ns.Schedule.isIncomplete(outcome)) }));
    deps.lastRun.save(record);
    return record;
  }

  /**
   * Sends the due consent warnings; a failure is reported, never thrown.
   * @param {AutoSyncDeps} deps - Services.
   * @returns {{sent: number, error: (string|undefined)}} Emails sent and the redacted errors joined, if any.
   */
  function sendWarnings(deps) {
    try {
      const result = ns.ConsentWarnings.sendDue({ consents: deps.consents(), props: deps.props, send: deps.send, nowMs: deps.nowMs });
      return { sent: result.sent, error: result.errors.length ? result.errors.join('; ') : undefined };
    } catch (e) {
      return { sent: 0, error: ns.Redact.text(e.name + ': ' + e.message) };
    }
  }

  /**
   * @param {{load: function(): ?Object}} lastRun - Last-run store.
   * @returns {number} The follow-up chain of the stored record (0 if none).
   */
  function previousChain(lastRun) {
    const previous = lastRun.load();
    return (previous && previous.chain) || 0;
  }

  /**
   * Records a Sync all run started from the menu, so Show status always shows the latest run.
   * The follow-up chain of the stored record is kept, so a manual run between a daily run and its
   * follow-ups does not raise the follow-up cap.
   * @param {{load: function(): ?Object, save: function(Object): void}} lastRun - Last-run store.
   * @param {{outcome: (Object|undefined), error: (Error|undefined)}} run - The outcome from `SyncAll.run`, or the error it threw.
   * @param {string} at - Local date and time.
   */
  function recordManual(lastRun, run, at) {
    const meta = { source: 'menu', at: at, chain: previousChain(lastRun), warningsSent: 0, followUp: false };
    lastRun.save(run.error ? ns.LastRun.fromError(run.error, meta) : ns.LastRun.fromOutcome(run.outcome, meta));
  }

  /**
   * Show status lines for the daily sync and the last run.
   * @param {Array<import('./schedule.js').TriggerInfo>} triggers - The project's triggers.
   * @param {?Object} last - Stored LastRun record.
   * @returns {Array<string>} Lines.
   */
  function statusLines(triggers, last) {
    const A = ns.Config.AUTOMATION;
    const on = ns.Schedule.isOn(triggers, A.DAILY_HANDLER);
    const monthly = ns.Schedule.isOn(triggers, A.REPORT_REQUEST_HANDLER);
    return ['Daily sync: ' + (on
      ? 'on, ' + scheduleText()
      : 'off (Enable daily sync to turn it on; consent warnings are emailed by it)'),
    'Monthly Trading 212 report: ' + (monthly ? 'on, requested ' + reportScheduleText() + ', saved about ' +
      Math.round(A.REPORT_FETCH_AFTER_MS / 60000) + ' minutes later' : 'off (Enable daily sync turns it on)')]
      .concat(ns.LastRun.describe(last));
  }

  ns.AutoSync = {
    enable: enable, disable: disable, runTriggered: runTriggered, recordManual: recordManual,
    statusLines: statusLines, scheduleText: scheduleText, reportScheduleText: reportScheduleText
  };
})(globalThis.BankSync);
