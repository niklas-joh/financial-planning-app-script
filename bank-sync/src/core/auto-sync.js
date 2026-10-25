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
   * @property {function(string, number): void} createAfter - One-off trigger for a handler after some ms.
   * @property {function(string): void} remove - Deletes the trigger with this id.
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
   * Applies a trigger plan.
   * @param {TriggerService} triggers - Trigger service.
   * @param {{create: boolean, remove: Array<string>}} plan - From `Schedule.plan`.
   * @param {function(): void} create - Creates the trigger when the plan says so.
   */
  function apply(triggers, plan, create) {
    plan.remove.forEach(function (id) { triggers.remove(id); });
    if (plan.create) create();
  }

  /**
   * Turns the daily sync on; repairs duplicates if there are any.
   * @param {TriggerService} triggers - Trigger service.
   * @returns {boolean} True if a trigger was created, false if one already existed.
   */
  function enable(triggers) {
    const A = ns.Config.AUTOMATION;
    const plan = ns.Schedule.plan(triggers.list(), A.DAILY_HANDLER, true);
    apply(triggers, plan, function () { triggers.createDaily(A.DAILY_HANDLER, A.DAILY_HOUR); });
    return plan.create;
  }

  /**
   * Turns the daily sync off, including a pending follow-up run.
   * @param {TriggerService} triggers - Trigger service.
   * @returns {number} Triggers deleted.
   */
  function disable(triggers) {
    const A = ns.Config.AUTOMATION;
    const list = triggers.list();
    const ids = ns.Schedule.plan(list, A.DAILY_HANDLER, false).remove
      .concat(ns.Schedule.plan(list, A.CONTINUE_HANDLER, false).remove);
    ids.forEach(function (id) { triggers.remove(id); });
    return ids.length;
  }

  /**
   * Runs Sync all from a trigger, emails due consent warnings and records the outcome.
   * @param {AutoSyncDeps} deps - Services.
   * @param {string} source - 'daily' or 'follow-up'.
   * @returns {Object} The stored LastRun record.
   * @throws {Error} Whatever stopped the run, except a busy lock (that schedules a follow-up),
   *   after recording it; Apps Script then emails its own failure notice too.
   */
  function runTriggered(deps, source) {
    const A = ns.Config.AUTOMATION;
    // One-off triggers stay listed after they fire, so clear old follow-ups before deciding anew.
    apply(deps.triggers, ns.Schedule.plan(deps.triggers.list(), A.CONTINUE_HANDLER, false), function () {});
    const previous = deps.lastRun.load();
    const chain = source === 'daily' ? 0 : ((previous && previous.chain) || 0) + 1;
    const followUp = function (unfinished) {
      const go = ns.Schedule.shouldContinue(unfinished, chain, A.MAX_CONTINUATIONS);
      if (go) deps.triggers.createAfter(A.CONTINUE_HANDLER, A.CONTINUE_AFTER_MS);
      return go;
    };

    let outcome;
    try {
      outcome = deps.sync();
    } catch (e) {
      const busy = e.name === 'LockError';
      const failed = ns.LastRun.fromError(e, { source: source, at: deps.nowLabel(), chain: chain, followUp: followUp(busy) });
      deps.lastRun.save(failed);
      if (!busy) throw e;
      return failed;
    }

    let warningsSent = 0;
    let warningsError;
    try {
      warningsSent = ns.ConsentWarnings.sendDue({ consents: deps.consents(), props: deps.props, send: deps.send, nowMs: deps.nowMs });
    } catch (e) {
      warningsError = ns.Redact.text(e.name + ': ' + e.message);
    }
    const record = ns.LastRun.fromOutcome(outcome, {
      source: source, at: deps.nowLabel(), chain: chain, warningsSent: warningsSent,
      warningsError: warningsError, followUp: followUp(ns.Schedule.isIncomplete(outcome))
    });
    deps.lastRun.save(record);
    return record;
  }

  /**
   * Records a Sync all run started from the menu, so Show status always shows the latest run.
   * @param {{save: function(Object): void}} lastRun - Last-run store.
   * @param {Object} outcome - From `SyncAll.run`.
   * @param {string} at - Local date and time.
   */
  function recordManual(lastRun, outcome, at) {
    lastRun.save(ns.LastRun.fromOutcome(outcome, { source: 'menu', at: at, chain: 0, warningsSent: 0, followUp: false }));
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
    return ['Daily sync: ' + (on
      ? 'on (between ' + A.DAILY_HOUR + ':00 and ' + (A.DAILY_HOUR + 1) + ':00, ' + ns.Config.TIMEZONE + ')'
      : 'off (Enable daily sync to turn it on; consent warnings are emailed by it)')]
      .concat(ns.LastRun.describe(last));
  }

  ns.AutoSync = { enable: enable, disable: disable, runTriggered: runTriggered, recordManual: recordManual, statusLines: statusLines };
})(globalThis.BankSync);
