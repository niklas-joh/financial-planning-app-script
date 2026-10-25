/**
 * @fileoverview Pure decisions behind the daily sync: which triggers to create or delete so
 * there is never more than one per handler, whether a run ended early (time budget) and
 * whether a follow-up run should continue it. No Apps Script services here.
 */
(function (ns) {
  /**
   * @typedef {Object} TriggerInfo
   * @property {string} id - Trigger unique id.
   * @property {string} handler - Name of the global function it calls.
   */

  /**
   * Plans the trigger changes that leave exactly one trigger for `handler` (on) or none (off).
   * @param {Array<TriggerInfo>} triggers - The project's current triggers.
   * @param {string} handler - Handler function name.
   * @param {boolean} wantOn - Whether a trigger should exist afterwards.
   * @returns {{create: boolean, remove: Array<string>}} Create one, and the ids to delete
   *   (duplicates are removed, the first one is kept).
   */
  function plan(triggers, handler, wantOn) {
    const ids = triggers.filter(function (t) { return t.handler === handler; }).map(function (t) { return t.id; });
    if (!wantOn) return { create: false, remove: ids };
    return { create: ids.length === 0, remove: ids.slice(1) };
  }

  /**
   * @param {Array<TriggerInfo>} triggers - The project's current triggers.
   * @param {string} handler - Handler function name.
   * @returns {boolean} Whether a trigger for `handler` exists.
   */
  function isOn(triggers, handler) {
    return triggers.some(function (t) { return t.handler === handler; });
  }

  /**
   * Whether a Sync all run stopped at the time budget, so running it again continues the work.
   * @param {{results: Array<{status: string, transactions: ?{complete: boolean}}>}} outcome - From `SyncAll.run`.
   * @returns {boolean} True if a connection was skipped or its transactions are unfinished.
   */
  function isIncomplete(outcome) {
    return outcome.results.some(function (r) {
      return r.status === 'skipped: time budget' || (r.status === 'ok' && r.transactions && !r.transactions.complete);
    });
  }

  /**
   * Decides whether to schedule a follow-up run. A daily run starts a fresh chain; each
   * follow-up adds one, up to `max`, so an unfinishable run cannot loop forever.
   * @param {boolean} unfinished - Whether this run left work (time budget, or the lock was busy).
   * @param {number} chain - Follow-ups already run since the last daily run, this one included.
   * @param {number} max - Most follow-ups per daily run.
   * @returns {boolean} True to schedule one more.
   */
  function shouldContinue(unfinished, chain, max) {
    return unfinished && chain < max;
  }

  ns.Schedule = { plan: plan, isOn: isOn, isIncomplete: isIncomplete, shouldContinue: shouldContinue };
})(globalThis.BankSync);
