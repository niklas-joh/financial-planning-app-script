/**
 * @fileoverview The last Sync all run's outcome, kept in one Script Property so a run nobody
 * watched (the daily trigger) is visible in Show status. Counts and redacted error messages
 * only: no transaction content, balances or account numbers.
 */
(function (ns) {
  /**
   * @typedef {Object} LastRun
   * @property {string} at - When the run ended (local date and time).
   * @property {string} source - 'daily', 'follow-up' or 'menu'.
   * @property {boolean} ok - False if a connection, the EUR conversion, the warnings or the whole run failed.
   * @property {Array<{connection: string, status: string, added: number, updated: number, removed: number,
   *   complete: boolean, error: (string|undefined)}>} connections - Per-connection counts.
   * @property {?{changed: number, missing: number, waiting: number, failed: number, error: (string|undefined)}} fx - EUR conversion counts.
   * @property {number} warningsSent - Consent warning emails sent.
   * @property {string} [warningsError] - Redacted error from the warning step.
   * @property {string} [error] - Redacted error that stopped the whole run.
   * @property {boolean} [followUp] - Whether a follow-up run was scheduled to finish the work.
   * @property {number} chain - Follow-up runs since the last daily run (0 for a daily or menu run).
   */

  /**
   * Builds the record of a finished run.
   * @param {{results: Array<Object>, fx: ?Object}} outcome - From `SyncAll.run`.
   * @param {{source: string, at: string, chain: number, warningsSent: number, warningsError: (string|undefined),
   *   followUp: boolean}} meta - Run context.
   * @returns {LastRun} The record.
   */
  function fromOutcome(outcome, meta) {
    const connections = outcome.results.map(function (r) {
      const t = r.transactions || { added: 0, updated: 0, removed: 0, complete: r.status !== 'skipped: time budget' };
      return {
        connection: r.connection, status: r.status, added: t.added, updated: t.updated,
        removed: t.removed, complete: t.complete, error: r.error
      };
    });
    const f = outcome.fx;
    const fx = f ? { changed: f.changed, missing: f.missing, waiting: f.waiting, failed: f.failed.length, error: f.error } : null;
    const ok = !connections.some(function (c) { return c.status === 'error'; }) &&
      !(fx && (fx.error || fx.failed)) && !meta.warningsError;
    return {
      at: meta.at, source: meta.source, ok: ok, connections: connections, fx: fx,
      warningsSent: meta.warningsSent, warningsError: meta.warningsError, followUp: meta.followUp, chain: meta.chain
    };
  }

  /**
   * Builds the record of a run that failed as a whole (e.g. the lock was busy).
   * @param {Error} error - What stopped it.
   * @param {{source: string, at: string, chain: number, followUp: boolean}} meta - Run context.
   * @returns {LastRun} The record.
   */
  function fromError(error, meta) {
    return {
      at: meta.at, source: meta.source, ok: false, connections: [], fx: null, warningsSent: 0,
      error: ns.Redact.text(error.name + ': ' + error.message), followUp: meta.followUp, chain: meta.chain
    };
  }

  /**
   * Turns a record into Show status lines.
   * @param {?LastRun} run - Stored record, or null.
   * @returns {Array<string>} Lines.
   */
  function describe(run) {
    if (!run) return ['Last sync: none recorded yet'];
    const lines = ['Last sync: ' + run.at + ' (' + run.source + '): ' + (run.ok ? 'OK' : 'FAILED') +
      (run.followUp ? ', unfinished; a follow-up run is scheduled' : '')];
    if (run.error) lines.push('  ' + run.error);
    run.connections.forEach(function (c) {
      lines.push('  ' + c.connection + ': ' + (c.status === 'ok'
        ? c.added + ' added, ' + c.updated + ' updated, ' + c.removed + ' pending removed' + (c.complete ? '' : ' (unfinished)')
        : c.status + (c.error ? ' (' + c.error + ')' : '')));
    });
    if (run.fx) {
      lines.push('  EUR conversion: ' + run.fx.changed + ' converted' +
        (run.fx.missing ? ', ' + run.fx.missing + ' without a rate' : '') +
        (run.fx.failed ? ', rates unavailable for ' + run.fx.failed + ' currencies' : '') +
        (run.fx.error ? ', failed (' + run.fx.error + ')' : ''));
    }
    if (run.warningsSent) lines.push('  Consent warnings emailed: ' + run.warningsSent);
    if (run.warningsError) lines.push('  Consent warnings failed (' + run.warningsError + ')');
    return lines;
  }

  /**
   * Creates the store.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): *}} props - Script Properties.
   * @returns {{load: function(): ?LastRun, save: function(LastRun): void}} Store.
   */
  function create(props) {
    const KEY = ns.Config.PROPERTY_KEYS.LAST_RUN;
    return {
      load: function () { const raw = props.getProperty(KEY); return raw ? JSON.parse(raw) : null; },
      save: function (run) { props.setProperty(KEY, JSON.stringify(run)); }
    };
  }

  ns.LastRun = { fromOutcome: fromOutcome, fromError: fromError, describe: describe, create: create };
})(globalThis.BankSync);
