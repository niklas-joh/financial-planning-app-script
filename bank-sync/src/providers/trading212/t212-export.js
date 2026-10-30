/**
 * @fileoverview Trading 212 report export: asks the API for a CSV report (`POST /equity/history/exports`),
 * checks the reports list until it is finished, downloads it and saves it to the holder's exports folder
 * (`CSV_FOLDER_T212_<HOLDER>`). That folder is also a CSV folder, so the next Sync all or *Import CSV files*
 * applies the report's details like a report uploaded by hand. A report is built in the background, so
 * the request and the download run in separate executions: the request schedules a check after a few
 * minutes, and a check that finds the report unfinished schedules the next, up to a limit. The state per
 * key lives in `T212_EXPORT_<HOLDER>` (dates, ids and counts only; never the download link).
 */
(function (ns) {
  const T212 = ns.Config.TRADING212;
  const A = ns.Config.AUTOMATION;
  const KEYS = ns.Config.PROPERTY_KEYS;

  /** What a report contains: everything the report format reads; orders are not imported. */
  const DATA_INCLUDED = Object.freeze({ includeDividends: true, includeInterest: true, includeOrders: false, includeTransactions: true });

  /**
   * @typedef {Object} PendingReport
   * @property {number} reportId - Id from the request.
   * @property {string} from - First day covered (ISO date).
   * @property {string} to - Last day covered (ISO date, the request day).
   * @property {string} requestedAt - ISO date-time of the request.
   * @property {number} checks - Checks of the reports list so far.
   */

  /**
   * @typedef {Object} ExportState
   * @property {?PendingReport} pending - The report asked for and not saved yet.
   * @property {string} coveredTo - Last day covered by a saved report, or '' before the first.
   * @property {?{file: string, from: string, to: string, at: string}} saved - The last saved report.
   * @property {string} error - The last problem (redacted), or ''.
   * @property {?number} catchUp - While older history is still being fetched (a saved report ended before today):
   *   the follow-up requests that failed in a row, retried with each check up to `REPORT_MAX_CHECKS`; null otherwise.
   */

  /**
   * @typedef {Object} ExportClient
   * @property {function(Object): {reportId: number}} requestReport - `T212Client` method.
   * @property {function(): Array<Object>} reports - `T212Client` method.
   * @property {function(string): string} download - `T212Client` method.
   */

  /**
   * @typedef {Object} ExportDeps
   * @property {Array<import('../../core/config.js').ConnectionConfig>} connections - Trading 212 connections with key and secret.
   * @property {function(import('../../core/config.js').ConnectionConfig): ExportClient} clientFor - API client per connection.
   * @property {Object} props - Script Properties.
   * @property {function(string): {create: function(string, string): void}} folder - Drive folder by id (saves a CSV file).
   * @property {import('../../core/auto-sync.js').TriggerService} triggers - Trigger service.
   * @property {function(function(): *): *} withLock - Runs an action under the script lock (the one Sync all takes);
   *   throws LockError if it is busy.
   * @property {function(): number} nowMs - Clock.
   * @property {function(): string} todayIso - Today's date (script time zone).
   * @property {function(import('../../core/config.js').ConnectionConfig): string} [firstDate] - Booking date of
   *   the connection's oldest stored row, or '' (where the first report starts).
   */

  /** @returns {ExportState} State before any report. */
  function emptyState() {
    return { pending: null, coveredTo: '', saved: null, error: '', catchUp: null };
  }

  /**
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {string} Script Property naming its exports folder.
   */
  function folderProperty(connection) {
    return KEYS.T212_EXPORT_FOLDER_PREFIX + connection.propertySuffix;
  }

  /**
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {string} Its exports folder id, or '' when not set.
   */
  function folderId(props, connection) {
    return String(props.getProperty(folderProperty(connection)) || '').trim();
  }

  /**
   * The state store of one connection.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): *}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {{load: function(): ExportState, save: function(ExportState): void}} Store; an unreadable state starts over.
   */
  function store(props, connection) {
    const key = KEYS.T212_EXPORT_PREFIX + connection.propertySuffix;
    return {
      load: function () {
        try {
          return Object.assign(emptyState(), JSON.parse(props.getProperty(key) || 'null') || {});
        } catch (e) {
          return emptyState();
        }
      },
      save: function (state) { props.setProperty(key, JSON.stringify(state)); }
    };
  }

  /**
   * The period the next report covers, at most `EXPORT_MAX_DAYS` long (Trading 212 refuses longer ones): the
   * first one starts at the oldest stored row (else `EXPORT_FIRST_FROM`), each later one a few days before the
   * last saved report's end (a missed month is caught up; the overlap gives the same ids, so it changes nothing).
   * @param {ExportState} state - Current state.
   * @param {string} today - Today (ISO date).
   * @param {string} [firstDate] - The connection's oldest stored booking date, or ''.
   * @returns {{from: string, to: string}} ISO dates, inclusive.
   */
  function rangeFor(state, today, firstDate) {
    let from = state.coveredTo ? ns.DateRange.addDays(state.coveredTo, -T212.EXPORT_OVERLAP_DAYS) : firstDate || T212.EXPORT_FIRST_FROM;
    if (from > today) from = today;
    const last = ns.DateRange.addDays(from, T212.EXPORT_MAX_DAYS - 1);
    return { from: from, to: last < today ? last : today };
  }

  /**
   * @param {number} ms - Epoch ms.
   * @returns {string} ISO date-time in UTC to the second, e.g. `2026-11-01T04:10:00Z`.
   */
  function isoSeconds(ms) {
    return new Date(ms).toISOString().replace(/\.\d{3}Z$/, 'Z');
  }

  /**
   * @param {ExportState} state - State.
   * @returns {boolean} Whether older history is still to be requested and the follow-up has retries left.
   */
  function catchingUp(state) {
    return typeof state.catchUp === 'number' && state.catchUp < A.REPORT_MAX_CHECKS;
  }

  /**
   * Where the first report starts: the oldest booking date among a connection's stored rows. It reaches back as
   * far as the Trading 212 history backfill has (*Show status*: "history: complete").
   * @param {Array<{connection: string, bookingDate: string}>} rows - Stored rows (all three tabs).
   * @param {string} connectionKey - Connection, e.g. 'trading212-niklas'.
   * @returns {string} ISO date, or '' when the connection has no dated row.
   */
  function oldestDate(rows, connectionKey) {
    return rows.reduce(function (first, t) {
      const date = t.connection === connectionKey ? String(t.bookingDate || '') : '';
      return date && (!first || date < first) ? date : first;
    }, '');
  }

  /**
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @param {PendingReport} report - The report.
   * @returns {string} File name, e.g. `trading212-report-niklas-2016-01-01-to-2026-11-01.csv`.
   */
  function fileName(connection, report) {
    return 'trading212-report-' + connection.holder + '-' + report.from + '-to-' + report.to + '.csv';
  }

  /**
   * @param {Error} e - Error.
   * @returns {string} Redacted one-line text.
   */
  function describe(e) {
    return ns.Redact.text(e.name + ': ' + e.message).slice(0, 300);
  }

  /**
   * Requests a report per connection whose exports folder is set. A report still pending is replaced
   * (the new one covers its period too).
   * @param {ExportDeps} deps - Services.
   * @param {Array<string>} [only] - Connection keys to request for; all when omitted.
   * @returns {Array<{connection: string, status: string, from: (string|undefined), to: (string|undefined)}>}
   *   One result per connection: 'requested', 'no folder' or 'error: ...'.
   */
  function request(deps, only) {
    return deps.connections.filter(function (c) { return !only || only.indexOf(c.key) !== -1; }).map(function (c) {
      if (!folderId(deps.props, c)) {
        return { connection: c.key, status: 'no folder: set ' + folderProperty(c) + ' in Script Properties' };
      }
      const states = store(deps.props, c);
      const state = states.load();
      const today = deps.todayIso();
      const range = rangeFor(state, today, state.coveredTo || !deps.firstDate ? '' : deps.firstDate(c));
      try {
        const out = deps.clientFor(c).requestReport({
          dataIncluded: DATA_INCLUDED,
          timeFrom: range.from + 'T00:00:00Z',
          timeTo: range.to === today ? isoSeconds(deps.nowMs()) : range.to + 'T23:59:59Z'
        });
        if (out.reportId === undefined || out.reportId === null) throw ns.Errors.create('DataError', 'Trading 212 returned no report id');
        state.pending = { reportId: out.reportId, from: range.from, to: range.to, requestedAt: new Date(deps.nowMs()).toISOString(), checks: 0 };
        state.error = '';
        state.catchUp = null;
        states.save(state);
        return { connection: c.key, status: 'requested', from: range.from, to: range.to };
      } catch (e) {
        state.error = describe(e);
        if (typeof state.catchUp === 'number') state.catchUp++;
        states.save(state);
        return { connection: c.key, status: 'error: ' + state.error };
      }
    });
  }

  /**
   * Checks one pending report and saves it once finished.
   * @param {ExportDeps} deps - Services.
   * @param {import('../../core/config.js').ConnectionConfig} c - Connection with a pending report.
   * @param {ExportState} state - Its state (updated).
   * @returns {string} 'saved', 'waiting', 'gave up' or 'failed: ...'.
   */
  function check(deps, c, state) {
    const report = state.pending;
    const client = deps.clientFor(c);
    const found = client.reports().filter(function (r) { return String(r.reportId) === String(report.reportId); })[0];
    const status = found ? String(found.status) : 'not listed';
    if (status === 'Finished') {
      if (!found.downloadLink) throw ns.Errors.create('DataError', 'The finished report has no download link');
      const name = fileName(c, report);
      deps.folder(folderId(deps.props, c)).create(name, client.download(found.downloadLink));
      state.saved = { file: name, from: report.from, to: report.to, at: new Date(deps.nowMs()).toISOString() };
      state.coveredTo = report.to;
      state.pending = null;
      state.error = '';
      return 'saved';
    }
    if (status === 'Failed' || status === 'Canceled') {
      state.pending = null;
      state.error = 'Trading 212 report ' + status.toLowerCase();
      return 'failed: ' + state.error;
    }
    return 'waiting';
  }

  /**
   * Checks every pending report: saves the finished ones; an unfinished one (or a failed check) counts one
   * check, and after `REPORT_MAX_CHECKS` it stays pending without further checks until the next request.
   * @param {ExportDeps} deps - Services.
   * @returns {{results: Array<{connection: string, status: string}>, again: boolean, more: Array<string>}} One
   *   result per connection with a pending report, whether a later check is needed, and the connections whose
   *   saved report ended before today, or whose request for the next period failed and has retries left (the
   *   history continues with the next period).
   */
  function fetchPending(deps) {
    let again = false;
    const results = [];
    const more = [];
    deps.connections.forEach(function (c) {
      const states = store(deps.props, c);
      const state = states.load();
      if (!state.pending || state.pending.checks >= A.REPORT_MAX_CHECKS) {
        if (catchingUp(state)) more.push(c.key);
        return;
      }
      let status;
      try {
        status = check(deps, c, state);
      } catch (e) {
        state.error = describe(e);
        status = 'error: ' + state.error;
      }
      if (state.pending) {
        state.pending.checks++;
        if (state.pending.checks < A.REPORT_MAX_CHECKS) again = true;
        else status = 'gave up after ' + state.pending.checks + ' checks (' + status + ')';
      }
      if (status === 'saved') state.catchUp = state.coveredTo < deps.todayIso() ? 0 : null;
      states.save(state);
      if (catchingUp(state) && !state.pending) more.push(c.key);
      results.push({ connection: c.key, status: status });
    });
    return { results: results, again: again, more: more };
  }

  /**
   * Deletes the fetch triggers: one-off triggers stay listed after they fire, so they are cleared before
   * deciding anew (never more than one is left).
   * @param {import('../../core/auto-sync.js').TriggerService} triggers - Trigger service.
   */
  function clearFetchTriggers(triggers) {
    const stale = ns.Schedule.plan(triggers.list(), A.REPORT_FETCH_HANDLER, false).remove;
    if (stale.length) triggers.remove(stale);
  }

  /**
   * @param {ExportDeps} deps - Services.
   * @returns {boolean} Whether a connection has a pending report with checks left, or older history to request
   *   with retries left.
   */
  function hasOpenReport(deps) {
    return deps.connections.some(function (c) {
      const state = store(deps.props, c).load();
      return state.pending ? state.pending.checks < A.REPORT_MAX_CHECKS : catchingUp(state);
    });
  }

  /**
   * Requests the reports under the script lock, then leaves one check scheduled while any report is pending
   * with checks left: also when this request failed and an earlier report is still waiting (monthly trigger and menu).
   * @param {ExportDeps} deps - Services.
   * @returns {Array<{connection: string, status: string, from: (string|undefined), to: (string|undefined)}>} From `request`.
   * @throws {Error} LockError if a sync or another export holds the lock (nothing is requested then).
   */
  function requestAndSchedule(deps) {
    const results = deps.withLock(function () { return request(deps); });
    clearFetchTriggers(deps.triggers);
    if (hasOpenReport(deps)) deps.triggers.createAfter(A.REPORT_FETCH_HANDLER, A.REPORT_FETCH_AFTER_MS);
    return results;
  }

  /**
   * Checks the pending reports under the script lock (fetch trigger). A saved report that ended before today
   * (the history is fetched one `EXPORT_MAX_DAYS` period at a time) is followed at once by the request for
   * the next period. Another check stays scheduled while a report is pending with checks left. A busy lock
   * counts no check: the check is simply scheduled again.
   * @param {ExportDeps} deps - Services.
   * @returns {{results: Array<{connection: string, status: string}>, again: boolean, more: Array<string>,
   *   requested: Array<Object>, busy: (boolean|undefined)}} From `fetchPending` plus the follow-up requests, or
   *   `busy` with no results when the lock was held.
   */
  function fetchAndSchedule(deps) {
    clearFetchTriggers(deps.triggers);
    let out;
    try {
      out = deps.withLock(function () {
        const done = fetchPending(deps);
        done.requested = done.more.length ? request(deps, done.more) : [];
        return done;
      });
    } catch (e) {
      if (e.name !== 'LockError') throw e;
      out = { results: [], again: true, more: [], requested: [], busy: true };
    }
    if (out.busy || hasOpenReport(deps)) deps.triggers.createAfter(A.REPORT_FETCH_HANDLER, A.REPORT_FETCH_AFTER_MS);
    return out;
  }

  /**
   * @param {Array<{connection: string, status: string, from: (string|undefined), to: (string|undefined)}>} results - From `request`.
   * @returns {string} Summary for the menu (names, dates and statuses only).
   */
  function summarize(results) {
    if (!results.length) return 'No Trading 212 key with its secret is set (see the README).';
    return results.map(function (r) {
      return r.connection + ': ' + (r.status === 'requested'
        ? 'report requested for ' + r.from + ' to ' + r.to + '. It is saved to the exports folder in about ' +
          Math.round(A.REPORT_FETCH_AFTER_MS / 60000) + ' minutes; then Import CSV files applies it (the daily sync does too).'
        : r.status);
    }).join('\n');
  }

  /**
   * Status line for one connection.
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {string} Folder set/MISSING, the last saved report, a pending one and the last error (dates only).
   */
  function statusLine(props, connection) {
    const state = store(props, connection).load();
    const parts = [folderProperty(connection) + ': ' + (folderId(props, connection) ? 'set' : 'MISSING')];
    parts.push(state.saved
      ? 'last saved ' + state.saved.at.slice(0, 10) + ' (' + state.saved.from + ' to ' + state.saved.to + ')'
      : 'none saved yet');
    if (state.pending) {
      parts.push('requested ' + state.pending.requestedAt.slice(0, 16).replace('T', ' ') + ' UTC, ' + state.pending.checks + ' checks' +
        (state.pending.checks >= A.REPORT_MAX_CHECKS ? ', no more checks until the next request' : ''));
    }
    if (!state.pending && typeof state.catchUp === 'number') {
      parts.push(catchingUp(state)
        ? 'older history: next period retried with the next check (' + state.catchUp + ' failed)'
        : 'older history: stopped after ' + state.catchUp + ' failed requests; Request Trading 212 report continues it');
    }
    if (state.error) parts.push('last error: ' + state.error);
    return ns.Connections.label(connection) + ' report export - ' + parts.join('; ');
  }

  ns.T212Export = {
    DATA_INCLUDED: DATA_INCLUDED, store: store, rangeFor: rangeFor, oldestDate: oldestDate, fileName: fileName, request: request, fetchPending: fetchPending,
    requestAndSchedule: requestAndSchedule, fetchAndSchedule: fetchAndSchedule, summarize: summarize, statusLine: statusLine
  };
})(globalThis.BankSync);
