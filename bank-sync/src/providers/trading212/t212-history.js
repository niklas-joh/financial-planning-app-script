/**
 * @fileoverview One-time backfill of a Trading 212 account's full history (cash movements and
 * dividends), on top of the normal incremental sync. Each feed is read newest first from where
 * the previous run stopped, a bounded number of pages per run; the position (`nextPagePath`) is
 * kept in Script Properties (`T212_HISTORY_<HOLDER>`), so a rate limit or the time budget only
 * pauses it until the next sync. Once a feed reaches its last page it is done for good.
 */
(function (ns) {
  /** History feeds, in the order they are backfilled. */
  const FEEDS = ['transactions', 'dividends'];

  /**
   * @typedef {Object} FeedState
   * @property {?string} next - `nextPagePath` to read next; null means the first page.
   * @property {boolean} done - True once the last page was read.
   */

  /** @returns {Object<string, FeedState>} State before anything was read. */
  function emptyState() {
    const state = {};
    FEEDS.forEach(function (feed) { state[feed] = { next: null, done: false }; });
    return state;
  }

  /**
   * Whether an error only pauses the backfill (it resumes on the next sync) rather than failing it:
   * the time budget, or any HTTP failure (a rate limit the retry gave up on, a server error). The
   * incremental sync has already read the same endpoints in the same run, so a lasting problem
   * (e.g. a revoked key) still surfaces there.
   * @param {Error} e - Error from reading a page.
   * @returns {boolean} True to pause.
   */
  function isPause(e) {
    return Boolean(e) && (e.name === 'TimeBudget' || e.name === 'HttpError');
  }

  /**
   * Reads up to `maxPages` pages of one feed from its saved position. Pure apart from `readPage`.
   * @param {function(?string): {items: Array<Object>, nextPagePath: ?string}} readPage - Page reader.
   * @param {FeedState} feed - Where the feed stands.
   * @param {number} maxPages - Page budget for this run.
   * @returns {{items: Array<Object>, feed: FeedState}} Items read and the new position (the input is not mutated).
   * @throws {Error} Any error that is not a pause (see `isPause`).
   */
  function step(readPage, feed, maxPages) {
    let items = [];
    let next = feed.next;
    let done = feed.done;
    for (let pages = 0; pages < maxPages && !done; pages++) {
      let page;
      try {
        page = readPage(next);
      } catch (e) {
        if (isPause(e)) break;
        throw e;
      }
      items = items.concat(page.items);
      done = !page.nextPagePath;
      next = page.nextPagePath || null;
    }
    return { items: items, feed: { next: next, done: done } };
  }

  /**
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {string} Script Property holding its backfill state.
   */
  function propertyName(connection) {
    return ns.Config.PROPERTY_KEYS.T212_HISTORY_PREFIX + connection.propertySuffix;
  }

  /**
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {?Object<string, FeedState>} Stored state, or null if none (or unreadable: it then restarts).
   */
  function read(props, connection) {
    try {
      const raw = props.getProperty(propertyName(connection));
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }

  /**
   * Creates the state store of one connection.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): *}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {{load: function(): Object<string, FeedState>, save: function(Object<string, FeedState>): void}} Store.
   */
  function store(props, connection) {
    return {
      load: function () { return Object.assign(emptyState(), read(props, connection) || {}); },
      save: function (state) { props.setProperty(propertyName(connection), JSON.stringify(state)); }
    };
  }

  /**
   * Show status line for one connection.
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - Trading 212 connection.
   * @returns {string} e.g. "Trading 212 (niklas) history: complete".
   */
  function statusLine(props, connection) {
    const state = read(props, connection);
    const done = state && FEEDS.every(function (feed) { return state[feed] && state[feed].done; });
    return ns.Connections.label(connection) + ' history: ' +
      (done ? 'complete' : state ? 'loading older history (continues with each sync)' : 'older history loads with the next sync');
  }

  ns.T212History = { FEEDS: FEEDS, step: step, store: store, statusLine: statusLine };
})(globalThis.BankSync);
