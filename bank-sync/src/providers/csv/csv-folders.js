/**
 * @fileoverview CSV connections come from Script Properties: one Drive folder per property,
 * `CSV_FOLDER_<NAME>` = the folder id (pasted by the owner; not a secret). One folder can hold several people's
 * exports: the name only labels the folder, and each account's owner comes from `AccountOwners`. What the last read of
 * each folder found is kept in `CSV_IMPORTS_<HOLDER>` (counts and dates only, never amounts or
 * text) for Show status.
 */
(function (ns) {
  const CSV = ns.Config.CSV;
  const KEYS = ns.Config.PROPERTY_KEYS;

  /**
   * @typedef {Object} SourceRead
   * @property {number} files - Files of this source read.
   * @property {number} rows - Transactions found (after removing the overlap between files).
   * @property {string} from - Earliest booking date, or ''.
   * @property {string} to - Latest booking date, or ''.
   */

  /**
   * @typedef {Object} ImportRecord
   * @property {string} at - ISO date-time of the read.
   * @property {Object<string, SourceRead>} sources - Per format id.
   * @property {Array<string>} skipped - Files not imported, each with the reason (redacted); the first
   *   `CSV.MAX_SKIPPED_LISTED` only, each cut to `CSV.MAX_SKIPPED_CHARS`.
   * @property {number} [skippedMore] - Further files not imported and not named.
   */

  /**
   * @param {string} suffix - Holder suffix of the property name.
   * @returns {import('../../core/config.js').ConnectionConfig} Connection keyed 'csv-<holder>'.
   */
  function connectionFor(suffix) {
    const holder = ns.Connections.slug(suffix);
    return {
      key: CSV.PROVIDER_ID + '-' + holder,
      holder: holder,
      provider: CSV.PROVIDER_ID,
      propertySuffix: suffix,
      bank: { aspspName: CSV.BANK_NAME, country: '' }
    };
  }

  /**
   * @param {{getKeys: function(): Array<string>}} props - Script Properties.
   * @returns {Array<import('../../core/config.js').ConnectionConfig>} One connection per folder property.
   */
  function connections(props) {
    const prefix = KEYS.CSV_FOLDER_PREFIX;
    return (props.getKeys() || [])
      .filter(function (k) { return k.indexOf(prefix) === 0 && k.length > prefix.length; })
      .map(function (k) { return k.slice(prefix.length); })
      .sort()
      .map(connectionFor);
  }

  /**
   * @param {{getProperty: function(string): ?string}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - CSV connection.
   * @returns {string} The folder id, or '' if the property is empty.
   */
  function folderId(props, connection) {
    return String(props.getProperty(KEYS.CSV_FOLDER_PREFIX + connection.propertySuffix) || '').trim();
  }

  /**
   * The last-read record of one folder.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): void}} props - Script Properties.
   * @param {import('../../core/config.js').ConnectionConfig} connection - CSV connection.
   * @returns {{load: function(): ?ImportRecord, save: function(ImportRecord): void}} Store.
   */
  function imports(props, connection) {
    const key = KEYS.CSV_IMPORTS_PREFIX + connection.propertySuffix;
    return {
      load: function () {
        try {
          return JSON.parse(props.getProperty(key) || 'null');
        } catch (e) {
          return null;
        }
      },
      save: function (record) { props.setProperty(key, JSON.stringify(record)); }
    };
  }

  /**
   * @param {ImportRecord} record - Last read.
   * @returns {Array<string>} One line per source, then one per skipped file.
   */
  function describe(record) {
    const sources = Object.keys(record.sources).sort().map(function (id) {
      const s = record.sources[id];
      return '  ' + id + ': ' + s.files + (s.files === 1 ? ' file, ' : ' files, ') + s.rows + ' transactions' +
        (s.from ? ', ' + s.from + ' to ' + s.to : '');
    });
    const skipped = record.skipped.map(function (s) { return '  skipped ' + s; });
    if (record.skippedMore) skipped.push('  and ' + record.skippedMore + ' more skipped files');
    return sources.concat(skipped);
  }

  /**
   * Status lines for Show status: the folder property set/MISSING and the last read per source.
   * @param {{getKeys: function(): Array<string>, getProperty: function(string): ?string}} props - Script Properties.
   * @returns {Array<string>} Lines (names, counts and dates only).
   */
  function statusLines(props) {
    const all = connections(props);
    if (!all.length) return [CSV.BANK_NAME + ': no folder set (optional; see the README)'];
    return all.reduce(function (lines, c) {
      const name = KEYS.CSV_FOLDER_PREFIX + c.propertySuffix;
      const record = imports(props, c).load();
      lines.push(ns.Connections.label(c) + ' - ' + name + ': ' + (folderId(props, c) ? 'set' : 'MISSING') +
        '; last read: ' + (record ? record.at.slice(0, 16).replace('T', ' ') + ' UTC' : 'never'));
      return record ? lines.concat(describe(record)) : lines;
    }, []);
  }

  ns.CsvFolders = { connections: connections, folderId: folderId, imports: imports, statusLines: statusLines };
})(globalThis.BankSync);
