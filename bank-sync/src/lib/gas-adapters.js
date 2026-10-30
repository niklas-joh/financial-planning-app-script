/**
 * @fileoverview The only place that touches Apps Script services directly for
 * crypto, HTTP and storage. Everything else receives these as dependencies.
 */
(function (ns) {
  ns.Gas = {
    /** @returns {import('./jwt.js').JwtDeps} Encoder/signer backed by Utilities. */
    jwtDeps: function () {
      return {
        base64UrlEncode: function (input) {
          const encoded = typeof input === 'string'
            ? Utilities.base64EncodeWebSafe(input, Utilities.Charset.UTF_8)
            : Utilities.base64EncodeWebSafe(input);
          return encoded.replace(/=+$/, '');
        },
        signRs256: function (data, privateKey) {
          return Utilities.computeRsaSha256Signature(data, privateKey);
        }
      };
    },

    /**
     * @param {Object} [options] - Extra `Http.create` deps, e.g. `retryDelayMs`.
     * @returns {{request: function(Object): *}} HTTP client backed by UrlFetchApp.
     */
    http: function (options) {
      return ns.Http.create(Object.assign({
        fetch: function (url, params) { return UrlFetchApp.fetch(url, params); },
        sleep: function (ms) { Utilities.sleep(ms); }
      }, options || {}));
    },

    /**
     * @param {string} text - Text.
     * @returns {string} Its Base64 encoding (UTF-8).
     */
    base64: function (text) {
      return Utilities.base64Encode(text, Utilities.Charset.UTF_8);
    },

    /** @returns {Object} Script Properties service. */
    props: function () {
      return PropertiesService.getScriptProperties();
    },

    /**
     * @param {string} timeZone - IANA timezone.
     * @returns {function(Date): string} Formats a Date as YYYY-MM-DD in that timezone.
     */
    dateFormatter: function (timeZone) {
      return function (date) { return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd'); };
    },

    /**
     * One Drive folder (the only DriveApp use; needs `drive`): its files, and saving a new CSV file in it
     * (the Trading 212 report export). Nothing is ever changed, moved or deleted. Google Docs formats are
     * left out of the files: a CSV converted to a Sheet on upload has no CSV content to read.
     * @param {string} id - Folder id.
     * @returns {{files: function(): Array<import('../providers/csv/csv-provider.js').CsvFile>,
     *   create: function(string, string): void}} Its files, each decoded on demand, and a CSV file writer
     *   (name, text).
     */
    driveFolder: function (id) {
      return {
        create: function (name, text) {
          DriveApp.getFolderById(id).createFile(name, text, 'text/csv');
        },
        files: function () {
          const out = [];
          const it = DriveApp.getFolderById(id).getFiles();
          while (it.hasNext()) {
            const file = it.next();
            if (file.getMimeType().indexOf('application/vnd.google-apps') === 0) continue;
            out.push({
              name: file.getName(),
              size: file.getSize(),
              text: function (charset) { return file.getBlob().getDataAsString(charset); }
            });
          }
          return out;
        }
      };
    },

    /** @returns {Object} Script cache service. */
    cache: function () {
      return CacheService.getScriptCache();
    },

    /**
     * @param {string} timeZone - IANA timezone.
     * @returns {function(Date): string} Formats a Date as YYYY-MM-DD HH:mm in that timezone.
     */
    dateTimeFormatter: function (timeZone) {
      return function (date) { return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd HH:mm'); };
    },

    /** @returns {import('../core/auto-sync.js').TriggerService} Project triggers backed by ScriptApp. */
    triggers: function () {
      return {
        list: function () {
          return ScriptApp.getProjectTriggers().map(function (t) {
            return { id: t.getUniqueId(), handler: t.getHandlerFunction() };
          });
        },
        createDaily: function (handler, hour) {
          ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(hour).create();
        },
        createMonthly: function (handler, day, hour) {
          ScriptApp.newTrigger(handler).timeBased().onMonthDay(day).atHour(hour).create();
        },
        createAfter: function (handler, ms) {
          ScriptApp.newTrigger(handler).timeBased().after(ms).create();
        },
        remove: function (ids) {
          ScriptApp.getProjectTriggers().forEach(function (t) {
            if (ids.indexOf(t.getUniqueId()) !== -1) ScriptApp.deleteTrigger(t);
          });
        }
      };
    },

    /**
     * Emails the account the script runs as (the trigger owner); the address is read at run
     * time and never stored.
     * @returns {function(string, string): void} Sends a plain-text email (subject, body).
     * @throws {Error} ConfigError if the address cannot be read.
     */
    mailOwner: function () {
      return function (subject, body) {
        const to = Session.getEffectiveUser().getEmail();
        if (!to) throw ns.Errors.config('Could not read your email address to send the warning');
        MailApp.sendEmail(to, subject, body);
      };
    }
  };
})(globalThis.BankSync);
