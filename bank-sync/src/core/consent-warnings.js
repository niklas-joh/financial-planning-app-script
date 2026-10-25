/**
 * @fileoverview Consent-expiry warnings. Each connection's consent gets one email per
 * threshold in `Config.CONSENT.WARN_DAYS` (e.g. 30, 14, 7 and 1 days before) and one after it
 * lapses. What was sent is kept per connection in a Script Property together with the expiry
 * date it was sent for, so renewing the consent (a new date) starts afresh. The emails name the
 * connection, the date and how to renew: no balances, amounts, counterparties or account numbers.
 */
(function (ns) {
  const DAY_MS = 86400000;
  const EXPIRED = 0;

  /**
   * @typedef {Object} WarnedRecord
   * @property {string} validUntil - The consent expiry the warnings were sent for.
   * @property {Array<number>} sent - Thresholds already sent (days; 0 = lapsed).
   */

  /**
   * @typedef {Object} DueWarning
   * @property {number} threshold - The threshold to warn about now (the smallest one reached; 0 = lapsed).
   * @property {number} daysLeft - Whole days until the consent lapses (negative once lapsed).
   * @property {WarnedRecord} record - What to store once the email is sent: every reached threshold,
   *   so a run that reaches several at once sends a single email.
   */

  /**
   * @param {string} validUntil - ISO date-time the consent lapses.
   * @param {number} nowMs - Now, epoch ms.
   * @returns {number} Whole days left (floor; negative once lapsed), NaN if the date is invalid.
   */
  function daysLeft(validUntil, nowMs) {
    return Math.floor((Date.parse(validUntil) - nowMs) / DAY_MS);
  }

  /**
   * Works out whether a warning is due for one consent.
   * @param {string} validUntil - ISO date-time the consent lapses.
   * @param {?WarnedRecord} record - What was sent before (null if nothing).
   * @param {number} nowMs - Now, epoch ms.
   * @param {Array<number>} thresholds - Days before expiry to warn on.
   * @returns {?DueWarning} The warning to send, or null if none is due (or the date is invalid).
   */
  function due(validUntil, record, nowMs, thresholds) {
    const left = daysLeft(validUntil, nowMs);
    if (isNaN(left)) return null;
    const lapsed = Date.parse(validUntil) <= nowMs;
    const already = record && record.validUntil === validUntil ? record.sent : [];
    const reached = thresholds.filter(function (t) { return left <= t; }).concat(lapsed ? [EXPIRED] : []);
    const pending = reached.filter(function (t) { return already.indexOf(t) === -1; });
    if (!pending.length) return null;
    const sent = already.concat(pending).sort(function (a, b) { return b - a; });
    return { threshold: Math.min.apply(null, pending), daysLeft: left, record: { validUntil: validUntil, sent: sent } };
  }

  /**
   * Builds the warning email.
   * @param {import('./config.js').ConnectionConfig} connection - The connection.
   * @param {string} validUntil - ISO date-time the consent lapses.
   * @param {DueWarning} warning - From `due`.
   * @param {Array<number>} thresholds - Days before expiry warnings are sent on (for the footer).
   * @returns {{subject: string, body: string}} Plain-text email.
   */
  function message(connection, validUntil, warning, thresholds) {
    const label = ns.Connections.label(connection);
    const date = String(validUntil).slice(0, 10);
    const lapsed = warning.threshold === EXPIRED;
    const when = lapsed
      ? 'lapsed on ' + date + '. Sync all cannot read this connection until you renew it.'
      : 'lapses on ' + date + ' (' + (warning.daysLeft < 1 ? 'in less than a day' : 'in ' + warning.daysLeft + ' days') + ').';
    return {
      subject: 'Bank Sync: consent for ' + label + (lapsed ? ' has lapsed' : ' lapses on ' + date),
      body: [
        'The bank consent for ' + label + ' ' + when,
        '',
        'To renew it:',
        '1. Open the Bank Sync sheet.',
        '2. Bank Sync > Connect a bank...: enter the country (' + connection.bank.country + '), the bank (' +
          connection.bank.aspspName + ') and the same holder name (' + connection.holder + '), so this connection is renewed.',
        '3. Open the link, log in at the bank and approve access.',
        '4. Bank Sync > Show status shows the new expiry date.',
        '',
        'This warning is sent ' + thresholds.join(', ') + ' days before the consent lapses, and once after.'
      ].join('\n')
    };
  }

  /**
   * Short note for Show status.
   * @param {string} validUntil - ISO date-time the consent lapses.
   * @param {number} nowMs - Now, epoch ms.
   * @param {Array<number>} thresholds - Warning thresholds; "renew soon" from the largest one.
   * @returns {string} '' when far off, else ' - renew soon' or ' - LAPSED: renew with Connect a bank...'.
   */
  function statusNote(validUntil, nowMs, thresholds) {
    if (Date.parse(validUntil) <= nowMs) return ' - LAPSED: renew with Connect a bank...';
    return daysLeft(validUntil, nowMs) <= Math.max.apply(null, thresholds) ? ' - renew soon (Connect a bank...)' : '';
  }

  /**
   * Sends the due warnings for every connected consent and records them. An email is recorded
   * only after it was sent, so a failed send is retried on the next run.
   * @param {Object} deps - Services.
   * @param {Array<{connection: import('./config.js').ConnectionConfig, validUntil: string}>} deps.consents - Connected consents.
   * @param {{getProperty: function(string): ?string, setProperty: function(string, string): *}} deps.props - Script Properties.
   * @param {function(string, string): void} deps.send - Sends an email (subject, body) to the owner.
   * @param {function(): number} deps.nowMs - Clock.
   * @returns {{sent: number, errors: Array<string>}} Emails sent, and one redacted error per connection
   *   that failed; a failure never stops the other connections.
   */
  function sendDue(deps) {
    const thresholds = ns.Config.CONSENT.WARN_DAYS;
    const prefix = ns.Config.PROPERTY_KEYS.CONSENT_WARNED_PREFIX;
    const result = { sent: 0, errors: [] };
    deps.consents.forEach(function (c) {
      const key = prefix + c.connection.key;
      try {
        const warning = due(c.validUntil, readRecord(deps.props.getProperty(key)), deps.nowMs(), thresholds);
        if (!warning) return;
        const mail = message(c.connection, c.validUntil, warning, thresholds);
        deps.send(mail.subject, mail.body);
        deps.props.setProperty(key, JSON.stringify(warning.record));
        result.sent += 1;
      } catch (e) {
        result.errors.push(c.connection.key + ': ' + ns.Redact.text(e.name + ': ' + e.message));
      }
    });
    return result;
  }

  /**
   * Parses a stored record; an unreadable one counts as nothing sent yet.
   * @param {?string} raw - Stored JSON.
   * @returns {?WarnedRecord} The record, or null.
   */
  function readRecord(raw) {
    try {
      const record = raw ? JSON.parse(raw) : null;
      return record && Array.isArray(record.sent) ? record : null;
    } catch (e) {
      return null;
    }
  }

  ns.ConsentWarnings = { daysLeft: daysLeft, due: due, message: message, statusNote: statusNote, sendDue: sendDue };
})(globalThis.BankSync);
