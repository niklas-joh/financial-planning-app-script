/**
 * @fileoverview Web-app endpoint that receives the bank redirect. Deployed
 * "execute as me, accessible only by me".
 */
(function (ns) {
  /**
   * Handles the redirect: completes the session or reports the bank's error.
   * @param {{parameter: Object<string,string>}} e - doGet event.
   * @returns {string} Plain-text result, free of secrets.
   */
  function handleRedirect(e) {
    const p = e.parameter || {};
    if (p.error) return 'Authorization failed: ' + ns.Redact.text(p.error);
    if (!p.code || !p.state) return 'Missing code or state in the redirect.';
    try {
      const summary = ns.Container.services().auth.complete(p.code, p.state);
      return 'Authorized ' + summary.connectionKey + ' (' + summary.accountCount +
        ' accounts). Consent valid until ' + summary.validUntil + '. You can close this tab.';
    } catch (err) {
      return ns.Connections.explainNotEnabled(err) || 'Authorization failed: ' + ns.Redact.text(err.name + ': ' + err.message);
    }
  }

  ns.WebApp = { handleRedirect: handleRedirect };
})(globalThis.BankSync);

/**
 * Apps Script web-app entry point.
 * @param {{parameter: Object<string,string>}} e - Request event.
 * @returns {GoogleAppsScript.Content.TextOutput} Result text.
 */
function doGet(e) {
  return ContentService.createTextOutput(BankSync.WebApp.handleRedirect(e));
}
