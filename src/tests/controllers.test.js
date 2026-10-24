/**
 * Financial Planning Tools - Controller / global entry point tests
 *
 * Covers the global functions Apps Script calls directly: onOpen (menu),
 * onEdit (routing) and doGet (SaltEdge callback page).
 */
(function() {
  const T = FinancialPlanner.Testing;
  let gas;

  T.registerTest("Controllers", "onOpen builds the Financial Tools menu", function() {
    gas = FakeGas.install(global);
    global.onOpen();
    T.assertEquals(1, gas.ui.menus.filter((m) => m.added).length, "One top-level menu should be added.");
    const items = [];
    (function collect(menu) {
      menu.items.forEach((i) => { if (i.fn) items.push(i.fn); if (i.submenu) collect(i.submenu._menu); });
    })(gas.ui.menus.find((m) => m.added));
    T.assertTrue(items.length > 20, "Menu should list the feature entries.");
    items.forEach(function(fn) {
      T.assertEquals('function', typeof global[fn], "Menu handler '" + fn + "' must be a global function.");
    });
  });

  T.registerTest("Controllers", "doGet escapes the error message from the callback URL", function() {
    gas = FakeGas.install(global);
    const page = global.doGet({ parameter: { stage: 'error', error_message: '<img src=x onerror=alert(1)>' } });
    const html = page.getContent();
    T.assertTrue(html.indexOf('<img') === -1, "Raw HTML from the URL must not reach the page.");
    T.assertTrue(html.indexOf('&lt;img src=x onerror=alert(1)&gt;') !== -1, "The message should be shown escaped.");
  });

  T.registerTest("Controllers", "doGet shows the success page", function() {
    gas = FakeGas.install(global);
    const html = global.doGet({ parameter: { stage: 'success' } }).getContent();
    T.assertTrue(html.indexOf('Bank Connected!') !== -1, "Success message should be shown.");
  });

  T.registerTest("Controllers", "plaid-link.html globals call the Plaid client", function() {
    gas = FakeGas.install(global);
    const original = FinancialPlanner.PlaidClient.createLinkToken;
    FinancialPlanner.PlaidClient.createLinkToken = () => ({ link_token: 'link-sandbox-123' });
    try {
      T.assertEquals('link-sandbox-123', global.plaidCreateLinkTokenGlobal().link_token, "Global should delegate to PlaidClient.");
    } finally {
      FinancialPlanner.PlaidClient.createLinkToken = original;
    }
  });
})();
