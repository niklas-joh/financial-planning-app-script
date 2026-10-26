/**
 * @fileoverview Trading 212 history report (app export or `POST /equity/history/exports`), written
 * from the owner's export (2026-10-25): comma-separated, UTF-8, times `YYYY-MM-DD HH:MM:SS+00:00` in
 * UTC. Its `ID` is the API's `reference`, so it names rows the Trading 212 provider already stores,
 * and every record gives its row what the API lacks: `category` = the merchant category, else the
 * action as a code (`CARD_DEBIT`, `SPENDING_CASHBACK`, `DEPOSIT`, ...), so a card payment is told
 * apart from a withdrawal and cashback from a deposit; `counterparty` = the merchant name, else the
 * notes (e.g. `Bank Transfer`); `kind` = 'card' for card spending (a merchant, or a card or cashback
 * action), else 'investment', which decides the tab the row is stored on. Columns that are empty for
 * the exported period may be missing (merchant columns without card payments), so only the ones every
 * report has are required.
 */
(function (ns) {
  /** Action codes that are card spending even without a merchant (e.g. cashback). */
  const CARD_ACTIONS = ['SPENDING_CASHBACK'];

  /**
   * Whether a report record is card spending: it names a merchant, or its action is a card action
   * (`CARD_DEBIT`, a card refund, ...) or cashback.
   * @param {string} merchant - Merchant name or category ('' if none).
   * @param {string} action - Action as a code.
   * @returns {boolean} True for card spending.
   */
  function isCard(merchant, action) {
    return Boolean(merchant) || action.indexOf('CARD_') === 0 || CARD_ACTIONS.indexOf(action) !== -1;
  }

  /** @type {import('../csv-formats.js').CsvFormat} */
  const FORMAT = {
    id: 'trading212-report',
    kind: 'details',
    enriches: ns.Config.TRADING212.PROVIDER_ID,
    bankName: ns.Config.TRADING212.BANK_NAME,
    separator: ',',
    encodings: ['UTF-8'],
    columns: ['Action', 'Time (UTC)', 'ID', 'Total'],
    row: function (r) {
      const S = ns.CsvParse;
      const id = S.squeeze(r.ID);
      if (!id) throw ns.Errors.create('DataError', 'A Trading 212 report row has no ID');
      // The date only feeds the last-read record.
      const date = S.date(S.squeeze(r['Time (UTC)']).slice(0, 10), 'YYYY-MM-DD');
      const merchant = S.squeeze(r['Merchant name']);
      const merchantCategory = S.squeeze(r['Merchant category']);
      const action = S.code(r.Action);
      return {
        id: id,
        date: date,
        counterparty: merchant || S.squeeze(r.Notes),
        category: merchantCategory || action,
        kind: isCard(merchant || merchantCategory, action) ? 'card' : 'investment'
      };
    }
  };

  ns.CsvT212Report = { FORMAT: FORMAT, isCard: isCard };
})(globalThis.BankSync);
