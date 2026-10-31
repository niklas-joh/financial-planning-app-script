/**
 * @fileoverview Who each account belongs to (its `owner`, e.g. 'niklas', 'common' or 'marie-aude'), so the
 * rows of one CSV folder, or of one login, can be split by person. Owners are kept on the `account_owners`
 * tab, one row per account, and the owner types or changes them there. A sync adds a row for every account
 * the tab does not list yet: the owner is filled in from the connection's holder for a bank login or a
 * Trading 212 key, and left empty for a CSV account (one folder holds everyone's exports, so only the owner
 * can say whose a file is). The owner is a label like the bank or account name (see `AccountLabels`): it is
 * stamped on every row of the account, so a changed owner reaches the stored rows at the account's next sync.
 * Owners are compared as slugs ('Marie-Aude' is stored as 'marie-aude'), the form connection keys use; an owner with
 * no Latin letter or digit slugs to '' and counts as no owner.
 */
(function (ns) {
  /** Most accounts without an owner named in a summary or status line. */
  const MAX_LISTED = 10;

  /**
   * @typedef {Object} OwnerRow
   * @property {string} owner - Owner as typed on the tab, or ''.
   * @property {string} bank - Bank label (to recognise the account).
   * @property {string} accountName - Account label.
   * @property {string} ibanMasked - Account number label (the full number when masking is off), or ''.
   * @property {string} connection - Connection key the account was first read through.
   * @property {string} provider - Provider id.
   * @property {string} accountId - Account id (the key).
   */

  /**
   * @typedef {Object} OwnersTab
   * @property {function(): Array<Object>} load - The tab's rows, keyed by header ([] when there is no tab).
   * @property {function(Array<OwnerRow>): void} append - Adds rows below the last one (creates the tab).
   */

  /**
   * @typedef {Object} OwnersOutcome
   * @property {number} added - Accounts added to the tab in this run.
   * @property {Array<string>} missing - Accounts read in this run that have no owner (bank and account name).
   * @property {string} [error] - Redacted message when the tab could not be read or written.
   */

  /**
   * @param {Object} row - Row read from the tab.
   * @param {string} key - Column.
   * @returns {string} The cell as text, '' when the column is missing.
   */
  function cell(row, key) {
    return String(row[key] === undefined ? '' : row[key]).trim();
  }

  /**
   * @param {import('./config.js').ConnectionConfig} connection - Connection an account was read through.
   * @returns {string} The owner a new account starts with: the holder of a login or API key, '' for a CSV folder.
   */
  function defaultOwner(connection) {
    return connection.provider === ns.Config.CSV.PROVIDER_ID ? '' : ns.Connections.slug(connection.holder || '');
  }

  /**
   * @param {{bank: string, accountName: string}} row - Owner row.
   * @returns {string} e.g. "BIL Current Private".
   */
  function nameOf(row) {
    return [row.bank, row.accountName].filter(Boolean).join(' ') || row.accountId;
  }

  /**
   * @param {Array<string>} names - Account names.
   * @returns {string} The first `MAX_LISTED`, then "and N more".
   */
  function listed(names) {
    const more = names.length - MAX_LISTED;
    return names.slice(0, MAX_LISTED).join(', ') + (more > 0 ? ' and ' + more + ' more' : '');
  }

  /**
   * Owners for one run. The tab is read once, on the first account.
   * @param {OwnersTab} tab - The `account_owners` tab.
   * @returns {{assign: function(Array<import('../providers/bank-provider.js').Account>,
   *   import('./config.js').ConnectionConfig): Array<import('../providers/bank-provider.js').Account>,
   *   save: function(): OwnersOutcome}} `assign` gives each account its `owner` (accounts are returned as they are,
   *   without one, when the tab cannot be read, so stored owners are kept); `save` adds the new accounts to the tab.
   */
  function create(tab) {
    let byId = null;
    let error = '';
    const added = [];
    const seen = {};

    const known = function () {
      if (byId) return byId;
      byId = {};
      try {
        tab.load().forEach(function (r) {
          const id = cell(r, 'accountId');
          if (id && !byId[id]) byId[id] = r;
        });
      } catch (e) {
        error = ns.Redact.text(e.name + ': ' + e.message);
      }
      return byId;
    };

    return {
      assign: function (accounts, connection) {
        const rows = known();
        if (error) return accounts;
        return accounts.map(function (a) {
          let row = rows[a.accountId];
          if (!row) {
            const labels = ns.AccountLabels.lookup([a])(a.accountId);
            row = {
              owner: defaultOwner(connection),
              bank: labels.bank,
              accountName: labels.accountName,
              ibanMasked: labels.ibanMasked,
              connection: labels.connection,
              provider: a.provider,
              accountId: a.accountId
            };
            rows[a.accountId] = row;
            added.push(row);
          }
          const owner = ns.Connections.slug(cell(row, 'owner'));
          seen[a.accountId] = { owner: owner, name: nameOf(row) };
          return Object.assign({}, a, { owner: owner });
        });
      },
      save: function () {
        const outcome = { added: 0, missing: [] };
        if (!error && added.length) {
          try {
            tab.append(added);
            outcome.added = added.length;
          } catch (e) {
            error = ns.Redact.text(e.name + ': ' + e.message);
          }
        }
        outcome.missing = Object.keys(seen).filter(function (id) { return !seen[id].owner; })
          .map(function (id) { return seen[id].name; }).sort();
        if (error) outcome.error = error;
        return outcome;
      }
    };
  }

  /**
   * One summary line for a run, or '' when there is nothing to say.
   * @param {?OwnersOutcome} outcome - From `save`.
   * @returns {string} Line (bank and account names only).
   */
  function summaryLine(outcome) {
    if (!outcome) return '';
    const tabName = ns.Config.TABS.ACCOUNT_OWNERS;
    if (outcome.error) return 'Account owners: ' + tabName + ' could not be used (' + outcome.error + '); owners left as they were.';
    const parts = [];
    if (outcome.added) parts.push(outcome.added + (outcome.added === 1 ? ' account' : ' accounts') + ' added to ' + tabName);
    if (outcome.missing.length) {
      parts.push('no owner yet for ' + listed(outcome.missing) + ' (type it in the owner column of ' + tabName +
        '; it applies from the next sync)');
    }
    return parts.length ? 'Account owners: ' + parts.join('; ') + '.' : '';
  }

  /**
   * Show status line: how many accounts the tab lists and which have no owner.
   * @param {OwnersTab} tab - The `account_owners` tab.
   * @returns {string} Line (bank and account names only).
   */
  function statusLine(tab) {
    const tabName = ns.Config.TABS.ACCOUNT_OWNERS;
    let rows;
    try {
      rows = tab.load().filter(function (r) { return cell(r, 'accountId'); });
    } catch (e) {
      return 'Account owners: ' + tabName + ' cannot be read (' + ns.Redact.text(e.name + ': ' + e.message) + ')';
    }
    if (!rows.length) return 'Account owners: none listed yet (the next sync lists every account on ' + tabName + ')';
    const missing = rows.filter(function (r) { return !ns.Connections.slug(cell(r, 'owner')); })
      .map(nameOf).sort();
    return 'Account owners: ' + rows.length + (rows.length === 1 ? ' account' : ' accounts') + ' on ' + tabName +
      (missing.length ? '; no owner yet: ' + listed(missing) : ', all with an owner');
  }

  ns.AccountOwners = { create: create, defaultOwner: defaultOwner, summaryLine: summaryLine, statusLine: statusLine };
})(globalThis.BankSync);
