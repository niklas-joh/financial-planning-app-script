/**
 * @fileoverview Creates the single `BankSync` namespace. Must load first.
 * Every other module attaches itself to this object, so the only globals in
 * the project are `BankSync` and the explicit entry-point functions.
 */
(function (root) {
  /** @namespace BankSync */
  root.BankSync = root.BankSync || {};
})(typeof globalThis !== 'undefined' ? globalThis : this);

if (typeof module !== 'undefined' && module.exports) {
  module.exports = globalThis.BankSync;
}
