/**
 * @fileoverview Creates the single `BankSync` namespace. Must load first.
 * Every other module attaches itself to this object, so the only globals in
 * the project are `BankSync` and the explicit entry-point functions.
 */
globalThis.BankSync = globalThis.BankSync || {};
