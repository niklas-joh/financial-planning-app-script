// Loads the Node-testable modules in clasp push order and returns the namespace.
// Apps Script-only files (adapters, entry points) are skipped.
const path = require('node:path');
const clasp = require('../.clasp.json');

const SKIP = /^(entry\/|lib\/gas-)/;

clasp.filePushOrder
  .filter((file) => !SKIP.test(file))
  .forEach((file) => require(path.join('..', clasp.rootDir, file)));

module.exports = globalThis.BankSync;
