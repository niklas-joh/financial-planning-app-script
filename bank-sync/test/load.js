// Loads the Node-testable modules in clasp push order and returns the namespace.
// Apps Script-only files (adapters, entry points) are skipped.
const path = require('node:path');
const clasp = require('../.clasp.json');

const SKIP = /^(entry\/|lib\/gas-)/;

// filePushOrder entries include rootDir (e.g. "src/lib/jwt.js"), as clasp expects.
const prefix = clasp.rootDir + '/';

clasp.filePushOrder
  .map((file) => (file.startsWith(prefix) ? file.slice(prefix.length) : file))
  .filter((file) => !SKIP.test(file))
  .forEach((file) => require(path.join('..', clasp.rootDir, file)));

module.exports = globalThis.BankSync;
