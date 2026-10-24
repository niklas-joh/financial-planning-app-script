const test = require('node:test');
const assert = require('node:assert');

require('../src/00_namespace.js');
const Config = require('../src/core/config.js');

test('config is frozen, including nested sections', () => {
  assert.ok(Object.isFrozen(Config));
  assert.ok(Object.isFrozen(Config.ENABLE_BANKING));
  assert.ok(Object.isFrozen(Config.PROPERTY_KEYS));
});

test('reporting currency is EUR', () => {
  assert.strictEqual(Config.REPORTING_CURRENCY, 'EUR');
});

test('config holds property key names only, never secret material', () => {
  const dump = JSON.stringify(Config);
  assert.ok(!dump.includes('BEGIN'), 'no PEM material in config');
  assert.strictEqual(Config.PROPERTY_KEYS.EB_PRIVATE_KEY, 'EB_PRIVATE_KEY');
});
