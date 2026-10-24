const test = require('node:test');
const assert = require('node:assert');
const { Config } = require('./load');

test('config is frozen, including nested sections', () => {
  assert.ok(Object.isFrozen(Config));
  assert.ok(Object.isFrozen(Config.ENABLE_BANKING));
  assert.ok(Object.isFrozen(Config.PROPERTY_KEYS));
});

test('reporting currency is EUR', () => {
  assert.strictEqual(Config.REPORTING_CURRENCY, 'EUR');
});

test('config holds property key names only, never secret material', () => {
  assert.ok(!JSON.stringify(Config).includes('BEGIN'));
  assert.strictEqual(Config.PROPERTY_KEYS.EB_PRIVATE_KEY, 'EB_PRIVATE_KEY');
});
