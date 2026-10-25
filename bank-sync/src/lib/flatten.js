/**
 * @fileoverview Flattens an API object into one level of dotted keys so every
 * field can be shown as a sheet column. IBAN-looking values are masked.
 */
(function (ns) {
  const IBAN = /^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/;
  const MAX_TEXT = 1000;

  /**
   * Converts one leaf to a cell-safe value.
   * @param {*} value - Leaf value.
   * @returns {string|number|boolean} Number or boolean as is; text truncated and IBANs masked; null as ''.
   */
  function leaf(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'number' || typeof value === 'boolean') return value;
    const text = String(value);
    return IBAN.test(text) ? ns.Redact.maskIban(text) : text.slice(0, MAX_TEXT);
  }

  /**
   * Flattens nested objects and arrays: `{a: {b: 1}, c: [{d: 2}]}` becomes
   * `{'a.b': 1, 'c[0].d': 2}`. Empty objects and arrays produce no key.
   * @param {*} value - Value to flatten.
   * @param {string} [prefix] - Key prefix (used while recursing).
   * @param {Object<string,*>} [out] - Accumulator (used while recursing).
   * @returns {Object<string,string|number|boolean>} Flat map.
   */
  function flat(value, prefix, out) {
    const result = out || {};
    if (value !== null && typeof value === 'object') {
      const isArray = Array.isArray(value);
      Object.keys(value).forEach(function (key) {
        const path = prefix ? (isArray ? prefix + '[' + key + ']' : prefix + '.' + key) : key;
        flat(value[key], path, result);
      });
    } else if (prefix) {
      result[prefix] = leaf(value);
    }
    return result;
  }

  ns.Flatten = { flat: flat };
})(globalThis.BankSync);
