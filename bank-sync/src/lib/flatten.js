/**
 * @fileoverview Flattens an API object into one level of dotted keys so every
 * field can be shown as a sheet column. IBANs and other account or card numbers are
 * masked to their last four characters (the holder's name is kept on purpose).
 */
(function (ns) {
  const IBAN = /^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/;
  /** Last path segments that hold an account or card number whatever its shape. */
  const SENSITIVE_KEY = /^(iban|bban|identification|account_?number|card_?number|pan|sort_?code)$/i;
  const MAX_TEXT = 1000;

  /**
   * Converts one leaf to a cell-safe value.
   * @param {*} value - Leaf value.
   * @param {string} path - Dotted path of the value.
   * @returns {string|number|boolean} Number or boolean as is; text truncated; IBAN-shaped values and
   *   values under account/card-number keys masked to their last four characters; null as ''.
   */
  function leaf(value, path) {
    if (value === null || value === undefined) return '';
    const last = path.split('.').pop().replace(/\[\d+\]$/, '');
    const text = String(value);
    if (IBAN.test(text) || (SENSITIVE_KEY.test(last) && text)) return ns.Redact.maskIban(text);
    if (typeof value === 'number' || typeof value === 'boolean') return value;
    return text.slice(0, MAX_TEXT);
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
      result[prefix] = leaf(value, prefix);
    }
    return result;
  }

  ns.Flatten = { flat: flat };
})(globalThis.BankSync);
