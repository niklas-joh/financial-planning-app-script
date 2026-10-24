/**
 * @fileoverview Tiny deterministic string hash, used to derive stable ids when a
 * bank supplies none. Not cryptographic.
 */
(function (ns) {
  /**
   * 32-bit FNV-1a hash.
   * @param {string} text - Input text.
   * @returns {string} 8 lowercase hex characters.
   */
  function fnv1a(text) {
    let hash = 0x811c9dc5;
    for (let i = 0; i < text.length; i++) {
      hash ^= text.charCodeAt(i);
      hash = Math.imul(hash, 0x01000193) >>> 0;
    }
    return ('00000000' + hash.toString(16)).slice(-8);
  }

  ns.Hash = { fnv1a: fnv1a };
})(globalThis.BankSync);
