// Turns Config.MASK_ACCOUNT_NUMBERS on or off for one test. Modules read the flag when they run,
// so swapping the frozen Config for a copy is enough; the original is put back afterwards.
function withMasking(BS, on, fn) {
  const saved = BS.Config;
  BS.Config = Object.freeze(Object.assign({}, saved, { MASK_ACCOUNT_NUMBERS: on }));
  try {
    return fn();
  } finally {
    BS.Config = saved;
  }
}

module.exports = { withMasking };
