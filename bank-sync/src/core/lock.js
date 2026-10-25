/**
 * @fileoverview One way to take the script lock, so syncs, the connect step and the bank
 * redirect never interleave their reads and writes of the shared Script Properties.
 */
(function (ns) {
  /**
   * Runs an action while holding the script lock.
   * @param {{getScriptLock: function(): {tryLock: function(number): boolean, releaseLock: function(): void}}} lockService - `LockService`.
   * @param {number} waitMs - How long to wait for the lock.
   * @param {function(): *} action - Work to run while holding the lock.
   * @returns {*} What the action returns.
   * @throws {Error} LockError if the lock is not free within `waitMs`; the action is not run then.
   */
  function run(lockService, waitMs, action) {
    const lock = lockService.getScriptLock();
    if (!lock.tryLock(waitMs)) throw ns.Errors.create('LockError', 'Another sync or login is running; try again in a minute');
    try {
      return action();
    } finally {
      lock.releaseLock();
    }
  }

  ns.Lock = { run: run };
})(globalThis.BankSync);
