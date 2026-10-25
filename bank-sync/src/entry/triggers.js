/**
 * @fileoverview Global functions the time-driven triggers call (names in
 * `Config.AUTOMATION`). They run Sync all through `AutoSync.runTriggered`.
 */

/** Daily trigger: Sync all, consent warnings, outcome recorded for Show status. */
function triggerDailySync() { BankSync.AutoSync.runTriggered(BankSync.Container.autoSyncDeps(), 'daily'); }

/** One-off trigger: continues a run that ended at the time budget or found the lock busy. */
function triggerSyncContinue() { BankSync.AutoSync.runTriggered(BankSync.Container.autoSyncDeps(), 'follow-up'); }
