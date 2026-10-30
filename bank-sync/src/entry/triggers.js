/**
 * @fileoverview Global functions the time-driven triggers call (names in
 * `Config.AUTOMATION`). They run Sync all through `AutoSync.runTriggered`, and the Trading 212
 * report export through `T212Export`.
 */

/** Daily trigger: Sync all, consent warnings, outcome recorded for Show status. */
function triggerDailySync() { BankSync.AutoSync.runTriggered(BankSync.Container.autoSyncDeps(), 'daily'); }

/** One-off trigger: continues a run that ended at the time budget or found the lock busy. */
function triggerSyncContinue() { BankSync.AutoSync.runTriggered(BankSync.Container.autoSyncDeps(), 'follow-up'); }

/** Monthly trigger: asks Trading 212 for a CSV report per key and schedules its download. */
function triggerT212ReportRequest() { BankSync.T212Export.requestAndSchedule(BankSync.Container.exportDeps()); }

/** One-off trigger: saves finished Trading 212 reports to the exports folder; checks again while one is unfinished. */
function triggerT212ReportFetch() { BankSync.T212Export.fetchAndSchedule(BankSync.Container.exportDeps()); }
