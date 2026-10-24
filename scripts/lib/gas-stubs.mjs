// Minimal stand-ins for the Apps Script services, enough for the application
// files to load in Node. Tests install richer mocks on `global` themselves.
//
// Every service method throws by default, so a test that forgets to mock a
// service it calls fails loudly instead of passing against a silent stub.

function unmocked(path) {
  return () => {
    throw new Error(`${path}() is not mocked in this test. Stub it via global.${path.split('.')[0]} = …`);
  };
}

/** An object whose every property access returns a throwing function. */
function strictService(name, overrides = {}) {
  return new Proxy(overrides, {
    get(target, prop) {
      if (prop in target) return target[prop];
      if (typeof prop === 'symbol') return undefined;
      return unmocked(`${name}.${String(prop)}`);
    },
  });
}

/** Installs the stubs on a vm context's global object. */
export function installGasStubs(g) {

  g.Logger = {
    log: (...args) => g.console.log(args.map(String).join(' ')),
  };

  const memoryCache = () => {
    const store = new Map();
    return {
      get: (k) => (store.has(k) ? store.get(k) : null),
      put: (k, v) => store.set(k, v),
      remove: (k) => store.delete(k),
      removeAll: (keys) => keys.forEach((k) => store.delete(k)),
      getAll: (keys) => Object.fromEntries(keys.filter((k) => store.has(k)).map((k) => [k, store.get(k)])),
      putAll: (obj) => Object.entries(obj).forEach(([k, v]) => store.set(k, v)),
    };
  };
  const scriptCache = memoryCache();
  const userCache = memoryCache();
  const documentCache = memoryCache();
  g.CacheService = {
    getScriptCache: () => scriptCache,
    getUserCache: () => userCache,
    getDocumentCache: () => documentCache,
  };

  const memoryProperties = () => {
    const store = new Map();
    return {
      getProperty: (k) => (store.has(k) ? store.get(k) : null),
      setProperty(k, v) {
        store.set(k, String(v));
        return this;
      },
      deleteProperty: (k) => store.delete(k),
      getProperties: () => Object.fromEntries(store),
      setProperties: (obj) => Object.entries(obj).forEach(([k, v]) => store.set(k, String(v))),
      getKeys: () => [...store.keys()],
      deleteAllProperties: () => store.clear(),
    };
  };
  const scriptProps = memoryProperties();
  const userProps = memoryProperties();
  const documentProps = memoryProperties();
  g.PropertiesService = {
    getScriptProperties: () => scriptProps,
    getUserProperties: () => userProps,
    getDocumentProperties: () => documentProps,
  };

  g.Session = strictService('Session', {
    getScriptTimeZone: () => 'Europe/Luxembourg',
  });

  g.Utilities = strictService('Utilities', {
    formatDate: (date, timeZone, pattern) => {
      // Covers the yyyy-MM-dd / dd/MM/yyyy style patterns used in this project.
      const parts = Object.fromEntries(
        new Intl.DateTimeFormat('en-GB', {
          timeZone,
          year: 'numeric',
          month: '2-digit',
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
          hour12: false,
        })
          .formatToParts(date)
          .map((p) => [p.type, p.value]),
      );
      return pattern
        .replace('yyyy', parts.year)
        .replace('MM', parts.month)
        .replace('dd', parts.day)
        .replace('HH', parts.hour)
        .replace('mm', parts.minute)
        .replace('ss', parts.second);
    },
    sleep: () => {},
    getUuid: () => '00000000-0000-4000-8000-000000000000',
  });

  for (const name of [
    'SpreadsheetApp',
    'HtmlService',
    'UrlFetchApp',
    'ScriptApp',
    'LockService',
    'MailApp',
    'GmailApp',
    'DriveApp',
    'Charts',
    'ContentService',
  ]) {
    g[name] = strictService(name);
  }
}
