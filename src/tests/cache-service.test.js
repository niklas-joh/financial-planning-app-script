/**
 * Financial Planning Tools - Cache Service Tests
 *
 * Exercises the real FinancialPlanner.CacheService. Its two layers are
 * observed from the outside: Apps Script's script cache is replaced with a
 * recording fake, and the in-memory layer is inferred from whether
 * computeFunction / the script cache are consulted.
 */
(function() {
  const T = FinancialPlanner.Testing;
  const Cache = FinancialPlanner.CacheService;
  const Config = FinancialPlanner.Config;
  const KNOWN_KEY = Config.getSection('CACHE').KEYS.CATEGORY_COMBINATIONS;

  let store;   // namespacedKey -> { value, expiresAt }
  let calls;   // script-cache calls, for assertions
  let now;     // fake clock (ms)

  const realNow = Date.now;

  function setup() {
    Config.reset();
    store = {};
    calls = [];
    now = 1700000000000;
    Date.now = () => now;
    global.CacheService = {
      getScriptCache: () => ({
        get: (k) => {
          calls.push(['get', k]);
          const e = store[k];
          return e && e.expiresAt > now ? e.value : null;
        },
        put: (k, v, ttl) => {
          calls.push(['put', k, ttl]);
          store[k] = { value: v, expiresAt: now + ttl * 1000 };
        },
        remove: (k) => { calls.push(['remove', k]); delete store[k]; },
        removeAll: (keys) => { calls.push(['removeAll', keys]); keys.forEach((k) => delete store[k]); },
      }),
    };
    // Clear the service's private memory layer between tests.
    Cache.invalidateAll();
    calls = [];
  }

  function teardown() {
    Date.now = realNow;
    Config.reset();
  }

  function test(name, fn) {
    T.registerTest("CacheService", name, function() {
      setup();
      try { fn(); } finally { teardown(); }
    });
  }

  function counter(value) {
    const f = () => { f.count++; return value; };
    f.count = 0;
    return f;
  }

  test("get should compute value on cache miss and store it in the script cache", function() {
    const compute = counter({ total: 42 });
    const result = Cache.get("miss_key", compute);

    T.assertEquals(1, compute.count, "Compute function should run once.");
    T.assertDeepEquals({ total: 42 }, result, "Should return the computed value.");
    T.assertEquals(JSON.stringify({ total: 42 }), store.fp_miss_key.value, "Script cache should hold the JSON value under the fp_ namespace.");
    T.assertEquals(21600 * 1000, store.fp_miss_key.expiresAt - now, "Default expiry should come from CACHE.EXPIRY_SECONDS.");
  });

  test("get should return value from memory cache on hit", function() {
    Cache.get("mem_key", counter("first"));
    calls = [];
    const compute = counter("second");
    const result = Cache.get("mem_key", compute);

    T.assertEquals("first", result, "Should return the cached value.");
    T.assertEquals(0, compute.count, "Compute function should not run on a memory hit.");
    T.assertEquals(0, calls.length, "Script cache should not be consulted on a memory hit.");
  });

  test("get should return value from script cache on memory miss", function() {
    store.fp_script_key = { value: JSON.stringify({ id: 123 }), expiresAt: now + 10000 };
    const compute = counter("should not compute");
    const result = Cache.get("script_key", compute);

    T.assertDeepEquals({ id: 123 }, result, "Should return the parsed script-cache value.");
    T.assertEquals(0, compute.count, "Compute function should not run on a script hit.");

    calls = [];
    Cache.get("script_key", compute);
    T.assertEquals(0, calls.length, "A script hit should populate the memory cache.");
  });

  test("get should recompute when the memory entry has expired", function() {
    Cache.get("expiring", counter("old"), 60);
    now += 61 * 1000; // both layers expire after 60 s
    const compute = counter("new");
    T.assertEquals("new", Cache.get("expiring", compute, 60), "Should return the recomputed value.");
    T.assertEquals(1, compute.count, "Compute function should run once after expiry.");
  });

  test("get should recompute when the script cache value is corrupt", function() {
    store.fp_corrupt = { value: "{not json", expiresAt: now + 10000 };
    const compute = counter("fresh");
    T.assertEquals("fresh", Cache.get("corrupt", compute), "Corrupt entries should be recomputed.");
    T.assertEquals(1, compute.count, "Compute function should run once.");
  });

  test("get should fall back to computing when the script cache throws", function() {
    global.CacheService = { getScriptCache: () => { throw new Error("Service unavailable"); } };
    const compute = counter("computed");
    T.assertEquals("computed", Cache.get("broken", compute), "Should still return the computed value.");
  });

  test("get should not use cache if disabled", function() {
    Config.update({ CACHE: { ENABLED: false } });
    const compute = counter("direct");
    T.assertEquals("direct", Cache.get("disabled", compute), "Should return the computed value.");
    Cache.get("disabled", compute);
    T.assertEquals(2, compute.count, "Compute function should run on every call when disabled.");
    T.assertEquals(0, calls.length, "Script cache should not be touched when disabled.");
  });

  test("put should populate both caches", function() {
    Cache.put("put_key", { a: 1 }, 120);
    T.assertEquals(JSON.stringify({ a: 1 }), store.fp_put_key.value, "Script cache should hold the value.");
    T.assertEquals(120 * 1000, store.fp_put_key.expiresAt - now, "Explicit expiry should be used.");
    const compute = counter("unused");
    T.assertDeepEquals({ a: 1 }, Cache.get("put_key", compute), "get should return the put value.");
    T.assertEquals(0, compute.count, "Compute function should not run after put.");
  });

  test("put should not populate caches if disabled", function() {
    Config.update({ CACHE: { ENABLED: false } });
    Cache.put("put_disabled", "value");
    T.assertTrue(store.fp_put_disabled === undefined, "Script cache should stay empty when disabled.");
  });

  test("invalidate should remove from both caches", function() {
    Cache.put("inv_key", "value");
    Cache.invalidate("inv_key");
    T.assertTrue(store.fp_inv_key === undefined, "Script cache entry should be removed.");
    const compute = counter("recomputed");
    T.assertEquals("recomputed", Cache.get("inv_key", compute), "Value should be recomputed after invalidate.");
  });

  test("invalidateByPrefix should clear matching memory entries only", function() {
    Cache.put("report_a", 1);
    Cache.put("report_b", 2);
    Cache.put("other", 3);
    // Drop the script layer so only the memory layer can answer.
    store = {};
    Cache.invalidateByPrefix("report_");

    const computeA = counter("A");
    const computeOther = counter("O");
    T.assertEquals("A", Cache.get("report_a", computeA), "Prefixed entry should be recomputed.");
    T.assertEquals(3, Cache.get("other", computeOther), "Non-matching entry should stay in memory.");
    T.assertEquals(0, computeOther.count, "Non-matching entry should not be recomputed.");
  });

  test("invalidateAll should clear memory and known script keys", function() {
    Cache.put(KNOWN_KEY, "value1");
    Cache.put("unknown_key", "value2");
    Cache.invalidateAll();

    T.assertTrue(store['fp_' + KNOWN_KEY] === undefined, "Known script-cache key should be removed.");
    T.assertNotNull(store.fp_unknown_key, "Unknown script-cache keys are not tracked and stay (documented limitation).");
    const compute = counter("recomputed");
    T.assertEquals("recomputed", Cache.get(KNOWN_KEY, compute), "Known key should be recomputed.");
    T.assertEquals("value2", Cache.get("unknown_key", counter("x")), "Unknown key is then served from the script cache.");
  });
})();
