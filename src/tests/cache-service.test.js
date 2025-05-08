/**
 * Financial Planning Tools - Cache Service Tests
 *
 * This file contains tests for the FinancialPlanner.CacheService module.
 * It includes mocking Google Apps Script's CacheService.
 */
(function() {
  // Alias for easier access
  const T = FinancialPlanner.Testing;

  // --- Mock Dependencies & Globals ---
  let mockScriptCacheStore = {};
  let mockMemoryCacheStore = {}; // Simulate internal memory cache if needed for verification
  let scriptCacheGetError = null;
  let scriptCachePutError = null;
  let scriptCacheRemoveError = null;
  let scriptCacheRemoveAllError = null;
  let jsonParseError = null;
  let jsonStringifyError = null;

  const mockScriptCache = {
    get: function(key) {
      if (scriptCacheGetError) throw scriptCacheGetError;
      // Simulate expiry (simplified) - real CacheService handles this internally
      const item = mockScriptCacheStore[key];
      if (item && item.expiry && item.expiry < Date.now()) {
          // console.log(`Mock Cache: Key '${key}' expired.`);
          delete mockScriptCacheStore[key];
          return null;
      }
      // console.log(`Mock Cache: Getting key '${key}', found: ${item ? item.value : 'null'}`);
      return item ? item.value : null;
    },
    put: function(key, value, ttl) {
      if (scriptCachePutError) throw scriptCachePutError;
      const expiry = ttl ? Date.now() + ttl * 1000 : null;
      // console.log(`Mock Cache: Putting key '${key}' with value '${value}', ttl: ${ttl}, expiry: ${expiry}`);
      mockScriptCacheStore[key] = { value: value, expiry: expiry };
    },
    remove: function(key) {
      if (scriptCacheRemoveError) throw scriptCacheRemoveError;
      // console.log(`Mock Cache: Removing key '${key}'`);
      delete mockScriptCacheStore[key];
    },
    removeAll: function(keys) {
       if (scriptCacheRemoveAllError) throw scriptCacheRemoveAllError;
       // console.log(`Mock Cache: Removing keys: ${keys.join(', ')}`);
       keys.forEach(key => delete mockScriptCacheStore[key]);
    }
  };

  // Global mock for CacheService
  global.CacheService = {
    getScriptCache: function() {
      return mockScriptCache;
    }
    // getDocumentCache, getUserCache etc. could be mocked if needed
  };
  
  // Mock JSON methods to simulate errors if needed
  const originalJSONParse = JSON.parse;
  const originalJSONStringify = JSON.stringify;

  global.JSON.parse = function(text) {
      if (jsonParseError) throw jsonParseError;
      return originalJSONParse(text);
  };
   global.JSON.stringify = function(value) {
      if (jsonStringifyError) throw jsonStringifyError;
      return originalJSONStringify(value);
  };

  // Mock Config dependency
  let mockCacheEnabled = true;
  const mockConfig = {
      _cacheSettings: {
          ENABLED: true,
          EXPIRY_SECONDS: 600, // 10 minutes default for tests
          KEYS: {
              KEY1: "test_key_1",
              KEY2: "test_key_2"
          }
      },
      getSection: function(section) {
          if (section === 'CACHE') {
              // Return a copy to prevent tests modifying the mock config directly
              const settings = { ...this._cacheSettings };
              settings.ENABLED = mockCacheEnabled; // Use the controllable flag
              return settings;
          }
          return {};
      },
      // Add get() if CacheService uses it directly
      get: function() {
          return { CACHE: this.getSection('CACHE') };
      }
  };

  // --- Test Suite Setup ---
   // Redefine the service for testing, injecting mocks
   // Need to re-run the IIFE with mocks
   const TestCacheService = (function(config) {
       // --- Copy of CacheService Implementation Start ---
        const memoryCache = {}; // Use a local var for the test instance's memory cache
        
        function isCacheEnabled() {
            // Use the provided mock config's current state
            return config.getSection('CACHE').ENABLED === true;
        }
        
        function getDefaultExpirySeconds() {
            return config.getSection('CACHE').EXPIRY_SECONDS || 3600;
        }
        
        function generateNamespacedKey(key) {
            return `fp_${key}`;
        }
        
        return {
            get: function(key, computeFunction, expirySeconds) {
            if (!isCacheEnabled()) {
                return computeFunction();
            }
            
            if (expirySeconds === undefined) {
                expirySeconds = getDefaultExpirySeconds();
            }
            
            const namespacedKey = generateNamespacedKey(key);
            
            // Check memory cache first
            if (memoryCache[namespacedKey] && memoryCache[namespacedKey].expiry > Date.now()) {
                // console.log(`TestCacheService: Memory cache hit for ${key}`);
                return memoryCache[namespacedKey].value;
            }
            
            // Then check script cache
            try {
                const cache = CacheService.getScriptCache(); // Uses global mock
                const cached = cache.get(namespacedKey);
                
                if (cached != null) {
                try {
                    const value = JSON.parse(cached); // Uses global mock JSON
                    // console.log(`TestCacheService: Script cache hit for ${key}`);
                    memoryCache[namespacedKey] = {
                    value: value,
                    expiry: Date.now() + (expirySeconds * 1000)
                    };
                    return value;
                } catch (parseError) {
                    console.warn(`Failed to parse cached value for key ${key}:`, parseError);
                    // Fall through to compute
                }
                }
                
                // Cache miss - compute
                // console.log(`TestCacheService: Cache miss for ${key}, computing...`);
                const result = computeFunction();
                
                try {
                const jsonResult = JSON.stringify(result); // Uses global mock JSON
                cache.put(namespacedKey, jsonResult, expirySeconds);
                memoryCache[namespacedKey] = {
                    value: result,
                    expiry: Date.now() + (expirySeconds * 1000)
                };
                } catch (cacheError) {
                console.warn(`Failed to cache result for key ${key}:`, cacheError);
                }
                
                return result;
            } catch (error) {
                console.warn(`Cache operation failed for key ${key}:`, error);
                return computeFunction();
            }
            },
            
            put: function(key, value, expirySeconds) {
                if (!isCacheEnabled()) return;
                
                if (expirySeconds === undefined) {
                    expirySeconds = getDefaultExpirySeconds();
                }
                
                const namespacedKey = generateNamespacedKey(key);
                
                memoryCache[namespacedKey] = {
                    value: value,
                    expiry: Date.now() + (expirySeconds * 1000)
                };
                
                try {
                    const cache = CacheService.getScriptCache();
                    cache.put(namespacedKey, JSON.stringify(value), expirySeconds);
                } catch (error) {
                    console.warn(`Failed to put value in cache for key ${key}:`, error);
                }
            },

            invalidate: function(key) {
                if (!isCacheEnabled()) return;
                const namespacedKey = generateNamespacedKey(key);
                delete memoryCache[namespacedKey];
                try {
                    const cache = CacheService.getScriptCache();
                    cache.remove(namespacedKey);
                } catch (error) {
                    console.warn(`Failed to invalidate cache for key ${key}:`, error);
                }
            },
            
            invalidateByPrefix: function(prefix) {
                if (!isCacheEnabled()) return;
                const namespacedPrefix = generateNamespacedKey(prefix);
                Object.keys(memoryCache).forEach(key => {
                    if (key.startsWith(namespacedPrefix)) {
                    delete memoryCache[key];
                    }
                });
                // Script cache invalidation by prefix is not directly supported by mock/GAS CacheService
            },
            
            invalidateAll: function() {
                if (!isCacheEnabled()) return;
                // Clear memory cache
                Object.keys(memoryCache).forEach(key => delete memoryCache[key]);
                // Clear script cache for known keys
                try {
                    const cache = CacheService.getScriptCache();
                    const keys = Object.values(config.getSection('CACHE').KEYS || {})
                    .map(key => generateNamespacedKey(key));
                    if (keys.length > 0) {
                    cache.removeAll(keys);
                    }
                } catch (error) {
                    console.warn("Failed to invalidate all cache entries:", error);
                }
            },
             // Helper for tests to inspect memory cache
            _getMemoryCache: function() { return memoryCache; }
        };
       // --- Copy of CacheService Implementation End ---
   })(mockConfig); // Pass mock config


  // --- Helper to reset state before each test ---
  function setupTestState() {
      mockScriptCacheStore = {};
      // Reset the internal memory cache of the TestCacheService instance
      TestCacheService.invalidateAll(); // This should clear its internal memory cache
      scriptCacheGetError = null;
      scriptCachePutError = null;
      scriptCacheRemoveError = null;
      scriptCacheRemoveAllError = null;
      jsonParseError = null;
      jsonStringifyError = null;
      mockCacheEnabled = true; // Default to enabled
  }

  // --- Test Cases ---

  T.registerTest("CacheService", "get should compute value on cache miss", function() {
    setupTestState();
    let computeCount = 0;
    const key = "miss_key";
    const expectedValue = { data: "computed" };
    const computeFunc = () => { computeCount++; return expectedValue; };

    const result = TestCacheService.get(key, computeFunc);

    T.assertEquals(1, computeCount, "Compute function should be called once on miss.");
    T.assertDeepEquals(expectedValue, result, "Should return the computed value.");
    // Check script cache was populated
    const namespacedKey = `fp_${key}`;
    T.assertNotNull(mockScriptCacheStore[namespacedKey], "Script cache should contain the key.");
    T.assertEquals(JSON.stringify(expectedValue), mockScriptCacheStore[namespacedKey].value, "Script cache should contain the stringified value.");
     // Check memory cache was populated
    T.assertNotNull(TestCacheService._getMemoryCache()[namespacedKey], "Memory cache should contain the key.");
    T.assertDeepEquals(expectedValue, TestCacheService._getMemoryCache()[namespacedKey].value, "Memory cache should contain the original value.");
  });

  T.registerTest("CacheService", "get should return value from memory cache on hit", function() {
     setupTestState();
     let computeCount = 0;
     const key = "mem_hit_key";
     const expectedValue = "from memory";
     const computeFunc = () => { computeCount++; return "should not compute"; };
     const namespacedKey = `fp_${key}`;

     // Pre-populate memory cache
     TestCacheService._getMemoryCache()[namespacedKey] = { value: expectedValue, expiry: Date.now() + 10000 };

     const result = TestCacheService.get(key, computeFunc);

     T.assertEquals(0, computeCount, "Compute function should not be called on memory hit.");
     T.assertEquals(expectedValue, result, "Should return the value from memory cache.");
  });
  
   T.registerTest("CacheService", "get should return value from script cache on hit (memory miss)", function() {
     setupTestState();
     let computeCount = 0;
     const key = "script_hit_key";
     const expectedValue = { id: 123, status: "active" };
     const computeFunc = () => { computeCount++; return "should not compute"; };
     const namespacedKey = `fp_${key}`;

     // Pre-populate script cache
     mockScriptCacheStore[namespacedKey] = { value: JSON.stringify(expectedValue), expiry: Date.now() + 10000 };
     // Ensure memory cache is empty for this key
     delete TestCacheService._getMemoryCache()[namespacedKey];

     const result = TestCacheService.get(key, computeFunc);

     T.assertEquals(0, computeCount, "Compute function should not be called on script hit.");
     T.assertDeepEquals(expectedValue, result, "Should return the value from script cache.");
      // Check memory cache was populated after script hit
     T.assertNotNull(TestCacheService._getMemoryCache()[namespacedKey], "Memory cache should be populated after script hit.");
     T.assertDeepEquals(expectedValue, TestCacheService._getMemoryCache()[namespacedKey].value, "Memory cache should contain the value after script hit.");
  });

  T.registerTest("CacheService", "get should recompute if memory cache expired", function() {
    setupTestState();
    let computeCount = 0;
    const key = "mem_expired_key";
    const initialValue = "expired value";
    const newValue = "newly computed";
    const computeFunc = () => { computeCount++; return newValue; };
    const namespacedKey = `fp_${key}`;

    // Pre-populate memory cache with expired item
    TestCacheService._getMemoryCache()[namespacedKey] = { value: initialValue, expiry: Date.now() - 10000 }; // Expired 10s ago

    const result = TestCacheService.get(key, computeFunc);

    T.assertEquals(1, computeCount, "Compute function should be called once when memory cache expired.");
    T.assertEquals(newValue, result, "Should return the newly computed value.");
  });
  
   T.registerTest("CacheService", "get should recompute if script cache expired", function() {
    setupTestState();
    let computeCount = 0;
    const key = "script_expired_key";
    const initialValue = "expired script value";
    const newValue = "newly computed again";
    const computeFunc = () => { computeCount++; return newValue; };
    const namespacedKey = `fp_${key}`;

    // Pre-populate script cache with expired item
    mockScriptCacheStore[namespacedKey] = { value: JSON.stringify(initialValue), expiry: Date.now() - 10000 }; // Expired
     // Ensure memory cache is empty
     delete TestCacheService._getMemoryCache()[namespacedKey];

    const result = TestCacheService.get(key, computeFunc);

    T.assertEquals(1, computeCount, "Compute function should be called once when script cache expired.");
    T.assertEquals(newValue, result, "Should return the newly computed value.");
  });

  T.registerTest("CacheService", "get should not use cache if disabled", function() {
    setupTestState();
    mockCacheEnabled = false; // Disable cache via mock config
    let computeCount = 0;
    const key = "disabled_key";
    const expectedValue = "computed when disabled";
    const computeFunc = () => { computeCount++; return expectedValue; };
    const namespacedKey = `fp_${key}`;

    // Pre-populate caches
    TestCacheService._getMemoryCache()[namespacedKey] = { value: "memory value", expiry: Date.now() + 10000 };
    mockScriptCacheStore[namespacedKey] = { value: JSON.stringify("script value"), expiry: Date.now() + 10000 };

    const result = TestCacheService.get(key, computeFunc);

    T.assertEquals(1, computeCount, "Compute function should always be called when cache is disabled.");
    T.assertEquals(expectedValue, result, "Should return the computed value when cache is disabled.");
    // Optionally check caches weren't modified (though put also checks isCacheEnabled)
  });

  T.registerTest("CacheService", "put should populate both caches", function() {
      setupTestState();
      const key = "put_key";
      const value = { item: "value to put" };
      const namespacedKey = `fp_${key}`;

      TestCacheService.put(key, value);

      // Check memory cache
      T.assertNotNull(TestCacheService._getMemoryCache()[namespacedKey], "Memory cache should contain the key after put.");
      T.assertDeepEquals(value, TestCacheService._getMemoryCache()[namespacedKey].value, "Memory cache value should match after put.");
      // Check script cache
      T.assertNotNull(mockScriptCacheStore[namespacedKey], "Script cache should contain the key after put.");
      T.assertEquals(JSON.stringify(value), mockScriptCacheStore[namespacedKey].value, "Script cache value should match after put.");
  });
  
   T.registerTest("CacheService", "put should not populate caches if disabled", function() {
      setupTestState();
      mockCacheEnabled = false;
      const key = "put_disabled_key";
      const value = "no cache put";
      const namespacedKey = `fp_${key}`;

      TestCacheService.put(key, value);

      T.assertTrue(TestCacheService._getMemoryCache()[namespacedKey] === undefined, "Memory cache should be empty after put when disabled.");
      T.assertTrue(mockScriptCacheStore[namespacedKey] === undefined, "Script cache should be empty after put when disabled.");
  });

  T.registerTest("CacheService", "invalidate should remove from both caches", function() {
      setupTestState();
      const key = "invalidate_key";
      const value = "to be invalidated";
      const namespacedKey = `fp_${key}`;

      // Populate caches
      TestCacheService.put(key, value);
      T.assertNotNull(TestCacheService._getMemoryCache()[namespacedKey], "Memory cache should have key before invalidate.");
      T.assertNotNull(mockScriptCacheStore[namespacedKey], "Script cache should have key before invalidate.");

      TestCacheService.invalidate(key);

      T.assertTrue(TestCacheService._getMemoryCache()[namespacedKey] === undefined, "Memory cache should be empty after invalidate.");
      T.assertTrue(mockScriptCacheStore[namespacedKey] === undefined, "Script cache should be empty after invalidate.");
  });
  
   T.registerTest("CacheService", "invalidateAll should clear memory and known script keys", function() {
      setupTestState();
      const key1 = mockConfig._cacheSettings.KEYS.KEY1; // Known key
      const key2 = "unknown_key";
      const nsKey1 = `fp_${key1}`;
      const nsKey2 = `fp_${key2}`;

      // Populate caches
      TestCacheService.put(key1, "value1");
      TestCacheService.put(key2, "value2"); // Put an unknown key too
      
      T.assertNotNull(TestCacheService._getMemoryCache()[nsKey1], "Memory cache should have key1 before invalidateAll.");
      T.assertNotNull(TestCacheService._getMemoryCache()[nsKey2], "Memory cache should have key2 before invalidateAll.");
      T.assertNotNull(mockScriptCacheStore[nsKey1], "Script cache should have key1 before invalidateAll.");
       T.assertNotNull(mockScriptCacheStore[nsKey2], "Script cache should have key2 before invalidateAll.");


      TestCacheService.invalidateAll();

      T.assertTrue(TestCacheService._getMemoryCache()[nsKey1] === undefined, "Memory cache key1 should be empty after invalidateAll.");
       T.assertTrue(TestCacheService._getMemoryCache()[nsKey2] === undefined, "Memory cache key2 should be empty after invalidateAll.");
      T.assertTrue(mockScriptCacheStore[nsKey1] === undefined, "Script cache known key1 should be empty after invalidateAll.");
      // Note: The mock invalidateAll only removes *known* keys from script cache
      T.assertNotNull(mockScriptCacheStore[nsKey2], "Script cache unknown key2 should *not* be removed by invalidateAll.");
  });


})(); // End IIFE
