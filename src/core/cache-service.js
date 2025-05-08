/**
 * Financial Planning Tools - Cache Service
 * 
 * This file provides a centralized service for caching operations,
 * helping to improve performance by caching expensive operations.
 */

// Create the CacheService module within the FinancialPlanner namespace
FinancialPlanner.CacheService = (function(config) {
  // Private variables
  
  // In-memory cache for ultra-fast access to frequently used data
  const memoryCache = {};
  
  // Private functions
  
  /**
   * Checks if the cache is enabled in the configuration.
   * @return {boolean} True if caching is enabled, false otherwise.
   * @private
   */
  function isCacheEnabled() {
    return config.getSection('CACHE').ENABLED === true;
  }
  
  /**
   * Gets the default cache expiry time in seconds from the configuration.
   * @return {number} Cache expiry time in seconds.
   * @private
   */
  function getDefaultExpirySeconds() {
    return config.getSection('CACHE').EXPIRY_SECONDS || 3600; // Default to 1 hour
  }
  
  /**
   * Generates a cache key with a namespace prefix to avoid collisions.
   * @param {string} key The base key.
   * @return {string} The namespaced key (e.g., "fp_myKey").
   * @private
   */
  function generateNamespacedKey(key) {
    return `fp_${key}`;
  }
  
  // Public API
  return {
    /**
     * Gets a value from the cache. If the value is not found or is expired,
     * it computes the value using the provided function, caches it, and then returns it.
     * If caching is disabled via configuration, it directly calls the computeFunction.
     *
     * @param {string} key The unique key for the cache entry.
     * @param {function(): any} computeFunction A function that computes the value if it's not in the cache.
     *                                        This function should return the value to be cached.
     * @param {number} [expirySeconds] The number of seconds for which the item should be cached.
     *                                 Defaults to the value from `config.getSection('CACHE').EXPIRY_SECONDS` or 3600.
     * @return {any} The cached or computed value.
     *
     * @example
     * const expensiveData = FinancialPlanner.CacheService.get('myDataKey', function() {
     *   return someExpensiveCalculation();
     * }, 600); // Cache for 10 minutes
     *
     * @example
     * // Using default expiry
     * const anotherData = FinancialPlanner.CacheService.get('anotherKey', function() {
     *   return fetchSomeData();
     * });
     */
    get: function(key, computeFunction, expirySeconds) {
      // If caching is disabled, just compute the value
      if (!isCacheEnabled()) {
        return computeFunction();
      }
      
      // Set default expiry if not provided
      if (expirySeconds === undefined) {
        expirySeconds = getDefaultExpirySeconds();
      }
      
      // Generate namespaced key
      const namespacedKey = generateNamespacedKey(key);
      
      // Check memory cache first (fastest)
      if (memoryCache[namespacedKey] && memoryCache[namespacedKey].expiry > Date.now()) {
        return memoryCache[namespacedKey].value;
      }
      
      // Then check script cache
      try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get(namespacedKey);
        
        if (cached != null) {
          try {
            const value = JSON.parse(cached);
            
            // Store in memory cache too for faster access next time
            memoryCache[namespacedKey] = {
              value: value,
              expiry: Date.now() + (expirySeconds * 1000)
            };
            
            return value;
          } catch (parseError) {
            // If parsing fails, treat as cache miss
            console.warn(`Failed to parse cached value for key ${key}:`, parseError);
          }
        }
        
        // Cache miss - compute the value
        const result = computeFunction();
        
        // Store in both caches
        try {
          // Convert to JSON string for storage
          const jsonResult = JSON.stringify(result);
          
          // Store in script cache
          cache.put(namespacedKey, jsonResult, expirySeconds);
          
          // Store in memory cache
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
        // Fall back to direct computation
        return computeFunction();
      }
    },
    
    /**
     * Invalidates (removes) a specific cache entry from both memory and script cache.
     * Does nothing if caching is disabled.
     *
     * @param {string} key The cache key to invalidate.
     *
     * @example
     * FinancialPlanner.CacheService.invalidate('staleDataKey');
     */
    invalidate: function(key) {
      if (!isCacheEnabled()) return;
      
      const namespacedKey = generateNamespacedKey(key);
      
      // Remove from memory cache
      delete memoryCache[namespacedKey];
      
      // Remove from script cache
      try {
        const cache = CacheService.getScriptCache();
        cache.remove(namespacedKey);
      } catch (error) {
        console.warn(`Failed to invalidate cache for key ${key}:`, error);
      }
    },
    
    /**
     * Invalidates all cache entries in the memory cache that start with the given prefix.
     * Note: This currently only affects the in-memory cache due to limitations
     * with Google Apps Script's CacheService prefix removal.
     * Does nothing if caching is disabled.
     *
     * @param {string} prefix The prefix of keys to invalidate (e.g., "user_settings_").
     *
     * @example
     * FinancialPlanner.CacheService.invalidateByPrefix('user_specific_data_');
     */
    invalidateByPrefix: function(prefix) {
      if (!isCacheEnabled()) return;
      
      // Remove from memory cache
      const namespacedPrefix = generateNamespacedKey(prefix);
      
      Object.keys(memoryCache).forEach(key => {
        if (key.startsWith(namespacedPrefix)) {
          delete memoryCache[key];
        }
      });
      
      // Note: CacheService doesn't provide a way to remove by prefix,
      // so we need to track keys with the same prefix separately if needed
    },
    
    /**
     * Invalidates all known cache entries.
     * This clears the entire in-memory cache and attempts to remove known keys
     * (defined in `config.getSection('CACHE').KEYS`) from the script cache.
     * Does nothing if caching is disabled.
     *
     * @example
     * FinancialPlanner.CacheService.invalidateAll();
     */
    invalidateAll: function() {
      if (!isCacheEnabled()) return;
      
      // Clear memory cache
      Object.keys(memoryCache).forEach(key => {
        delete memoryCache[key];
      });
      
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
    
    /**
     * Puts a value directly into the cache (both memory and script cache).
     * If caching is disabled, this operation does nothing.
     *
     * @param {string} key The unique key for the cache entry.
     * @param {any} value The value to cache. Must be serializable to JSON for script cache.
     * @param {number} [expirySeconds] The number of seconds for which the item should be cached.
     *                                 Defaults to the value from `config.getSection('CACHE').EXPIRY_SECONDS` or 3600.
     *
     * @example
     * FinancialPlanner.CacheService.put('userPreferences', { theme: 'dark', notifications: true }, 86400); // Cache for 1 day
     */
    put: function(key, value, expirySeconds) {
      if (!isCacheEnabled()) return;
      
      // Set default expiry if not provided
      if (expirySeconds === undefined) {
        expirySeconds = getDefaultExpirySeconds();
      }
      
      const namespacedKey = generateNamespacedKey(key);
      
      // Store in memory cache
      memoryCache[namespacedKey] = {
        value: value,
        expiry: Date.now() + (expirySeconds * 1000)
      };
      
      // Store in script cache
      try {
        const cache = CacheService.getScriptCache();
        cache.put(namespacedKey, JSON.stringify(value), expirySeconds);
      } catch (error) {
        console.warn(`Failed to put value in cache for key ${key}:`, error);
      }
    }
  };
})(FinancialPlanner.Config);
