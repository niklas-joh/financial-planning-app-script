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
   * Checks if the cache is enabled in the configuration
   * @return {Boolean} True if caching is enabled
   * @private
   */
  function isCacheEnabled() {
    return config.getSection('CACHE').ENABLED === true;
  }
  
  /**
   * Gets the default cache expiry time in seconds
   * @return {Number} Cache expiry time in seconds
   * @private
   */
  function getDefaultExpirySeconds() {
    return config.getSection('CACHE').EXPIRY_SECONDS || 3600; // Default to 1 hour
  }
  
  /**
   * Generates a cache key with a namespace prefix
   * @param {String} key - The base key
   * @return {String} The namespaced key
   * @private
   */
  function generateNamespacedKey(key) {
    return `fp_${key}`;
  }
  
  // Public API
  return {
    /**
     * Gets a value from cache or computes it if not available
     * @param {String} key - Cache key
     * @param {Function} computeFunction - Function to compute value if not in cache
     * @param {Number} expirySeconds - Cache expiry in seconds (optional)
     * @return {any} The cached or computed value
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
     * Invalidates a specific cache entry
     * @param {String} key - Cache key to invalidate
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
     * Invalidates all cache entries with the given prefix
     * @param {String} prefix - Prefix of keys to invalidate
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
     * Invalidates all cache entries
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
     * Puts a value in the cache
     * @param {String} key - Cache key
     * @param {any} value - Value to cache
     * @param {Number} expirySeconds - Cache expiry in seconds (optional)
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
