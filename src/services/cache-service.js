/**
 * @fileoverview Cache Service Module for Financial Planning Tools.
 * Provides centralized caching operations to improve performance and reduce
 * reliance on direct API calls or expensive computations. It utilizes both
 * Google Apps Script's `CacheService` and an in-memory cache for faster access.
 * @module services/cache-service
 */

// Ensure the global FinancialPlanner namespace exists
// eslint-disable-next-line no-var, vars-on-top
var FinancialPlanner = FinancialPlanner || {};

/**
 * Cache Service - Provides centralized caching with in-memory and script cache.
 * Uses IIFE to keep memoryCache private via closure.
 * @namespace FinancialPlanner.CacheService
 */
FinancialPlanner.CacheService = (function() {
  /**
   * In-memory cache for ultra-fast access to frequently used items.
   * Stores objects with `value` and `expiry` (timestamp) properties.
   * This cache is per script execution and is faster than `CacheService.getScriptCache()`.
   * @type {Object<string, {value: *, expiry: number}>}
   * @private
   */
  const memoryCache = {};

  /**
   * Checks if caching is globally enabled via configuration.
   * @returns {boolean} True if caching is enabled, false otherwise.
   * @private
   */
  function isCacheEnabled() {
    try {
      return FinancialPlanner.Config.getSection('CACHE').ENABLED === true;
    } catch (e) {
      FinancialPlanner.ErrorService.log(
        FinancialPlanner.ErrorService.create('Failed to read CACHE.ENABLED config', { originalError: e, severity: 'medium' })
      );
      return false; // Default to cache disabled if config is broken
    }
  }

  /**
   * Retrieves the default cache expiry time from configuration.
   * Defaults to 1 hour (3600 seconds) if not specified or if config is broken.
   * @returns {number} The default cache expiry time in seconds.
   * @private
   */
  function getDefaultExpirySeconds() {
    try {
      return FinancialPlanner.Config.getSection('CACHE').EXPIRY_SECONDS || 3600; // Default to 1 hour
    } catch (e) {
      FinancialPlanner.ErrorService.log(
        FinancialPlanner.ErrorService.create('Failed to read CACHE.EXPIRY_SECONDS config', { originalError: e, severity: 'medium' })
      );
      return 3600; // Default to 1 hour if config is broken
    }
  }

  /**
   * Generates a namespaced key to prevent collisions in the global cache.
   * @param {string} key - The original cache key.
   * @returns {string} The namespaced cache key (e.g., "fp_myKey").
   * @private
   */
  function generateNamespacedKey(key) {
    return 'fp_' + key;
  }

  // Public API
  return {
    /**
     * Retrieves an item from the cache. If the item is not found or expired,
     * it executes the `computeFunction` to generate the value, caches it,
     * and then returns it. Handles both in-memory and script cache.
     *
     * @param {string} key - The unique key for the cache item.
     * @param {function(): *} computeFunction - A function that computes/retrieves the value
     *   if it's not found in the cache or is expired. This function should return the value.
     * @param {number} [expirySeconds] - Optional. The expiry time in seconds for this specific item.
     *   If not provided, the default cache expiry time from config is used.
     * @returns {*} The cached value or the newly computed value. Returns the result of
     *   `computeFunction` directly if caching is disabled or if errors occur during caching.
     * @memberof FinancialPlanner.CacheService
     */
    get: function(key, computeFunction, expirySeconds) {
      if (!isCacheEnabled()) {
        return computeFunction();
      }

      const effectiveExpirySeconds = expirySeconds !== undefined ? expirySeconds : getDefaultExpirySeconds();
      const namespacedKey = generateNamespacedKey(key);

      if (memoryCache[namespacedKey] && memoryCache[namespacedKey].expiry > Date.now()) {
        return memoryCache[namespacedKey].value;
      }

      try {
        const scriptCache = CacheService.getScriptCache();
        const cached = scriptCache.get(namespacedKey);

        if (cached != null) {
          try {
            const value = JSON.parse(cached);
            memoryCache[namespacedKey] = {
              value: value,
              expiry: Date.now() + (effectiveExpirySeconds * 1000),
            };
            return value;
          } catch (parseError) {
            FinancialPlanner.ErrorService.log(
              FinancialPlanner.ErrorService.create('Failed to parse cached value for key ' + key, { originalError: parseError, severity: 'warning' })
            );
          }
        }

        const result = computeFunction();
        try {
          const jsonResult = JSON.stringify(result);
          scriptCache.put(namespacedKey, jsonResult, effectiveExpirySeconds);
          memoryCache[namespacedKey] = {
            value: result,
            expiry: Date.now() + (effectiveExpirySeconds * 1000),
          };
        } catch (cachePutError) {
          FinancialPlanner.ErrorService.log(
            FinancialPlanner.ErrorService.create('Failed to cache result for key ' + key, { originalError: cachePutError, severity: 'warning' })
          );
        }
        return result;
      } catch (error) {
        FinancialPlanner.ErrorService.log(
          FinancialPlanner.ErrorService.create('Cache get operation failed for key ' + key, { originalError: error, severity: 'warning' })
        );
        return computeFunction(); // Fall back to direct computation
      }
    },

    /**
     * Stores an item in the cache (both in-memory and script cache).
     * The value must be JSON-serializable for script cache.
     *
     * @param {string} key - The unique key for the cache item.
     * @param {*} value - The value to be cached. It will be JSON.stringified for script cache.
     * @param {number} [expirySeconds] - Optional. The expiry time in seconds for this item.
     *   If not provided, the default cache expiry time from config is used.
     * @returns {void}
     * @memberof FinancialPlanner.CacheService
     */
    put: function(key, value, expirySeconds) {
      if (!isCacheEnabled()) return;

      const effectiveExpirySeconds = expirySeconds !== undefined ? expirySeconds : getDefaultExpirySeconds();
      const namespacedKey = generateNamespacedKey(key);

      memoryCache[namespacedKey] = {
        value: value,
        expiry: Date.now() + (effectiveExpirySeconds * 1000),
      };

      try {
        const scriptCache = CacheService.getScriptCache();
        scriptCache.put(namespacedKey, JSON.stringify(value), effectiveExpirySeconds);
      } catch (error) {
        FinancialPlanner.ErrorService.log(
          FinancialPlanner.ErrorService.create('Failed to put value in cache for key ' + key, { originalError: error, severity: 'warning' })
        );
      }
    },

    /**
     * Removes a specific item from both the in-memory and script cache.
     *
     * @param {string} key - The key of the item to remove.
     * @returns {void}
     * @memberof FinancialPlanner.CacheService
     */
    invalidate: function(key) {
      if (!isCacheEnabled()) return;
      const namespacedKey = generateNamespacedKey(key);

      delete memoryCache[namespacedKey];

      try {
        const scriptCache = CacheService.getScriptCache();
        scriptCache.remove(namespacedKey);
      } catch (error) {
        FinancialPlanner.ErrorService.log(
          FinancialPlanner.ErrorService.create('Failed to invalidate cache for key ' + key, { originalError: error, severity: 'warning' })
        );
      }
    },

    /**
     * Invalidates cache entries where the key starts with the given prefix.
     * **Important:** Due to Google Apps Script `CacheService` limitations, this method
     * only reliably clears matching entries from the **in-memory cache**.
     * It does not perform a prefix-based removal from the script cache.
     *
     * @param {string} prefix - The prefix of keys to invalidate from the in-memory cache.
     * @returns {void}
     * @memberof FinancialPlanner.CacheService
     */
    invalidateByPrefix: function(prefix) {
      if (!isCacheEnabled()) return;
      const namespacedPrefix = generateNamespacedKey(prefix);

      Object.keys(memoryCache).forEach(function(key) {
        if (key.indexOf(namespacedPrefix) === 0) {
          delete memoryCache[key];
        }
      });
      // Note: Script cache does not support removeByPrefix easily.
      // This could be logged as a known limitation or handled if specific keys are tracked.
      FinancialPlanner.ErrorService.log(
        FinancialPlanner.ErrorService.create('invalidateByPrefix only affects memory cache due to Apps Script CacheService limitations.', { severity: 'info', prefix: namespacedPrefix })
      );
    },

    /**
     * Clears all items from the in-memory cache.
     * Attempts to clear all **known** items (defined in `CONFIG.CACHE.KEYS`) from the script cache.
     * Due to Apps Script limitations, a true "clear all" for script cache is not
     * straightforward without iterating over all possible keys, which is not done here.
     *
     * @returns {void}
     * @memberof FinancialPlanner.CacheService
     */
    invalidateAll: function() {
      if (!isCacheEnabled()) return;

      Object.keys(memoryCache).forEach(function(key) {
        delete memoryCache[key];
      });

      try {
        const scriptCache = CacheService.getScriptCache();
        const cacheConfig = FinancialPlanner.Config.getSection('CACHE');
        const knownKeys = cacheConfig && cacheConfig.KEYS ? Object.values(cacheConfig.KEYS) : [];
        
        const namespacedKeysToRemove = knownKeys.map(function(key) {
          return generateNamespacedKey(key);
        });

        if (namespacedKeysToRemove.length > 0) {
          scriptCache.removeAll(namespacedKeysToRemove);
        }
        // To clear truly *all* script cache, one would have to iterate and remove,
        // or accept that only known keys are cleared. This implementation clears known keys.
      } catch (error) {
        FinancialPlanner.ErrorService.log(
          FinancialPlanner.ErrorService.create('Failed to invalidate all known cache entries from script cache', { originalError: error, severity: 'warning' })
        );
      }
    }
  };
})();
