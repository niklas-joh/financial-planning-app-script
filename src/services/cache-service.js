/**
 * @fileoverview Cache Service Module for Financial Planning Tools.
 * Provides centralized caching operations to improve performance and reduce
 * reliance on direct API calls or expensive computations. It utilizes both
 * Google Apps Script's `CacheService` and an in-memory cache for faster access.
 * This module is designed to be instantiated by `00_module_loader.js`.
 * @module services/cache-service
 */

/**
 * IIFE to encapsulate the CacheServiceModule logic and prevent global namespace pollution.
 * @returns {function} The CacheServiceModule constructor.
 */
// eslint-disable-next-line no-unused-vars
const CacheServiceModule = (function() {
  /**
   * Constructor for the CacheServiceModule.
   * Initializes the cache service with necessary dependencies and sets up
   * an in-memory cache.
   * @param {ConfigModule} configInstance - An instance of ConfigModule.
   * @param {ErrorServiceModule} errorServiceInstance - An instance of ErrorServiceModule.
   * @constructor
   * @alias CacheServiceModule
   * @memberof module:services/cache-service
   */
  function CacheServiceModuleConstructor(configInstance, errorServiceInstance) {
    /**
     * Instance of ConfigModule, used to access cache configurations.
     * @type {ConfigModule}
     * @private
     */
    this.config = configInstance;
    /**
     * Instance of ErrorServiceModule, used for logging cache-related errors.
     * @type {ErrorServiceModule}
     * @private
     */
    this.errorService = errorServiceInstance;
    /**
     * In-memory cache for ultra-fast access to frequently used items.
     * Stores objects with `value` and `expiry` (timestamp) properties.
     * This cache is per script execution and is faster than `CacheService.getScriptCache()`.
     * @type {Object<string, {value: *, expiry: number}>}
     * @private
     */
    this.memoryCache = {}; // In-memory cache for ultra-fast access
  }

  // Private helper methods (prefixed with _ and attached to prototype or defined in closure)

  /**
   * Checks if caching is globally enabled via configuration.
   * @returns {boolean} True if caching is enabled, false otherwise.
   * @private
   * @memberof CacheServiceModule
   */
  CacheServiceModuleConstructor.prototype._isCacheEnabled = function() {
    try {
      return this.config.getSection('CACHE').ENABLED === true;
    } catch (e) {
      this.errorService.log(this.errorService.create('Failed to read CACHE.ENABLED config', { originalError: e, severity: 'medium' }));
      return false; // Default to cache disabled if config is broken
    }
  };

  /**
   * Retrieves the default cache expiry time from configuration.
   * Defaults to 1 hour (3600 seconds) if not specified or if config is broken.
   * @returns {number} The default cache expiry time in seconds.
   * @private
   * @memberof CacheServiceModule
   */
  CacheServiceModuleConstructor.prototype._getDefaultExpirySeconds = function() {
    try {
      return this.config.getSection('CACHE').EXPIRY_SECONDS || 3600; // Default to 1 hour
    } catch (e) {
      this.errorService.log(this.errorService.create('Failed to read CACHE.EXPIRY_SECONDS config', { originalError: e, severity: 'medium' }));
      return 3600; // Default to 1 hour if config is broken
    }
  };

  /**
   * Generates a namespaced key to prevent collisions in the global cache.
   * @param {string} key - The original cache key.
   * @returns {string} The namespaced cache key (e.g., "fp_myKey").
   * @private
   * @memberof CacheServiceModule
   */
  CacheServiceModuleConstructor.prototype._generateNamespacedKey = function(key) {
    return `fp_${key}`;
  };

  // Public API methods

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
   * @memberof CacheServiceModule
   */
  CacheServiceModuleConstructor.prototype.get = function(key, computeFunction, expirySeconds) {
    if (!this._isCacheEnabled()) {
      return computeFunction();
    }

    const effectiveExpirySeconds = expirySeconds === undefined ? this._getDefaultExpirySeconds() : expirySeconds;
    const namespacedKey = this._generateNamespacedKey(key);

    if (this.memoryCache[namespacedKey] && this.memoryCache[namespacedKey].expiry > Date.now()) {
      return this.memoryCache[namespacedKey].value;
    }

    try {
      const scriptCache = CacheService.getScriptCache();
      const cached = scriptCache.get(namespacedKey);

      if (cached != null) {
        try {
          const value = JSON.parse(cached);
          this.memoryCache[namespacedKey] = {
            value: value,
            expiry: Date.now() + (effectiveExpirySeconds * 1000),
          };
          return value;
        } catch (parseError) {
          this.errorService.log(this.errorService.create(`Failed to parse cached value for key ${key}`, { originalError: parseError, severity: 'warning' }));
        }
      }

      const result = computeFunction();
      try {
        const jsonResult = JSON.stringify(result);
        scriptCache.put(namespacedKey, jsonResult, effectiveExpirySeconds);
        this.memoryCache[namespacedKey] = {
          value: result,
          expiry: Date.now() + (effectiveExpirySeconds * 1000),
        };
      } catch (cachePutError) {
        this.errorService.log(this.errorService.create(`Failed to cache result for key ${key}`, { originalError: cachePutError, severity: 'warning' }));
      }
      return result;
    } catch (error) {
      this.errorService.log(this.errorService.create(`Cache 'get' operation failed for key ${key}`, { originalError: error, severity: 'warning' }));
      return computeFunction(); // Fall back to direct computation
    }
  };

  /**
   * Stores an item in the cache (both in-memory and script cache).
   * The value must be JSON-serializable for script cache.
   *
   * @param {string} key - The unique key for the cache item.
   * @param {*} value - The value to be cached. It will be JSON.stringified for script cache.
   * @param {number} [expirySeconds] - Optional. The expiry time in seconds for this item.
   *   If not provided, the default cache expiry time from config is used.
   * @returns {void}
   * @memberof CacheServiceModule
   */
  CacheServiceModuleConstructor.prototype.put = function(key, value, expirySeconds) {
    if (!this._isCacheEnabled()) return;

    const effectiveExpirySeconds = expirySeconds === undefined ? this._getDefaultExpirySeconds() : expirySeconds;
    const namespacedKey = this._generateNamespacedKey(key);

    this.memoryCache[namespacedKey] = {
      value: value,
      expiry: Date.now() + (effectiveExpirySeconds * 1000),
    };

    try {
      const scriptCache = CacheService.getScriptCache();
      scriptCache.put(namespacedKey, JSON.stringify(value), effectiveExpirySeconds);
    } catch (error) {
      this.errorService.log(this.errorService.create(`Failed to put value in cache for key ${key}`, { originalError: error, severity: 'warning' }));
    }
  };

  /**
   * Removes a specific item from both the in-memory and script cache.
   *
   * @param {string} key - The key of the item to remove.
   * @returns {void}
   * @memberof CacheServiceModule
   */
  CacheServiceModuleConstructor.prototype.invalidate = function(key) {
    if (!this._isCacheEnabled()) return;
    const namespacedKey = this._generateNamespacedKey(key);

    delete this.memoryCache[namespacedKey];

    try {
      const scriptCache = CacheService.getScriptCache();
      scriptCache.remove(namespacedKey);
    } catch (error) {
      this.errorService.log(this.errorService.create(`Failed to invalidate cache for key ${key}`, { originalError: error, severity: 'warning' }));
    }
  };

  /**
   * Invalidates cache entries where the key starts with the given prefix.
   * **Important:** Due to Google Apps Script `CacheService` limitations, this method
   * only reliably clears matching entries from the **in-memory cache**.
   * It does not perform a prefix-based removal from the script cache.
   *
   * @param {string} prefix - The prefix of keys to invalidate from the in-memory cache.
   * @returns {void}
   * @memberof CacheServiceModule
   */
  CacheServiceModuleConstructor.prototype.invalidateByPrefix = function(prefix) {
    if (!this._isCacheEnabled()) return;
    const namespacedPrefix = this._generateNamespacedKey(prefix);

    Object.keys(this.memoryCache).forEach(key => {
      if (key.startsWith(namespacedPrefix)) {
        delete this.memoryCache[key];
      }
    });
    // Note: Script cache does not support removeByPrefix easily.
    // This could be logged as a known limitation or handled if specific keys are tracked.
    this.errorService.log(this.errorService.create('invalidateByPrefix only affects memory cache due to Apps Script CacheService limitations.', { severity: 'info', prefix: namespacedPrefix }));
  };

  /**
   * Clears all items from the in-memory cache.
   * Attempts to clear all **known** items (defined in `CONFIG.CACHE.KEYS`) from the script cache.
   * Due to Apps Script limitations, a true "clear all" for script cache is not
   * straightforward without iterating over all possible keys, which is not done here.
   *
   * @returns {void}
   * @memberof CacheServiceModule
   */
  CacheServiceModuleConstructor.prototype.invalidateAll = function() {
    if (!this._isCacheEnabled()) return;

    Object.keys(this.memoryCache).forEach(key => {
      delete this.memoryCache[key];
    });

    try {
      const scriptCache = CacheService.getScriptCache();
      const cacheConfig = this.config.getSection('CACHE');
      const knownKeys = cacheConfig && cacheConfig.KEYS ? Object.values(cacheConfig.KEYS) : [];
      
      const namespacedKeysToRemove = knownKeys.map(key => this._generateNamespacedKey(key));

      if (namespacedKeysToRemove.length > 0) {
        scriptCache.removeAll(namespacedKeysToRemove);
      }
      // To clear truly *all* script cache, one would have to iterate and remove,
      // or accept that only known keys are cleared. This implementation clears known keys.
    } catch (error) {
      this.errorService.log(this.errorService.create('Failed to invalidate all known cache entries from script cache', { originalError: error, severity: 'warning' }));
    }
  };

  return CacheServiceModuleConstructor;
})();
