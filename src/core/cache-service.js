/**
 * @fileoverview Cache Service Module for Financial Planning Tools.
 * Provides centralized caching operations to improve performance.
 * This module is designed to be instantiated by 00_module_loader.js.
 */

// eslint-disable-next-line no-unused-vars
const CacheServiceModule = (function() {
  /**
   * Constructor for the CacheServiceModule.
   * @param {object} configInstance - An instance of ConfigModule.
   * @param {object} errorServiceInstance - An instance of ErrorServiceModule.
   * @constructor
   */
  function CacheServiceModuleConstructor(configInstance, errorServiceInstance) {
    this.config = configInstance;
    this.errorService = errorServiceInstance;
    this.memoryCache = {}; // In-memory cache for ultra-fast access
  }

  // Private helper methods (prefixed with _ and attached to prototype or defined in closure)

  CacheServiceModuleConstructor.prototype._isCacheEnabled = function() {
    try {
      return this.config.getSection('CACHE').ENABLED === true;
    } catch (e) {
      this.errorService.log(this.errorService.create('Failed to read CACHE.ENABLED config', { originalError: e, severity: 'medium' }));
      return false; // Default to cache disabled if config is broken
    }
  };

  CacheServiceModuleConstructor.prototype._getDefaultExpirySeconds = function() {
    try {
      return this.config.getSection('CACHE').EXPIRY_SECONDS || 3600; // Default to 1 hour
    } catch (e) {
      this.errorService.log(this.errorService.create('Failed to read CACHE.EXPIRY_SECONDS config', { originalError: e, severity: 'medium' }));
      return 3600; // Default to 1 hour if config is broken
    }
  };

  CacheServiceModuleConstructor.prototype._generateNamespacedKey = function(key) {
    return `fp_${key}`;
  };

  // Public API methods

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
