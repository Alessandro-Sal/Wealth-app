/**
 * Generic caching wrapper implementing the "Cache-Aside" pattern.
 * Uses chunking to bypass the 100KB value limit of CacheService.
 * @param {string} key - The unique cache key.
 * @param {Function} fetchFunction - The callback to execute if cache misses (slow operation).
 * @param {number} [expirationSec=600] - Cache duration in seconds (Default: 600s / 10 min).
 * @return {Object|null} The parsed data or null on error.
 */
function getFromCache(key, fetchFunction, expirationSec = 600) {
  const cache = CacheService.getScriptCache();
  
  // Read chunk count
  const chunksStr = cache.get(key + '_chunks');
  let cachedStr = null;
  
  if (chunksStr) {
    const chunks = parseInt(chunksStr, 10);
    if (chunks === 1) {
      cachedStr = cache.get(key);
    } else {
      let str = '';
      for (let i = 0; i < chunks; i++) {
        const chunk = cache.get(key + '_' + i);
        if (chunk === null) {
          str = null; // Missed a chunk
          break;
        }
        str += chunk;
      }
      cachedStr = str;
    }
  }
  
  if (cachedStr) {
    try {
      return JSON.parse(cachedStr); // Return fast cached data
    } catch(e) {
      // JSON parse error, ignore and fetch again
    }
  }
  
  try {
    const data = fetchFunction(); // Execute the slow fetch operation
    if (data) {
      const strValue = JSON.stringify(data);
      const CHUNK_SIZE = 90000;
      
      if (strValue.length <= CHUNK_SIZE) {
        cache.put(key, strValue, expirationSec);
        cache.put(key + '_chunks', '1', expirationSec);
      } else {
        const chunks = Math.ceil(strValue.length / CHUNK_SIZE);
        cache.put(key + '_chunks', chunks.toString(), expirationSec);
        const obj = {};
        for (let i = 0; i < chunks; i++) {
          obj[key + '_' + i] = strValue.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
        }
        cache.putAll(obj, expirationSec);
      }
    }
    return data;
  } catch (e) {
    return null;
  }
}

/**
 * Removes a chunked cache key.
 */
function invalidateCacheKey(key) {
  try {
    const cache = CacheService.getScriptCache();
    const chunksStr = cache.get(key + '_chunks');
    if (chunksStr) {
      const chunks = parseInt(chunksStr, 10);
      let keysToRemove = [key, key + '_chunks'];
      if (chunks > 1) {
        for (let i = 0; i < chunks; i++) keysToRemove.push(key + '_' + i);
      }
      cache.removeAll(keysToRemove);
    }
  } catch(e) {}
}

/**
 * FIX (1.1/1.2): Invalida le chiavi di cache dei dati "stabili" (config).
 * Va chiamata da OGNI funzione che modifica Config_Category o Config_FixedExpenses.
 */
function invalidateConfigCache() {
  invalidateCacheKey('APP_CONFIG_V2');
}

/**
 * Invalida la cache del portafoglio Live (chiamare quando si aggiungono transazioni di investimento)
 */
function invalidatePortfolioCache() {
  invalidateCacheKey('LIVE_PORTFOLIO');
}

/**
 * Invalida la cache della Watchlist (chiamare quando si aggiungono/rimuovono ticket)
 */
function invalidateWatchlistCache() {
  invalidateCacheKey('WATCHLIST');
}