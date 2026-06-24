/**
 * Generic caching wrapper implementing the "Cache-Aside" pattern.
 * Checks the Script Cache first; if missing, executes the provided fetch function,
 * stores the result, and returns it.
 * * @param {string} key - The unique cache key.
 * @param {Function} fetchFunction - The callback to execute if cache misses (slow operation).
 * @param {number} [expirationSec=600] - Cache duration in seconds (Default: 600s / 10 min).
 * @return {Object|null} The parsed data or null on error.
 */
function getFromCache(key, fetchFunction, expirationSec = 600) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  
  if (cached) {
    return JSON.parse(cached); // Return fast cached data
  }
  
  try {
    const data = fetchFunction(); // Execute the slow fetch operation
    if (data) {
      cache.put(key, JSON.stringify(data), expirationSec);
    }
    return data;
  } catch (e) {
    return null;
  }
}

/**
 * FIX (1.1/1.2): Invalida le chiavi di cache dei dati "stabili" (config).
 * Va chiamata da OGNI funzione che modifica Config_Category o Config_FixedExpenses,
 * altrimenti dopo un salvataggio l'utente vedrebbe dati vecchi finché non scade il TTL.
 * Aggiungere qui eventuali nuove chiavi cache man mano che si attiva getFromCache().
 */
function invalidateConfigCache() {
  try {
    CacheService.getScriptCache().removeAll(['APP_CONFIG_V2']);
  } catch (e) {
    // Best-effort: un fallimento di invalidazione non deve bloccare la scrittura.
  }
}