// ==========================================
// ============ Cache helpers ===============
// ==========================================

// CacheService rejects values over 100 KB per key. Staying under that with a
// margin is cheaper than letting put() throw and break the calling feature.
const CACHE_MAX_VALUE_BYTES = 90000;

const CACHE_TTL_CALENDAR = 180;
const CACHE_TTL_AI_CONTEXT = 60;
const CACHE_TTL_CLICKUP_TASKS = 120;

// Every cache key the AI context is assembled from, so invalidation can be blanket
const AI_CONTEXT_CACHE_KEYS = ['ai_ctx_full', 'ai_ctx_tasks'];

/**
 * Reads a JSON value from the user cache. Returns null on a miss or on a
 * corrupted entry — callers always have a way to recompute.
 */
function cacheGetJson(key) {
  try {
    const raw = CacheService.getUserCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    Logger.log('Кеш: не вдалося прочитати ' + key + ': ' + e.message);
    return null;
  }
}

/**
 * Writes a JSON value to the user cache, silently skipping values that exceed
 * the per-key limit. Caching is an optimization — never let it break a feature.
 */
function cachePutJson(key, value, ttlSeconds) {
  try {
    const serialized = JSON.stringify(value);
    if (serialized.length > CACHE_MAX_VALUE_BYTES) {
      Logger.log('Кеш: ' + key + ' завеликий (' + serialized.length + ' байт), не кешую');
      return;
    }
    CacheService.getUserCache().put(key, serialized, ttlSeconds);
  } catch (e) {
    Logger.log('Кеш: не вдалося записати ' + key + ': ' + e.message);
  }
}

/**
 * Drops the cached AI context. Must run after anything that changes what the
 * assistant would see, otherwise it keeps answering from stale data for a minute.
 */
function invalidateAiContextCache() {
  try {
    CacheService.getUserCache().removeAll(AI_CONTEXT_CACHE_KEYS);
  } catch (e) {
    Logger.log('Кеш: не вдалося скинути контекст ШІ: ' + e.message);
  }
}

/**
 * Drops cached calendar events for every look-ahead window the app uses.
 */
function invalidateCalendarCache() {
  try {
    CacheService.getUserCache().removeAll(['cal_events_1', 'cal_events_7']);
  } catch (e) {
    Logger.log('Кеш: не вдалося скинути календар: ' + e.message);
  }
}
