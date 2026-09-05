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
 * Measures a string the way CacheService does — in UTF-8 bytes.
 *
 * String.length counts UTF-16 units, so every Cyrillic character was counted as
 * one byte instead of two. Values up to twice the real limit sailed past the
 * guard, put() threw the actual 100 KB error, and the throw was swallowed — so
 * past roughly a hundred tasks the cache silently stopped working altogether
 * and every assistant turn re-read all the sheets, calendars and ClickUp.
 * @param {string} text
 * @returns {number} Size in bytes
 */
function byteLength(text) {
  return Utilities.newBlob(text).getBytes().length;
}

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
    const bytes = byteLength(serialized);
    if (bytes > CACHE_MAX_VALUE_BYTES) {
      Logger.log('Кеш: ' + key + ' завеликий (' + bytes + ' байт), не кешую');
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
 *
 * The key is built from the caller's own daysAhead, so a hardcoded pair of
 * windows drifted out of date the moment a new one appeared: the assistant
 * caches a 2-day window that nothing ever cleared, and went on describing a
 * meeting for another three minutes after it was deleted. Clearing the whole
 * plausible range costs one call and cannot fall behind again.
 */
function invalidateCalendarCache() {
  try {
    const keys = [];
    for (var days = 1; days <= 31; days++) {
      keys.push('cal_events_' + days);
    }
    CacheService.getUserCache().removeAll(keys);
  } catch (e) {
    Logger.log('Кеш: не вдалося скинути календар: ' + e.message);
  }
}
