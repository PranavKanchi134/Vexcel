// Lightweight selector cache for stable Google Sheets controls.

const VexcelSelectorCache = (() => {
  const cache = new Map();

  function isUsable(value) {
    return !!(value && value.isConnected);
  }

  function get(key, resolver, validator = isUsable) {
    const cached = cache.get(key);
    if (cached && validator(cached)) return cached;
    const value = resolver();
    if (value && validator(value)) {
      cache.set(key, value);
      return value;
    }
    cache.delete(key);
    return null;
  }

  function set(key, value) {
    if (value) cache.set(key, value);
  }

  function invalidate(key) {
    cache.delete(key);
  }

  function invalidatePrefix(prefix) {
    for (const key of cache.keys()) {
      if (key.startsWith(prefix)) cache.delete(key);
    }
  }

  return { get, set, invalidate, invalidatePrefix, isUsable };
})();
