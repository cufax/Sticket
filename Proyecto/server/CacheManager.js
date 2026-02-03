const CacheManager = {
  TTL: 600,
  CHUNK_SIZE: 45000,

  getOrSet: function (key, fetcher) {
    const cache = CacheService.getScriptCache();
    const cachedData = this._getLarge(cache, key);
    if (cachedData) return JSON.parse(cachedData);

    const data = fetcher();
    if (data) this._putLarge(cache, key, JSON.stringify(data));
    return data;
  },

  invalidate: function (key) {
    const cache = CacheService.getScriptCache();
    cache.remove(key);
    const meta = cache.get(key + "_meta");
    if (meta) {
      const chunks = parseInt(meta);
      for (let i = 0; i < chunks; i++) cache.remove(key + "_" + i);
      cache.remove(key + "_meta");
    }
  },

  _putLarge: function (cache, key, value) {
    const chunks = [];
    let index = 0;
    while (index < value.length) {
      chunks.push(value.substr(index, this.CHUNK_SIZE));
      index += this.CHUNK_SIZE;
    }
    cache.put(key + "_meta", chunks.length.toString(), this.TTL);
    chunks.forEach((chunk, i) => cache.put(key + "_" + i, chunk, this.TTL));
    if (value.length < this.CHUNK_SIZE) cache.put(key, value, this.TTL);
  },

  _getLarge: function (cache, key) {
    const simple = cache.get(key);
    if (simple && !cache.get(key + "_meta")) return simple;
    const meta = cache.get(key + "_meta");
    if (!meta) return null;
    const numChunks = parseInt(meta);
    let fullString = "";
    for (let i = 0; i < numChunks; i++) {
      const chunk = cache.get(key + "_" + i);
      if (!chunk) return null;
      fullString += chunk;
    }
    return fullString;
  }
};
