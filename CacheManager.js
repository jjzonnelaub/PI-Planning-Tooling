/**
 * Unified Cache Manager for PI Planning Tool
 * All caching operations go through this module
 * Cache expires after 5 minutes
 */
const CacheManager = {
  CACHE_EXPIRATION_MINUTES: 5, // Changed from 60 to 5 minutes
  CHUNK_SIZE: 90000, // Leave room for overhead in cache
  
  /**
   * Get cached data by key
   * @param {string} key - Cache key
   * @return {any} Cached data or null
   */
  get: function(key) {
    try {
      const cache = CacheService.getScriptCache();
      const countKey = `${key}_count`;
      const chunkCount = cache.get(countKey);
      
      if (!chunkCount) {
        console.log(`Cache miss for key: ${key}`);
        return null;
      }
      
      const chunks = [];
      const count = parseInt(chunkCount);
      
      for (let i = 0; i < count; i++) {
        const chunkKey = `${key}_${i}`;
        const chunk = cache.get(chunkKey);
        if (!chunk) {
          console.log(`Cache chunk missing: ${chunkKey}`);
          return null;
        }
        chunks.push(chunk);
      }
      
      const fullData = chunks.join('');
      console.log(`Cache hit for key: ${key} (${fullData.length} bytes)`);
      return JSON.parse(fullData);
      
    } catch (error) {
      console.error('Cache read error:', error);
      return null;
    }
  },
  
  /**
   * Set cached data
   * @param {string} key - Cache key
   * @param {any} data - Data to cache
   * @return {boolean} Success status
   */
  set: function(key, data) {
    try {
      const cache = CacheService.getScriptCache();
      const serialized = JSON.stringify(data);
      
      // Check size
      if (serialized.length > 1000000) { // 1MB limit warning
        console.warn(`Large cache data for key ${key}: ${serialized.length} bytes`);
      }
      
      // Clear any existing data for this key
      this.clear(key);
      
      // Split into chunks
      const chunks = [];
      for (let i = 0; i < serialized.length; i += this.CHUNK_SIZE) {
        chunks.push(serialized.substring(i, i + this.CHUNK_SIZE));
      }
      
      // Store chunks
      const cacheData = {};
      chunks.forEach((chunk, index) => {
        cacheData[`${key}_${index}`] = chunk;
      });
      cacheData[`${key}_count`] = chunks.length.toString();
      
      // Batch put for efficiency
      cache.putAll(cacheData, this.CACHE_EXPIRATION_MINUTES * 60);
      
      console.log(`Cached data for key: ${key} (${chunks.length} chunks, expires in ${this.CACHE_EXPIRATION_MINUTES} minutes)`);
      return true;
      
    } catch (error) {
      console.error('Cache write error:', error);
      return false;
    }
  },
  
  /**
   * Clear cached data for a specific key
   * @param {string} key - Cache key
   */
  clear: function(key) {
    try {
      const cache = CacheService.getScriptCache();
      const countKey = `${key}_count`;
      const count = cache.get(countKey);
      
      if (count) {
        const keysToRemove = [];
        const chunkCount = parseInt(count);
        
        for (let i = 0; i < chunkCount; i++) {
          keysToRemove.push(`${key}_${i}`);
        }
        keysToRemove.push(countKey);
        
        cache.removeAll(keysToRemove);
        console.log(`Cleared cache for key: ${key}`);
      }
    } catch (error) {
      console.error('Cache clear error:', error);
    }
  },
  
  /**
   * Clear all PI-related caches
   * @param {string} piNumber - PI number
   */
  clearPI: function(piNumber) {
    try {
      const cache = CacheService.getScriptCache();
      const allValueStreams = getAvailableValueStreams();
      const keysToRemove = [];
      
      // Generate all possible cache key patterns for this PI
      keysToRemove.push(`pi_analysis_${piNumber}`);
      
      // Individual value streams
      allValueStreams.forEach(vs => {
        keysToRemove.push(`pi_analysis_${piNumber}_${vs}`);
      });
      
      // All possible combinations
      for (let i = 0; i < 20; i++) {
        keysToRemove.push(`pi_analysis_${piNumber}_${i}`);
        keysToRemove.push(`pi_analysis_${piNumber}_count`);
      }
      
      // Remove all at once
      cache.removeAll(keysToRemove);
      console.log(`Cleared all caches for PI ${piNumber}`);
      
    } catch (error) {
      console.error('Error clearing PI cache:', error);
    }
  },
  
  /**
   * Clear all caches (nuclear option)
   */
  clearAll: function() {
    try {
      const cache = CacheService.getScriptCache();
      // This is limited by quota, but it's the best we can do
      cache.removeAll([]);
      console.log('Cleared all caches');
    } catch (error) {
      console.error('Error clearing all caches:', error);
    }
  },
  
  /**
   * Check if caching is enabled (can be toggled for debugging)
   */
  isEnabled: function() {
    // You can add a script property to disable caching globally
    const props = PropertiesService.getScriptProperties();
    const cacheEnabled = props.getProperty('CACHE_ENABLED');
    return cacheEnabled !== 'false'; // Default to true if not set
  }
};