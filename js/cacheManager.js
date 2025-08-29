/**
 * CacheManager - Session-based cache storage for debugging and inspection
 * Only active on localhost for development purposes
 */

class CacheManager {
    constructor() {
        // Session-only cache storage (cleared on page refresh)
        this.cache = new Map();
        this.isLocalhost = window.location.href.includes('localhost') || 
                          window.location.href.includes('127.0.0.1') ||
                          window.location.href.includes('file://');
        
        config.debug('🗄️ CacheManager initialized', { isLocalhost: this.isLocalhost });
    }

    /**
     * Store data in cache with timestamp
     * @param {string} key - Cache key
     * @param {any} data - Data to cache
     * @param {string} description - Human-readable description
     */
    set(key, data, description = '') {
        if (!this.isLocalhost) return; // Only cache on localhost
        
        const cacheEntry = {
            data: data,
            timestamp: new Date().toISOString(),
            description: description,
            size: this.getDataSize(data)
        };
        
        this.cache.set(key, cacheEntry);
        config.debug(`📦 Cached: ${key}`, { size: cacheEntry.size, description });
    }

    /**
     * Retrieve data from cache
     * @param {string} key - Cache key
     * @returns {any} Cached data or null
     */
    get(key) {
        if (!this.isLocalhost) return null;
        
        const entry = this.cache.get(key);
        return entry ? entry.data : null;
    }

    /**
     * Get cache entry with metadata
     * @param {string} key - Cache key
     * @returns {object} Cache entry with metadata
     */
    getEntry(key) {
        if (!this.isLocalhost) return null;
        return this.cache.get(key) || null;
    }

    /**
     * Get all cache entries for inspection
     * @returns {object} All cache entries with metadata
     */
    getAllEntries() {
        if (!this.isLocalhost) return {};
        
        const entries = {};
        for (const [key, value] of this.cache.entries()) {
            entries[key] = {
                description: value.description,
                timestamp: value.timestamp,
                size: value.size,
                dataType: Array.isArray(value.data) ? 'Array' : typeof value.data,
                length: Array.isArray(value.data) ? value.data.length : undefined,
                sampleData: this.getSampleData(value.data)
            };
        }
        return entries;
    }

    /**
     * Clear specific cache entry
     * @param {string} key - Cache key to clear
     */
    clear(key) {
        if (!this.isLocalhost) return;
        
        this.cache.delete(key);
        config.debug(`🗑️ Cleared cache: ${key}`);
    }

    /**
     * Clear all cache entries
     */
    clearAll() {
        if (!this.isLocalhost) return;
        
        this.cache.clear();
        config.debug('🗑️ Cleared all cache entries');
    }

    /**
     * Check if cache inspector should be available
     * @returns {boolean} True if localhost
     */
    isAvailable() {
        return this.isLocalhost;
    }

    /**
     * Get approximate data size for display
     * @param {any} data - Data to measure
     * @returns {string} Human-readable size
     */
    getDataSize(data) {
        try {
            const jsonStr = JSON.stringify(data);
            const bytes = new Blob([jsonStr]).size;
            
            if (bytes < 1024) return `${bytes} bytes`;
            if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
            return `${Math.round(bytes / (1024 * 1024))} MB`;
        } catch (e) {
            return 'Unknown size';
        }
    }

    /**
     * Get sample data for preview
     * @param {any} data - Data to sample
     * @returns {any} Sample data for display
     */
    getSampleData(data) {
        if (Array.isArray(data)) {
            if (data.length === 0) return [];
            
            // For timeline and price data, show more comprehensive sample
            if (data.length > 0 && (data[0].date || data[0].price)) {
                // Show first 5, some middle samples, and last 5
                const sample = [];
                
                // First 5
                sample.push(...data.slice(0, Math.min(5, data.length)));
                
                if (data.length > 15) {
                    sample.push('...');
                    
                    // Some middle samples
                    const midStart = Math.floor(data.length * 0.3);
                    const midEnd = Math.floor(data.length * 0.7);
                    sample.push(...data.slice(midStart, Math.min(midStart + 3, midEnd)));
                    
                    sample.push('...');
                    
                    // Last 5
                    sample.push(...data.slice(-5));
                } else if (data.length > 10) {
                    sample.push('...');
                    sample.push(...data.slice(-5));
                }
                
                return sample;
            }
            
            // For other arrays, use original logic
            if (data.length <= 3) return data;
            return [data[0], '...', data[data.length - 1]];
        }
        
        if (typeof data === 'object' && data !== null) {
            const keys = Object.keys(data);
            const jsonSize = JSON.stringify(data).length;
            
            // Show full object if it's small (< 1KB) or has few properties
            if (keys.length <= 10 || jsonSize < 1000) {
                return data;
            }
            
            // Only truncate large objects
            const sample = {};
            keys.slice(0, 5).forEach(key => sample[key] = data[key]);
            sample['...'] = `${keys.length - 5} more properties`;
            return sample;
        }
        
        return data;
    }

    /**
     * Get full data for inspection (use carefully with large datasets)
     * @param {string} key - Cache key
     * @returns {any} Full cached data
     */
    getFullData(key) {
        if (!this.isLocalhost) return null;
        
        const entry = this.cache.get(key);
        return entry ? entry.data : null;
    }

    /**
     * Get cache statistics
     * @returns {object} Cache statistics
     */
    getStats() {
        if (!this.isLocalhost) return {};
        
        return {
            totalEntries: this.cache.size,
            keys: Array.from(this.cache.keys()),
            totalSize: Array.from(this.cache.values())
                .reduce((total, entry) => total + (entry.size ? 1 : 0), 0)
        };
    }
}

// Global cache manager instance
const cacheManager = new CacheManager();

// Export for use in other modules
window.CacheManager = cacheManager;