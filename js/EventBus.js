/**
 * EventBus - Decoupled communication system for Equate services
 * Enables loose coupling between app services without direct dependencies
 */
class EventBus {
    constructor() {
        this.events = new Map();
        this.debugEnabled = false;
    }

    /**
     * Enable/disable debug logging for event tracking
     * @param {boolean} enabled - Whether to log event emissions and subscriptions
     */
    setDebugMode(enabled) {
        this.debugEnabled = enabled;
    }

    /**
     * Subscribe to an event
     * @param {string} eventName - Name of the event to listen for
     * @param {Function} callback - Function to call when event is emitted
     * @param {Object} options - Optional configuration
     * @param {boolean} options.once - If true, callback will only be called once then removed
     * @returns {Function} Unsubscribe function to remove this listener
     */
    on(eventName, callback, options = {}) {
        if (typeof eventName !== 'string') {
            throw new Error('EventBus.on: eventName must be a string');
        }
        if (typeof callback !== 'function') {
            throw new Error('EventBus.on: callback must be a function');
        }

        if (!this.events.has(eventName)) {
            this.events.set(eventName, []);
        }

        const listener = {
            callback,
            once: options.once || false,
            id: Date.now() + Math.random() // Unique identifier
        };

        this.events.get(eventName).push(listener);

        if (this.debugEnabled) {
            console.log(`[EventBus] Subscribed to '${eventName}' (${options.once ? 'once' : 'persistent'})`);
        }

        // Return unsubscribe function
        return () => {
            const listeners = this.events.get(eventName);
            if (listeners) {
                const index = listeners.findIndex(l => l.id === listener.id);
                if (index !== -1) {
                    listeners.splice(index, 1);
                    if (this.debugEnabled) {
                        console.log(`[EventBus] Unsubscribed from '${eventName}'`);
                    }
                }
            }
        };
    }

    /**
     * Subscribe to an event that will only be called once
     * @param {string} eventName - Name of the event to listen for
     * @param {Function} callback - Function to call when event is emitted
     * @returns {Function} Unsubscribe function
     */
    once(eventName, callback) {
        return this.on(eventName, callback, { once: true });
    }

    /**
     * Emit an event to all subscribers
     * @param {string} eventName - Name of the event to emit
     * @param {any} data - Data to pass to event listeners
     * @returns {Promise<any[]>} Promise that resolves with array of all callback return values
     */
    async emit(eventName, data) {
        if (typeof eventName !== 'string') {
            throw new Error('EventBus.emit: eventName must be a string');
        }

        const listeners = this.events.get(eventName);
        if (!listeners || listeners.length === 0) {
            if (this.debugEnabled) {
                console.log(`[EventBus] No listeners for '${eventName}'`);
            }
            return [];
        }

        if (this.debugEnabled) {
            console.log(`[EventBus] Emitting '${eventName}' to ${listeners.length} listener(s)`, data);
        }

        const results = [];
        const toRemove = [];

        // Process all listeners
        for (const listener of listeners) {
            try {
                const result = await listener.callback(data);
                results.push(result);

                // Mark once listeners for removal
                if (listener.once) {
                    toRemove.push(listener);
                }
            } catch (error) {
                console.error(`[EventBus] Error in listener for '${eventName}':`, error);
                // Continue processing other listeners even if one fails
                results.push(null);
            }
        }

        // Remove once listeners
        if (toRemove.length > 0) {
            const remainingListeners = listeners.filter(l => !toRemove.includes(l));
            this.events.set(eventName, remainingListeners);
        }

        return results;
    }

    /**
     * Emit event synchronously (non-blocking)
     * Useful when you don't need to wait for listeners to complete
     * @param {string} eventName - Name of the event to emit
     * @param {any} data - Data to pass to event listeners
     */
    emitSync(eventName, data) {
        // Don't await - fire and forget
        this.emit(eventName, data).catch(error => {
            console.error(`[EventBus] Unhandled error in async emit for '${eventName}':`, error);
        });
    }

    /**
     * Remove all listeners for a specific event
     * @param {string} eventName - Name of the event to clear
     */
    off(eventName) {
        if (this.events.has(eventName)) {
            this.events.delete(eventName);
            if (this.debugEnabled) {
                console.log(`[EventBus] Cleared all listeners for '${eventName}'`);
            }
        }
    }

    /**
     * Remove all event listeners
     */
    clear() {
        this.events.clear();
        if (this.debugEnabled) {
            console.log('[EventBus] Cleared all event listeners');
        }
    }

    /**
     * Get list of all registered event names
     * @returns {string[]} Array of event names
     */
    getEventNames() {
        return Array.from(this.events.keys());
    }

    /**
     * Get number of listeners for a specific event
     * @param {string} eventName - Name of the event
     * @returns {number} Number of listeners
     */
    getListenerCount(eventName) {
        const listeners = this.events.get(eventName);
        return listeners ? listeners.length : 0;
    }

    /**
     * Wait for a specific event to be emitted
     * @param {string} eventName - Name of the event to wait for
     * @param {number} timeout - Optional timeout in milliseconds
     * @returns {Promise} Promise that resolves with event data
     */
    waitFor(eventName, timeout = null) {
        return new Promise((resolve, reject) => {
            let timeoutId = null;

            const unsubscribe = this.once(eventName, (data) => {
                if (timeoutId) clearTimeout(timeoutId);
                resolve(data);
            });

            if (timeout) {
                timeoutId = setTimeout(() => {
                    unsubscribe();
                    reject(new Error(`EventBus.waitFor: Timeout waiting for '${eventName}'`));
                }, timeout);
            }
        });
    }
}

// Common event names used throughout the application
EventBus.Events = {
    // Currency events
    CURRENCY_CHANGED: 'currency-changed',
    CURRENCY_DATA_LOADED: 'currency-data-loaded',
    CURRENCY_SELECTOR_INITIALIZED: 'currency-selector-initialized',
    
    // Calculation events  
    RECALCULATION_NEEDED: 'recalculation-needed',
    CALCULATIONS_UPDATED: 'calculations-updated',
    
    // UI events
    LOADING_STARTED: 'loading-started',
    LOADING_FINISHED: 'loading-finished',
    CHARTS_UPDATED: 'charts-updated',
    
    // Data events
    PORTFOLIO_DATA_LOADED: 'portfolio-data-loaded',
    HISTORICAL_DATA_LOADED: 'historical-data-loaded',
    
    // File processing events
    FILE_UPLOAD_STARTED: 'file-upload-started',
    FILE_ANALYSIS_COMPLETED: 'file-analysis-completed'
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = EventBus;
} else if (typeof window !== 'undefined') {
    window.EventBus = EventBus;
}