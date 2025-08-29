/**
 * Production Configuration for Equate Portfolio Analysis
 * Centralizes environment settings and feature flags
 */

class EquateConfig {
    constructor() {
        // Production vs Development mode
        this.isProduction = !window.location.href.includes('localhost') && 
                           !window.location.href.includes('127.0.0.1') &&
                           !window.location.href.includes('file://');
        
        // Debug logging (disabled in production)
        this.enableDebugLogging = !this.isProduction;
        
        // Version information
        this.version = '1.0.0';
        this.buildDate = new Date().toISOString().split('T')[0];
        
        // Feature flags
        this.features = {
            enableAdvancedCharts: true,
            enableExport: true,
            enableThemes: true
        };
        
        // Performance settings
        this.performance = {
            maxFileSize: 10 * 1024 * 1024, // 10MB
            maxHistoricalEntries: 10000
        };
    }
    
    /**
     * Debug logging wrapper - only logs in development
     * @param {string} message - Message to log
     * @param {...any} args - Additional arguments
     */
    debug(message, ...args) {
        if (this.enableDebugLogging) {
            console.log(message, ...args);
        }
    }
    
    /**
     * Info logging - reduced verbosity in production
     * @param {string} message - Message to log
     * @param {...any} args - Additional arguments
     */
    info(message, ...args) {
        if (this.enableDebugLogging) {
            console.log(message, ...args);
        }
    }
    
    /**
     * Warning logging - always enabled
     * @param {string} message - Warning message
     * @param {...any} args - Additional arguments
     */
    warn(message, ...args) {
        console.warn(message, ...args);
    }
    
    /**
     * Error logging - always enabled
     * @param {string} message - Error message
     * @param {...any} args - Additional arguments
     */
    error(message, ...args) {
        console.error(message, ...args);
    }
    
    /**
     * Success logging - only in development
     * @param {string} message - Success message
     * @param {...any} args - Additional arguments
     */
    success(message, ...args) {
        if (this.enableDebugLogging) {
            console.log(message, ...args);
        }
    }
}

// Global configuration instance
const config = new EquateConfig();

// Export for use in other modules
window.EquateConfig = config;