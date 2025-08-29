/**
 * V2 Currency Formatting Utilities
 * Provides currency symbols and formatting for display purposes only
 * NO CONVERSION LOGIC - V2 uses instant multi-currency switching via enhanced reference points
 */

class CurrencyConverter {
    constructor() {
        // V2: No conversion data needed, only formatting utilities
        config.debug('🎯 V2 CurrencyConverter: Formatting utilities only (no conversion logic)');
    }

    /**
     * Get currency symbol for display
     */
    getCurrencySymbol(currency) {
        if (!currency) return '';
        
        if (typeof CurrencyMappings !== 'undefined') {
            const mappings = new CurrencyMappings();
            const currencyInfo = mappings.getByCode(currency);
            if (currencyInfo && currencyInfo.symbol) {
                return currencyInfo.symbol;
            } else {
                return currency; // Use currency code as fallback
            }
        } else {
            console.warn('⚠️ CurrencyMappings not available');
            return currency;
        }
    }

    /**
     * Format currency value with symbol
     */
    formatCurrency(value, currency = null, decimals = 2) {
        const symbol = this.getCurrencySymbol(currency);
        
        // Get user's preferred number format from app instance
        const format = (typeof app !== 'undefined' && app.numberFormat) ? app.numberFormat : 'us';
        const locale = format === 'eu' ? 'de-DE' : 'en-US';
        
        return `${symbol} ${value.toLocaleString(locale, { 
            minimumFractionDigits: decimals, 
            maximumFractionDigits: decimals 
        })}`;
    }

    /**
     * Format percentage value
     */
    formatPercentage(value, decimals = 2) {
        // Get user's preferred number format from app instance
        const format = (typeof app !== 'undefined' && app.numberFormat) ? app.numberFormat : 'us';
        const locale = format === 'eu' ? 'de-DE' : 'en-US';
        
        return `${value.toLocaleString(locale, { 
            minimumFractionDigits: decimals, 
            maximumFractionDigits: decimals 
        })}%`;
    }

    /**
     * V2: Minimal clear method for compatibility
     */
    clear() {
        config.debug('🧹 V2 CurrencyConverter cleared (no conversion data to clear)');
    }
}

// Global instance for backward compatibility
const currencyConverter = new CurrencyConverter();