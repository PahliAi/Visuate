/**
 * Currency Code ↔ Symbol Mappings
 * Comprehensive mapping of currency codes to symbols and vice versa
 * Supports all major currencies that EquatePlus might use
 */

class CurrencyMappings {
    constructor() {
        // Comprehensive currency mappings (code → symbol, name)
        this.currencies = {
            // Major currencies
            'USD': { symbol: '$', name: 'US Dollar', regions: ['US'] },
            'EUR': { symbol: '€', name: 'Euro', regions: ['EU'] },
            'GBP': { symbol: '£', name: 'British Pound', regions: ['GB'] },
            'JPY': { symbol: '¥', name: 'Japanese Yen', regions: ['JP'] },
            'CNY': { symbol: '¥', name: 'Chinese Yuan', regions: ['CN'] },
            'CHF': { symbol: 'CHF', name: 'Swiss Franc', regions: ['CH'] },
            'CAD': { symbol: 'C$', name: 'Canadian Dollar', regions: ['CA'] },
            'AUD': { symbol: 'A$', name: 'Australian Dollar', regions: ['AU'] },
            
            // European currencies
            'SEK': { symbol: 'kr', name: 'Swedish Krona', regions: ['SE'] },
            'NOK': { symbol: 'kr', name: 'Norwegian Krone', regions: ['NO'] },
            'DKK': { symbol: 'kr', name: 'Danish Krone', regions: ['DK'] },
            'PLN': { symbol: 'zł', name: 'Polish Zloty', regions: ['PL'] },
            'CZK': { symbol: 'Kč', name: 'Czech Koruna', regions: ['CZ'] },
            'HUF': { symbol: 'Ft', name: 'Hungarian Forint', regions: ['HU'] },
            'RON': { symbol: 'lei', name: 'Romanian Leu', regions: ['RO'] },
            'BGN': { symbol: 'лв', name: 'Bulgarian Lev', regions: ['BG'] },
            'HRK': { symbol: 'kn', name: 'Croatian Kuna', regions: ['HR'] },
            
            // Other major currencies
            'RUB': { symbol: '₽', name: 'Russian Ruble', regions: ['RU'] },
            'TRY': { symbol: '₺', name: 'Turkish Lira', regions: ['TR'] },
            'INR': { symbol: '₹', name: 'Indian Rupee', regions: ['IN'] },
            'KRW': { symbol: '₩', name: 'South Korean Won', regions: ['KR'] },
            'SGD': { symbol: 'S$', name: 'Singapore Dollar', regions: ['SG'] },
            'HKD': { symbol: 'HK$', name: 'Hong Kong Dollar', regions: ['HK'] },
            'NZD': { symbol: 'NZ$', name: 'New Zealand Dollar', regions: ['NZ'] },
            'ZAR': { symbol: 'R', name: 'South African Rand', regions: ['ZA'] },
            'BRL': { symbol: 'R$', name: 'Brazilian Real', regions: ['BR'] },
            'MXN': { symbol: '$', name: 'Mexican Peso', regions: ['MX'] },
            
            // Middle East & Africa
            'ILS': { symbol: '₪', name: 'Israeli Shekel', regions: ['IL'] },
            'AED': { symbol: 'د.إ', name: 'UAE Dirham', regions: ['AE'] },
            'SAR': { symbol: '﷼', name: 'Saudi Riyal', regions: ['SA'] },
            'EGP': { symbol: 'E£', name: 'Egyptian Pound', regions: ['EG'] },
            'NGN': { symbol: '₦', name: 'Nigerian Naira', regions: ['NG'] },
            
            // Asia Pacific
            'THB': { symbol: '฿', name: 'Thai Baht', regions: ['TH'] },
            'VND': { symbol: '₫', name: 'Vietnamese Dong', regions: ['VN'] },
            'MYR': { symbol: 'RM', name: 'Malaysian Ringgit', regions: ['MY'] },
            'IDR': { symbol: 'Rp', name: 'Indonesian Rupiah', regions: ['ID'] },
            'PHP': { symbol: '₱', name: 'Philippine Peso', regions: ['PH'] },
            'TWD': { symbol: 'NT$', name: 'Taiwan Dollar', regions: ['TW'] },
            
            // Americas
            'CLP': { symbol: '$', name: 'Chilean Peso', regions: ['CL'] },
            'ARS': { symbol: '$', name: 'Argentine Peso', regions: ['AR'] },
            'COP': { symbol: '$', name: 'Colombian Peso', regions: ['CO'] },
            'PEN': { symbol: 'S/.', name: 'Peruvian Sol', regions: ['PE'] },
            
            // Other
            'ISK': { symbol: 'kr', name: 'Icelandic Krona', regions: ['IS'] },
            'UAH': { symbol: '₴', name: 'Ukrainian Hryvnia', regions: ['UA'] },
            'KZT': { symbol: '₸', name: 'Kazakhstani Tenge', regions: ['KZ'] }
        };
        
        // Create reverse mappings (symbol → possible codes)
        this.symbolToCodes = {};
        Object.keys(this.currencies).forEach(code => {
            const symbol = this.currencies[code].symbol;
            if (!this.symbolToCodes[symbol]) {
                this.symbolToCodes[symbol] = [];
            }
            this.symbolToCodes[symbol].push(code);
        });
        
        // Special compound patterns (US$, CN¥, etc.)
        this.compoundPatterns = {
            'US$': 'USD',
            'CN¥': 'CNY', 
            'JP¥': 'JPY',
            'HK$': 'HKD',
            'NZ$': 'NZD',
            'A$': 'AUD',
            'C$': 'CAD',
            'S$': 'SGD',
            'NT$': 'TWD',
            'R$': 'BRL'
        };
    }
    
    /**
     * Get currency info by code
     */
    getByCode(code) {
        return this.currencies[code.toUpperCase()] || null;
    }
    
    /**
     * Get currency info by symbol
     */
    getBySymbol(symbol) {
        const codes = this.symbolToCodes[symbol];
        if (!codes || codes.length === 0) return null;
        if (codes.length === 1) {
            return { code: codes[0], ...this.currencies[codes[0]] };
        }
        // Multiple currencies use same symbol - return all possibilities
        return codes.map(code => ({ code, ...this.currencies[code] }));
    }
    
    /**
     * Resolve compound pattern (US$, CN¥, etc.)
     */
    resolveCompoundPattern(pattern) {
        const code = this.compoundPatterns[pattern];
        return code ? { code, ...this.currencies[code] } : null;
    }
    
    /**
     * Get all currencies sorted by common usage
     */
    getAllCurrencies() {
        const commonOrder = [
            'USD', 'EUR', 'GBP', 'JPY', 'CNY', 'CHF', 'CAD', 'AUD',
            'SEK', 'NOK', 'DKK', 'PLN', 'CZK', 'HUF', 'RON', 'BGN', 'HRK',
            'RUB', 'TRY', 'INR', 'KRW', 'SGD', 'HKD', 'NZD', 'ZAR', 'BRL', 'MXN'
        ];
        
        const result = [];
        
        // Add common currencies first
        commonOrder.forEach(code => {
            if (this.currencies[code]) {
                result.push({ code, ...this.currencies[code] });
            }
        });
        
        // Add remaining currencies
        Object.keys(this.currencies).forEach(code => {
            if (!commonOrder.includes(code)) {
                result.push({ code, ...this.currencies[code] });
            }
        });
        
        return result;
    }
    
    /**
     * Convert currency codes to display format (symbol + code)
     */
    toDisplayFormat(code) {
        const currency = this.getByCode(code);
        if (!currency) return code;
        
        // For currencies with unique symbols, show symbol first
        if (currency.symbol && currency.symbol !== code) {
            return `${currency.symbol} (${code})`;
        }
        
        // For currencies without unique symbols, show code and name
        return `${code} (${currency.name})`;
    }
    
    /**
     * Get currency symbol for display
     */
    getSymbol(code) {
        const currency = this.getByCode(code);
        return currency ? currency.symbol : code;
    }
    
    /**
     * Create user-friendly currency options for selection
     */
    createCurrencyOptions(codes = null) {
        const currenciesToShow = codes ? 
            codes.map(code => ({ code, ...this.currencies[code] })).filter(c => c.symbol) :
            this.getAllCurrencies();
            
        return currenciesToShow.map(currency => ({
            code: currency.code,
            symbol: currency.symbol,
            name: currency.name,
            display: this.toDisplayFormat(currency.code),
            regions: currency.regions
        }));
    }
}

// Export for both browser and Node.js
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CurrencyMappings;
} else if (typeof window !== 'undefined') {
    window.CurrencyMappings = CurrencyMappings;
}