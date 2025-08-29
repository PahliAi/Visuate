/**
 * DOMManager - Centralized DOM element management and manipulation
 * Handles caching, retrieval, and updates of DOM elements for Equate services
 */
class DOMManager {
    constructor() {
        this.elements = {};
        this.cachedAt = null;
        this.cacheElements();
    }

    /**
     * Cache all DOM elements used by the application
     * Called on initialization and can be called again if DOM structure changes
     */
    cacheElements() {
        try {
            this.elements = {
                // Theme and Format
                themeSelect: document.getElementById('theme'),
                formatSelect: document.getElementById('formatSelect'),
                languageSelect: document.getElementById('languageSelect'),
                
                // Currency Selector
                currencySelector: document.getElementById('currencySelector'),
                currencyLabelContainer: document.getElementById('currencyLabelContainer'),
                currencySelect: document.getElementById('currencySelect'),
                
                // Price controls (moved to Results tab)
                manualPrice: document.getElementById('manualPriceInput'),
                manualPriceInput: document.getElementById('manualPriceInput'), // New location in Results tab
                customizeExportBtn: document.getElementById('customizeExportBtn'),
                
                // Upload
                portfolioInput: document.getElementById('portfolio'),
                transactionsInput: document.getElementById('transactions'),
                portfolioZone: document.getElementById('portfolioZone'),
                transactionsZone: document.getElementById('transactionsZone'),
                portfolioInfo: document.getElementById('portfolioInfo'),
                transactionsInfo: document.getElementById('transactionsInfo'),
                analyzeBtn: document.getElementById('analyzeBtn'),
                
                // Results
                priceSource: document.getElementById('priceSource'),
                
                // Metrics (all currency-formatted)
                userInvestment: document.getElementById('userInvestment'),
                companyMatch: document.getElementById('companyMatch'),
                freeShares: document.getElementById('freeShares'),
                dividendIncome: document.getElementById('dividendIncome'),
                totalInvestment: document.getElementById('totalInvestment'),
                totalSold: document.getElementById('totalSold'),
                currentValue: document.getElementById('currentValue'),

                currentPortfolioDisclaimer: document.getElementById('currentPortfolioDisclaimer'),
                totalValue: document.getElementById('totalValue'),
                totalReturn: document.getElementById('totalReturn'),
                returnOnTotalInvestment: document.getElementById('returnOnTotalInvestment'),
                returnPercentage: document.getElementById('returnPercentage'),
                returnPercentageOnTotalInvestment: document.getElementById('returnPercentageOnTotalInvestment'),
                availableShares: document.getElementById('availableShares'),
                xirrUserInvestment: document.getElementById('xirrUserInvestment'),
                xirrTotalInvestment: document.getElementById('xirrTotalInvestment'),
                
                // Chart
                chartPlaceholder: document.getElementById('chartPlaceholder'),
                
                // Loading
                loadingOverlay: document.getElementById('loadingOverlay')
            };

            this.cachedAt = new Date();

            // Check for critical missing elements
            const criticalElements = ['portfolioZone', 'transactionsZone', 'analyzeBtn'];
            const missingElements = criticalElements.filter(id => !this.elements[id]);
            
            if (missingElements.length > 0) {
                console.warn('DOMManager: Critical elements missing:', missingElements);
            }

        } catch (error) {
            console.error('DOMManager: Failed to cache DOM elements:', error);
            this.elements = {}; // Ensure elements is always an object
        }
    }

    /**
     * Get a cached DOM element by ID
     * @param {string} elementId - The element ID to retrieve
     * @returns {HTMLElement|null} The DOM element or null if not found
     */
    get(elementId) {
        return this.elements[elementId] || null;
    }

    /**
     * Check if an element exists in the cache
     * @param {string} elementId - The element ID to check
     * @returns {boolean} True if element exists and is still in DOM
     */
    exists(elementId) {
        const element = this.elements[elementId];
        return element && document.contains(element);
    }

    /**
     * Refresh cache for a specific element or all elements
     * @param {string} elementId - Optional specific element to refresh
     */
    refresh(elementId = null) {
        if (elementId) {
            // Refresh specific element
            const domElement = document.getElementById(elementId);
            if (domElement) {
                this.elements[elementId] = domElement;
            } else {
                console.warn(`DOMManager: Element '${elementId}' not found during refresh`);
                this.elements[elementId] = null;
            }
        } else {
            // Refresh all elements
            this.cacheElements();
        }
    }

    // =============================================
    // CURRENCY-SPECIFIC DOM METHODS
    // =============================================

    /**
     * Update all currency-formatted metric displays
     * @param {Object} calculations - Current calculations object
     * @param {string} currency - Currency code (e.g., 'EUR', 'USD')
     * @param {Object} converter - Currency converter instance
     * @param {number} decimals - Number of decimal places to display
     */
    updateCurrencyMetrics(calculations, currency, converter, decimals = 2) {
        if (!calculations || !currency || !converter) {
            console.warn('DOMManager: Missing required parameters for currency metrics update');
            return;
        }

        // Helper function to safely update an element
        const updateElement = (elementId, value, className = null) => {
            const element = this.get(elementId);
            if (element) {
                element.textContent = value;
                if (className) {
                    element.className = className;
                }
            } else if (this.exists(elementId)) {
                console.warn(`DOMManager: Element '${elementId}' found in cache but textContent update failed`);
            }
        };

        try {
            // Basic investment metrics
            updateElement('userInvestment', converter.formatCurrency(calculations.userInvestment, currency, decimals));
            updateElement('companyMatch', converter.formatCurrency(calculations.companyMatch, currency, decimals));
            updateElement('freeShares', converter.formatCurrency(calculations.freeShares, currency, decimals));
            updateElement('dividendIncome', converter.formatCurrency(calculations.dividendIncome, currency, decimals));
            updateElement('totalInvestment', converter.formatCurrency(calculations.totalInvestment, currency, decimals));
            
            // Current value metrics
            updateElement('currentValue', converter.formatCurrency(calculations.currentValue, currency, decimals));
            updateElement('totalSold', converter.formatCurrency(calculations.totalSold, currency, decimals));
            updateElement('totalValue', converter.formatCurrency(calculations.totalValue, currency, decimals));
            
            // Return metrics with styling
            const totalReturn = calculations.totalValue - calculations.userInvestment;
            const returnOnTotalInvestment = calculations.totalValue - calculations.totalInvestment;
            const returnPercentage = calculations.userInvestment > 0 ? (totalReturn / calculations.userInvestment) * 100 : 0;
            const returnPercentageOnTotalInvestment = calculations.totalInvestment > 0 ? (returnOnTotalInvestment / calculations.totalInvestment) * 100 : 0;
            
            updateElement('totalReturn', 
                converter.formatCurrency(totalReturn, currency, decimals),
                `metric-value ${totalReturn >= 0 ? 'positive' : 'negative'}`
            );
            
            updateElement('returnOnTotalInvestment',
                converter.formatCurrency(returnOnTotalInvestment, currency, decimals),
                `metric-value ${returnOnTotalInvestment >= 0 ? 'positive' : 'negative'}`
            );
            
            // Percentage returns
            updateElement('returnPercentage',
                converter.formatPercentage(returnPercentage, decimals),
                `metric-value ${returnPercentage >= 0 ? 'positive' : 'negative'}`
            );
            
            updateElement('returnPercentageOnTotalInvestment',
                converter.formatPercentage(returnPercentageOnTotalInvestment, decimals),
                `metric-value ${returnPercentageOnTotalInvestment >= 0 ? 'positive' : 'negative'}`
            );
            
            // XIRR metrics
            updateElement('xirrUserInvestment',
                converter.formatPercentage(calculations.xirrUserInvestment, decimals),
                `metric-value ${calculations.xirrUserInvestment >= 0 ? 'positive' : 'negative'}`
            );
            
            updateElement('xirrTotalInvestment',
                converter.formatPercentage(calculations.xirrTotalInvestment, decimals),
                `metric-value ${calculations.xirrTotalInvestment >= 0 ? 'positive' : 'negative'}`
            );
            
        } catch (error) {
            console.error('DOMManager: Error updating currency metrics:', error);
        }
    }

    /**
     * Update price source indicator with currency symbol
     * @param {Object} calculations - Current calculations object
     * @param {string} currency - Currency code
     * @param {Object} converter - Currency converter instance
     * @param {Object} translationManager - Translation manager instance
     */
    updatePriceSource(calculations, currency, converter, translationManager) {
        const indicator = this.get('priceSource');
        if (!indicator || !calculations || !currency || !converter) return;

        const currencySymbol = converter.getCurrencySymbol(currency);
        
        try {
            const text = indicator.querySelector('span') || indicator;
            
            if (calculations.priceSource === 'manual') {
                text.innerHTML = `<strong>${translationManager.t('price_label')}:</strong> ${currencySymbol}${app.formatNumber(calculations.currentPrice, 2)} (manual input)`;
            } else {
                text.innerHTML = `<strong>${translationManager.t('price_label')}:</strong> ${currencySymbol}${app.formatNumber(calculations.currentPrice, 2)} ${translationManager.t('as_of_label')} ${calculations.priceDate}`;
            }
        } catch (error) {
            console.error('DOMManager: Error updating price source:', error);
        }
    }

    /**
     * Populate currency selector dropdown
     * @param {Array} currencies - Array of currency objects with code, score, name
     * @param {string} savedCurrency - Currently selected currency code
     * @returns {boolean} Success status
     */
    populateCurrencySelector(currencies, savedCurrency = null) {
        const selectElement = this.get('currencySelect');
        if (!selectElement) {
            console.warn('DOMManager: Currency select element not found');
            return false;
        }

        try {
            // Clear existing options
            selectElement.innerHTML = '';
            
            // Add default option
            const defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = 'Select currency (data quality %)...';
            selectElement.appendChild(defaultOption);
            
            // Add currency options
            currencies.forEach(currency => {
                const option = document.createElement('option');
                option.value = currency.code;
                option.textContent = `${currency.code} (${currency.score}%) - ${currency.name}`;
                
                // Highlight high-quality currencies
                if (currency.score >= 90) {
                    option.style.fontWeight = 'bold';
                    option.style.color = '#2d5a2d';
                }
                
                selectElement.appendChild(option);
            });
            
            // Set saved currency if provided
            if (savedCurrency && selectElement.querySelector(`option[value="${savedCurrency}"]`)) {
                selectElement.value = savedCurrency;
            }
            
            return true;
        } catch (error) {
            console.error('DOMManager: Error populating currency selector:', error);
            selectElement.innerHTML = '<option value="">Error loading currencies</option>';
            return false;
        }
    }

    /**
     * Show currency selector UI elements
     */
    showCurrencySelector() {
        const selector = this.get('currencySelector');
        const labelContainer = this.get('currencyLabelContainer');
        
        if (selector) {
            selector.style.display = 'block';
        }
        if (labelContainer) {
            labelContainer.style.display = 'block';
        }
    }

    /**
     * Hide currency selector UI elements
     */
    hideCurrencySelector() {
        const selector = this.get('currencySelector');
        const labelContainer = this.get('currencyLabelContainer');
        
        if (selector) {
            selector.style.display = 'none';
        }
        if (labelContainer) {
            labelContainer.style.display = 'none';
        }
    }

    /**
     * Set up event listener for currency selection changes
     * @param {Function} callback - Function to call when currency changes
     * @returns {Function|null} Function to remove the event listener
     */
    onCurrencyChange(callback) {
        const selectElement = this.get('currencySelect');
        if (!selectElement) {
            console.warn('DOMManager: Cannot set up currency change listener - element not found');
            return null;
        }

        if (typeof callback !== 'function') {
            console.error('DOMManager: Currency change callback must be a function');
            return null;
        }

        selectElement.addEventListener('change', callback);
        
        // Return cleanup function
        return () => {
            selectElement.removeEventListener('change', callback);
        };
    }

    // =============================================
    // GENERAL DOM UTILITY METHODS
    // =============================================

    /**
     * Show loading overlay with message
     * @param {string} message - Loading message to display
     */
    showLoading(message = 'Loading...') {
        const overlay = this.get('loadingOverlay');
        if (overlay) {
            overlay.style.display = 'flex';
            const messageElement = overlay.querySelector('.loading-message');
            if (messageElement) {
                messageElement.textContent = message;
            }
        }
    }

    /**
     * Hide loading overlay
     */
    hideLoading() {
        const overlay = this.get('loadingOverlay');
        if (overlay) {
            overlay.style.display = 'none';
        }
    }

    /**
     * Get diagnostic information about cached elements
     * @returns {Object} Diagnostic information
     */
    getDiagnostics() {
        const total = Object.keys(this.elements).length;
        const existing = Object.keys(this.elements).filter(id => this.exists(id)).length;
        const missing = Object.keys(this.elements).filter(id => !this.exists(id));
        
        return {
            totalElements: total,
            existingElements: existing,
            missingElements: missing,
            cachedAt: this.cachedAt,
            cacheAge: this.cachedAt ? Date.now() - this.cachedAt.getTime() : null
        };
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DOMManager;
} else if (typeof window !== 'undefined') {
    window.DOMManager = DOMManager;
}