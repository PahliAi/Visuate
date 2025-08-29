/**
 * CurrencyService - Complete currency management for Equate
 * Handles currency detection, data loading, UI management, and conversion operations
 * Extracted from app.js to provide proper separation of concerns
 */
class CurrencyService {
    constructor(dependencies) {
        // Validate required dependencies
        this.validateDependencies(dependencies);
        
        // Inject dependencies (no coupling to app.js "this")
        this.database = dependencies.database;           // equateDB instance
        this.converter = dependencies.converter;         // currencyConverter instance
        this.mappings = dependencies.mappings;          // CurrencyMappings instance
        this.translator = dependencies.translator;       // translationManager instance
        this.calculator = dependencies.calculator;       // portfolioCalculator instance
        this.eventBus = dependencies.eventBus;         // EventBus instance
        this.dom = dependencies.domManager;            // DOMManager instance
        
        // Internal state (moved from app.js)
        this.detectedCurrency = null;                   // Currently detected portfolio currency
        this.cachedCurrencyData = null;                 // In-memory cache of historical currency data
        this.isChangingCurrency = false;                // Flag to prevent concurrent currency changes
        
        // Initialize service
        this.setupEventListeners();
    }

    /**
     * Validate that all required dependencies are provided
     * @param {Object} dependencies - Dependencies object
     * @throws {Error} If required dependencies are missing
     */
    validateDependencies(dependencies) {
        const required = ['database', 'converter', 'mappings', 'translator', 'calculator', 'eventBus', 'domManager'];
        const missing = required.filter(dep => !dependencies || !dependencies[dep]);
        
        if (missing.length > 0) {
            throw new Error(`CurrencyService: Missing required dependencies: ${missing.join(', ')}`);
        }
    }

    /**
     * Set up internal event listeners for service coordination
     */
    setupEventListeners() {
        // Listen for recalculation requests
        this.eventBus.on('recalculation-needed', async () => {
            await this.handleRecalculationRequest();
        });
        
        // Listen for portfolio data changes
        this.eventBus.on('portfolio-data-loaded', (portfolioData) => {
            this.handlePortfolioDataLoaded(portfolioData);
        });
    }

    // =============================================
    // PUBLIC API - STATE MANAGEMENT
    // =============================================

    /**
     * Get currently detected currency
     * @returns {string|null} Currency code or null
     */
    getDetectedCurrency() {
        return this.detectedCurrency;
    }

    /**
     * Set detected currency and notify listeners
     * @param {string} currency - Currency code (e.g., 'EUR', 'USD')
     */
    setDetectedCurrency(currency) {
        if (this.detectedCurrency !== currency) {
            this.detectedCurrency = currency;
            this.eventBus.emitSync(EventBus.Events.CURRENCY_CHANGED, currency);
        }
    }

    /**
     * Check if currency service is currently changing currencies
     * @returns {boolean} True if currency change is in progress
     */
    isChangingCurrencies() {
        return this.isChangingCurrency;
    }

    /**
     * Get cached currency data
     * @returns {Array|null} Cached currency data or null
     */
    getCachedCurrencyData() {
        return this.cachedCurrencyData;
    }

    /**
     * Set cached currency data
     * @param {Array|null} data - Currency data to cache
     */
    setCachedCurrencyData(data) {
        this.cachedCurrencyData = data;
    }

    /**
     * Get the changing currency flag (for preventing concurrent changes)
     * @returns {boolean} True if currently changing currency
     */
    getIsChangingCurrency() {
        return this.isChangingCurrency;
    }

    /**
     * Set the changing currency flag (for preventing concurrent changes)
     * @param {boolean} changing - Whether currency is currently being changed
     */
    setIsChangingCurrency(changing) {
        this.isChangingCurrency = changing;
    }

    // =============================================
    // PUBLIC API - DATA MANAGEMENT
    // =============================================

    /**
     * Load historical currency data from hist.xlsx Currency sheet
     * Manages memory and database caching with portfolio-aware quality scoring
     * @param {Object} portfolioData - Optional portfolio data for quality calculation
     * @returns {Promise<Array>} Array of historical currency entries
     */
    async loadHistoricalData(portfolioData = null, options = {}) {
        config.debug('🔍 loadHistoricalCurrencyData called with portfolio data:', !!portfolioData);
        try {
            // Check in-memory cache first
            if (this.cachedCurrencyData) {
                config.debug(`Historical currency data loaded from memory cache: ${this.cachedCurrencyData.length} entries`);
                
                // ⚡ FIX: If portfolio data is provided, recalculate portfolio-aware quality scores even with memory cache
                if (portfolioData && portfolioData.entries && portfolioData.entries.length > 0) {
                    config.debug('🎯 Recalculating portfolio-aware quality scores for memory cached data...');
                    // For cached data, we need to discover all currency columns (since we don't have original headers)
                    const currencyColumns = this.getAllColumns(this.cachedCurrencyData);
                    config.debug('🔍 Found currency columns from cached data:', currencyColumns.length, 'currencies');
                    const portfolioQualityScores = this.calculatePortfolioAwareScores(this.cachedCurrencyData, currencyColumns, portfolioData);
                    await this.database.saveHistoricalCurrencyData(this.cachedCurrencyData, portfolioQualityScores);
                    config.debug('✅ Portfolio-aware quality scores updated for memory cached data');
                    
                    // IMPORTANT: Initialize currency selector with updated quality scores
                    await this.initializeSelector();
                }
                
                return this.cachedCurrencyData;
            }
            
            // Check if we already have historical currency data in database
            const existingCurrencyData = await this.database.getHistoricalCurrencyData();
            if (existingCurrencyData.length > 0) {
                config.debug(`Historical currency data loaded from database cache: ${existingCurrencyData.length} entries`);
                
                // Store in memory cache for future requests
                this.cachedCurrencyData = existingCurrencyData;
                
                // ⚡ FIX: If portfolio data is provided, recalculate portfolio-aware quality scores
                config.debug('🔍 DEBUG PORTFOLIO DATA CHECK:', {
                    hasPortfolioData: !!portfolioData,
                    hasEntries: portfolioData?.entries ? true : false,
                    entriesLength: portfolioData?.entries?.length || 0,
                    portfolioDataKeys: portfolioData ? Object.keys(portfolioData) : []
                });
                
                if (portfolioData && portfolioData.entries && portfolioData.entries.length > 0) {
                    config.debug('🎯 Recalculating portfolio-aware quality scores for database cached data...');
                    // For cached data, we need to discover all currency columns (since we don't have original headers)
                    const currencyColumns = this.getAllColumns(existingCurrencyData);
                    config.debug('🔍 Found currency columns from cached data:', currencyColumns.length, 'currencies');
                    const portfolioQualityScores = this.calculatePortfolioAwareScores(existingCurrencyData, currencyColumns, portfolioData);
                    await this.database.saveHistoricalCurrencyData(existingCurrencyData, portfolioQualityScores);
                    config.debug('✅ Portfolio-aware quality scores updated for database cached data');
                } else {
                    config.debug('❌ Portfolio-aware quality scores NOT calculated due to missing data condition');
                }
                
                // IMPORTANT: Initialize currency selector even with cached data
                await this.initializeSelector();
                
                return existingCurrencyData;
            }

            config.debug('💱 Loading historical currency data from hist.xlsx Currency sheet...');

            // Environment-aware URL construction (same pattern as loadHistoricalPrices)
            let baseUrl = '';
            const currentUrl = window.location.href;
            
            if (currentUrl.includes('github.io')) {
                // Production: GitHub Pages - use raw GitHub URL
                const parts = currentUrl.split('.github.io')[0].split('//')[1];
                const repoName = currentUrl.split('/')[3] || 'Equate';
                baseUrl = `https://raw.githubusercontent.com/${parts}/${repoName}/main/`;
            } else {
                // Development: Local - use relative path
                baseUrl = './';
            }

            // Try to load currency data with cache buster
            const cacheBuster = `?v=${Date.now()}`;
            const fullUrl = baseUrl + 'hist_base.xlsx' + cacheBuster;
            config.debug('Attempting to fetch currency data from:', fullUrl);
            
            const response = await fetch(fullUrl);
            
            if (!response.ok) {
                config.warn(`Historical currency file fetch failed: ${response.status} ${response.statusText} - will use Synthetic fallback`);
                return [];
            }

            const arrayBuffer = await response.arrayBuffer();
            const data = XLSX.read(arrayBuffer, { type: 'array' });
            
            // Use the 'Currency' sheet
            const worksheet = data.Sheets['Currency'];
            if (!worksheet) {
                return [];
            }
            
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (rawData.length < 2) {
                config.warn('Currency sheet has no data rows - will use Synthetic fallback');
                return [];
            }

            // Get currency column headers (skip 'Date' in column 0)
            const headers = rawData[0];
            const currencyColumns = headers.slice(1); // Skip 'Date' column

            // Parse currency data - convert DD-MM-YYYY to YYYY-MM-DD and extract rates
            let currencyEntries = [];
            for (let i = 1; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row[0]) continue; // Skip rows without dates

                // Handle DD-MM-YYYY string dates (same pattern as stock prices)
                const dateStr = row[0];
                let isoDate;
                
                if (typeof dateStr === 'string' && dateStr.includes('-')) {
                    const parts = dateStr.split('-');
                    if (parts.length === 3) {
                        const day = parts[0].padStart(2, '0');
                        const month = parts[1].padStart(2, '0');
                        const year = parts[2];
                        isoDate = `${year}-${month}-${day}`;
                    }
                } else if (typeof dateStr === 'number') {
                    // Fallback: handle Excel serial dates if any remain
                    const jsDate = this.excelDateToJSDate(dateStr);
                    isoDate = jsDate.toISOString().split('T')[0];
                }
                
                if (!isoDate) {
                    config.warn(`Skipping invalid currency date format in row ${i}:`, dateStr);
                    continue;
                }

                // Extract currency rates for this date
                const dateRates = { date: isoDate };
                for (let j = 1; j < row.length && j <= currencyColumns.length; j++) {
                    const currencyCode = currencyColumns[j - 1];
                    const rate = parseFloat(row[j]);
                    
                    if (currencyCode && !isNaN(rate) && rate > 0) {
                        dateRates[currencyCode] = rate;
                    }
                }
                
                // Only add entry if it has at least one valid currency rate
                if (Object.keys(dateRates).length > 1) { // More than just 'date'
                    currencyEntries.push(dateRates);
                }
            }

            if (currencyEntries.length === 0) {
                config.warn('No valid currency entries found in hist.xlsx Currency sheet');
                return [];
            }

            // Sort by date
            currencyEntries.sort((a, b) => new Date(a.date) - new Date(b.date));
            
            // OPTIMIZATION: Filter by portfolio startDate to reduce data loading from ~200K to ~60K entries
            const originalLength = currencyEntries.length;
            if (portfolioData && portfolioData.startDate) {
                currencyEntries = currencyEntries.filter(entry => entry.date >= portfolioData.startDate);
                config.debug(`🚀 OPTIMIZATION: Filtered currency data from ${originalLength} to ${currencyEntries.length} entries (from ${portfolioData.startDate} onwards)`);
            } else {
                config.debug(`⚠️ No startDate in portfolioData - loading all ${currencyEntries.length} entries (no optimization)`);
            }
            
            config.debug(`About to save ${currencyEntries.length} currency entries to database`);
            config.debug('Sample currency entries:', currencyEntries.slice(0, 2));
            config.debug(`Available currencies:`, currencyColumns.slice(0, 20));
            config.debug(`Total currencies found:`, currencyColumns.length);
            
            // Calculate portfolio-aware quality scores ONLY if portfolio data is provided
            if (portfolioData && portfolioData.entries && portfolioData.entries.length > 0) {
                const portfolioQualityScores = this.calculatePortfolioAwareScores(currencyEntries, currencyColumns, portfolioData);
                await this.database.saveHistoricalCurrencyData(currencyEntries, portfolioQualityScores);
            } else {
                // No portfolio data - save currency data without overwriting existing quality scores
                try {
                    await this.database.saveHistoricalCurrencyData(currencyEntries, null);
                } catch (error) {
                    config.error('ERROR saving currency data without quality scores:', error);
                    // V2: Simple save without quality scores (good data quality assumed)
                    await this.database.saveHistoricalCurrencyData(currencyEntries, null);
                }
            }
            
            // Verify they were saved
            const savedCurrencyData = await this.database.getHistoricalCurrencyData();
            config.debug(`✅ Historical currency data loaded and cached: ${currencyEntries.length} entries`);
            config.debug(`✅ Verified in database: ${savedCurrencyData.length} entries`);
            config.debug(`Date range: ${currencyEntries[0]?.date} to ${currencyEntries[currencyEntries.length - 1]?.date}`);
            
            // Store freshly loaded data in memory cache
            this.cachedCurrencyData = currencyEntries;
            
            // Initialize currency selector with portfolio-aware quality scores
            await this.initializeSelector();
            
            return currencyEntries;
            
        } catch (error) {
            config.warn('Could not load historical currency data:', error.message);
            config.debug('Will use Synthetic fallback for currency conversion');
            return [];
        }
    }

    /**
     * Helper method for Excel date conversion
     * @param {number} excelDate - Excel serial date number
     * @returns {Date} JavaScript Date object
     */
    excelDateToJSDate(excelDate) {
        return new Date((excelDate - 25569) * 86400 * 1000);
    }

    /**
     * Get historical currency data (unified access point for all components)
     * Public API method used by calculator.js and other external modules
     * @returns {Promise<Array>} Historical currency data entries
     */
    async getData() {
        return await this.loadHistoricalData();
    }

    /**
     * Clear currency data cache
     * Called when data is updated or cleared
     */
    clearCache() {
        this.cachedCurrencyData = null;
        config.debug('🗑️ Currency data cache cleared');
    }

    // =============================================
    // PUBLIC API - UI MANAGEMENT
    // =============================================

    /**
     * Initialize currency selector with quality scores
     * Sets up dropdown and event listeners
     * @returns {Promise<void>}
     */
    async initializeSelector() {
        try {
            
            // Set up DOM event listeners first (even if no data yet)
            this.setupCurrencyEventListeners();
            
            // Get currency quality scores from database
            const qualityScores = await this.database.getCurrencyQualityScores();
            
            if (Object.keys(qualityScores).length === 0) {
                
                // Update the loading message
                if (this.dom.elements.currencySelect) {
                    this.dom.elements.currencySelect.innerHTML = '<option value="">No currencies loaded yet</option>';
                }
                return;
            }
            
            // Populate currency selector
            await this.populateSelector(qualityScores);
            
            
        } catch (error) {
            config.error('Error initializing currency selector:', error);
            // Show error in selector
            if (this.dom.elements.currencySelect) {
                this.dom.elements.currencySelect.innerHTML = '<option value="">Error loading currencies</option>';
            }
        }
    }

    /**
     * Show currency selector UI elements
     */
    showSelector() {
        if (this.dom.elements.currencySelector) {
            this.dom.elements.currencySelector.style.display = 'block';
        }
        if (this.dom.elements.currencyLabelContainer) {
            this.dom.elements.currencyLabelContainer.style.display = 'block';
        }
    }

    /**
     * Hide currency selector UI elements
     */
    hideSelector() {
        if (this.dom.elements.currencySelector) {
            this.dom.elements.currencySelector.style.display = 'none';
        }
        if (this.dom.elements.currencyLabelContainer) {
            this.dom.elements.currencyLabelContainer.style.display = 'none';
        }
    }

    /**
     * Populate currency selector with quality scores
     */
    async populateSelector(qualityScores) {
        if (!this.dom.elements.currencySelect) {
            return;
        }
        
        
        // Get currency mappings for display names
        const currencyMappings = new CurrencyMappings();
        
        // Convert quality scores to array and sort by quality (high to low), then alphabetically
        const sortedCurrencies = Object.entries(qualityScores)
            .map(([code, data]) => ({
                code,
                ...data,
                name: data.name || currencyMappings.getByCode(code)?.name || code
            }))
            .sort((a, b) => {
                // First sort by quality score (descending)
                if (b.score !== a.score) {
                    return b.score - a.score;
                }
                // Then alphabetically by code
                return a.code.localeCompare(b.code);
            });
        
        // Clear existing options
        this.dom.elements.currencySelect.innerHTML = '';
        
        // Add default option
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = 'Select currency...';
        this.dom.elements.currencySelect.appendChild(defaultOption);
        
        // Add currency options with quality scores and color coding
        sortedCurrencies.forEach(currency => {
            const option = document.createElement('option');
            option.value = currency.code;
            
            // Format: "USD - US Dollar"
            option.textContent = `${currency.code} - ${currency.name}`;
            
            // Add quality-based styling
            if (currency.score >= 90) {
                option.style.color = '#22c55e'; // Green
            } else if (currency.score >= 70) {
                option.style.color = '#eab308'; // Yellow  
            } else {
                option.style.color = '#f97316'; // Orange
            }
            
            this.dom.elements.currencySelect.appendChild(option);
        });
        
        // Load and set saved currency preference
        const savedCurrency = await this.database.getPreference('selectedCurrency');
        if (savedCurrency && this.dom.elements.currencySelect.querySelector(`option[value="${savedCurrency}"]`)) {
            this.dom.elements.currencySelect.value = savedCurrency;
            
            // CRITICAL: If we have calculations loaded, convert data to match saved currency
            // TODO: Fix in future session - use portfolioCalculator.getCurrentCalculations()
            if (window.app && window.app.currentCalculations && savedCurrency !== window.app.currentCalculations.currency) {
                
                // Block automatic currency restoration in synthetic mode
                const app = window.equateApp || window.app;
                if (app && app.timelineConversionStatus && app.timelineConversionStatus.synthetic) {
                    console.log('🚫 Automatic currency restoration blocked in synthetic mode');
                    // Reset to detected currency instead
                    this.dom.elements.currencySelect.value = app.detectedCurrency || '';
                    return;
                }
                
                // Trigger the existing change event to reuse all conversion logic
                this.dom.elements.currencySelect.dispatchEvent(new Event('change'));
            }
        }
        
        
        // Debug: show a few sample currencies that were added
        if (sortedCurrencies.length > 0) {
            const samples = sortedCurrencies.slice(0, 3);
        }
    }

    /**
     * Setup currency selector DOM event listeners (copied from app.js:2448-2515)
     * Handles currency selection changes and triggers recalculation
     */
    setupCurrencyEventListeners() {
        if (!this.dom.elements.currencySelect) return;
        
        // Currency selection change - with INSTANT switching support
        this.dom.elements.currencySelect.addEventListener('change', async (e) => {
            const selectedCurrency = e.target.value;
            
            // Track currency switch event
            if (typeof gtag !== 'undefined') {
                gtag('event', 'currency_switch', {
                    'currency': selectedCurrency,
                    'event_category': 'settings'
                });
            }
            
            const isChanging = this.getIsChangingCurrency();
            if (!selectedCurrency || !window.app || !window.app.currentCalculations || isChanging) {
                return;
            }
            
            // Use the unified changeCurrency method for consistency
            await this.changeCurrency(selectedCurrency);
        });
    }

    /**
     * Update all currency-formatted displays with current calculations
     * @param {Object} calculations - Current calculations object
     */
    updateDisplays(calculations) {
        if (!calculations || !this.detectedCurrency) {
            console.warn('CurrencyService: Cannot update displays - missing calculations or currency');
            return;
        }

        // Use DOMManager to update all currency displays
        this.dom.updateCurrencyMetrics(calculations, this.detectedCurrency, this.converter);
        this.dom.updatePriceSource(calculations, this.detectedCurrency, this.converter, this.translator);
    }

    // =============================================
    // PUBLIC API - BUSINESS LOGIC
    // =============================================

    /**
     * Change to a new currency
     * Handles the complete currency switching workflow with INSTANT switching for pre-computed currencies
     * @param {string} newCurrency - New currency code
     * @returns {Promise<void>}
     */
    async changeCurrency(newCurrency) {
        if (this.isChangingCurrency) {
            console.log('CurrencyService: Currency change already in progress, ignoring request');
            return;
        }

        if (!newCurrency || newCurrency === this.detectedCurrency) {
            console.log('CurrencyService: No currency change needed');
            return;
        }

        // CRITICAL: Block currency switching in synthetic mode
        // Check both global window references for synthetic mode detection
        const app = window.equateApp || window.app;
        if (app && app.timelineConversionStatus && app.timelineConversionStatus.synthetic) {
            console.warn('🚫 Currency switching blocked: System is in synthetic mode. Multi-currency conversion not available for this portfolio.');
            
            // Reset selector to prevent confusion
            if (this.dom.elements.currencySelect) {
                this.dom.elements.currencySelect.value = app.detectedCurrency || '';
            }
            
            return;
        }

        this.isChangingCurrency = true;

        try {
            console.log('🚀 V2 Dynamic Currency System: Instant switching to', newCurrency);
            this.eventBus.emitSync(EventBus.Events.LOADING_STARTED, `Switching to ${newCurrency}...`);

            // Save preference
            await this.database.savePreference('selectedCurrency', newCurrency);
            
            // Update internal state
            this.setDetectedCurrency(newCurrency);
            
            // V2 INSTANT SWITCHING: Use calculator's new changeCurrency method
            this.calculator.changeCurrency(newCurrency);
            
            // Update cached portfolio data with new currency
            const portfolioData = await this.database.getPortfolioData();
            if (portfolioData) {
                portfolioData.currency = newCurrency;
                
                await this.database.savePortfolioData(
                    portfolioData.portfolioData,
                    portfolioData.transactionData,
                    portfolioData.userId,
                    portfolioData.isEnglish,
                    portfolioData.company,
                    portfolioData.portfolioData?.detectedInfo?.currency, // Preserve original currency
                    portfolioData.transactionData?.detectedInfo?.company
                );

                // Get historical data for timeline from multiCurrencyPrices
                const multiCurrencyData = await this.database.getMultiCurrencyPrices();
                if (multiCurrencyData && multiCurrencyData.length > 0) {
                    const instantPrices = multiCurrencyData
                        .map(entry => ({
                            date: entry.date,
                            price: entry[newCurrency]
                        }))
                        .filter(entry => entry.price != null && entry.price > 0);
                    
                    // Set historical prices for timeline
                    this.calculator.historicalPrices = instantPrices;
                    
                    // Update current price to reflect the new currency
                    this.calculator.updateCurrentPrice();
                }
                
                // Quick recalculation with V2 enhanced reference points
                const calculations = await this.calculator.calculate();
                
                // Update UI directly
                if (window.app) {
                    window.app.currentCalculations = calculations;
                    window.app.displayResults();
                    window.app.createPortfolioChart();
                    window.app.createBreakdownCharts();
                }
                
                console.log('💥 V2 INSTANT currency switch completed!');
            }

        } catch (error) {
            console.error('CurrencyService: Error changing currency:', error);
            throw error;
        } finally {
            this.isChangingCurrency = false;
            this.eventBus.emitSync(EventBus.Events.LOADING_FINISHED);
        }
    }

    /**
     * Show currency warning if there's a mismatch
     * @param {Object} calculations - Current calculations with warning info
     */
    showWarning() {
        // Get current calculations from app for now (will be refactored in Phase 4.x)
        const currentCalculations = window.app ? window.app.currentCalculations : null;
        
        if (!currentCalculations || !currentCalculations.currencyWarning) {
            return;
        }

        // Check if warning already exists
        let warningElement = document.getElementById('currencyWarning');
        
        if (!warningElement) {
            // Create warning element
            warningElement = document.createElement('div');
            warningElement.id = 'currencyWarning';
            warningElement.className = 'historical-price-warning';
            warningElement.innerHTML = `
                <div class="warning-content">
                    <div class="warning-icon">⚠️</div>
                    <div class="warning-text">
                        <strong>Currency Mismatch Warning</strong>
                        <p>${currentCalculations.currencyWarning}</p>
                    </div>
                    <button class="warning-close" onclick="this.parentElement.parentElement.remove()">×</button>
                </div>
            `;
            
            // Insert after price source indicator
            const priceSource = document.getElementById('priceSource');
            if (priceSource && priceSource.parentNode) {
                priceSource.parentNode.insertBefore(warningElement, priceSource.nextSibling);
            }
        } else {
            // Update existing warning
            const warningText = warningElement.querySelector('.warning-text p');
            if (warningText) {
                warningText.textContent = currentCalculations.currencyWarning;
            }
        }
    }

    /**
     * Get currency match status for transaction files
     * @param {string} transactionCurrency - Currency from transaction file
     * @returns {string} Match status message
     */
    getMatchStatus(transactionCurrency) {
        if (!this.detectedCurrency) {
            return ' (no portfolio currency detected)';
        }
        return transactionCurrency === this.detectedCurrency ? ' (matches portfolio)' : ' (⚠️ mismatch)';
    }


    // =============================================
    // INTERNAL EVENT HANDLERS
    // =============================================

    /**
     * Handle recalculation requests from other services
     * @private
     */
    async handleRecalculationRequest() {
        try {
            // Trigger calculator recalculation
            const calculations = await this.calculator.calculate();
            
            // Update displays with new calculations
            this.updateDisplays(calculations);
            
            // Notify other services that calculations are updated
            this.eventBus.emitSync(EventBus.Events.CALCULATIONS_UPDATED, calculations);
            
        } catch (error) {
            console.error('CurrencyService: Error handling recalculation request:', error);
        }
    }

    /**
     * Handle portfolio data loading events
     * @param {Object} portfolioData - Loaded portfolio data
     * @private
     */
    handlePortfolioDataLoaded(portfolioData) {
        if (portfolioData && portfolioData.currency) {
            this.setDetectedCurrency(portfolioData.currency);
        }
    }

    /**
     * Get ALL currency columns from ALL entries, not just the first entry
     * This ensures sparse currencies are included even if they don't appear in the first entry
     * @param {Array} currencyEntries - Currency data entries
     * @returns {Array} Array of all unique currency column names (excluding 'date')
     */
    getAllColumns(currencyEntries) {
        if (!currencyEntries || currencyEntries.length === 0) {
            return [];
        }
        
        const allColumns = new Set();
        
        // Iterate through ALL entries to collect ALL possible currency columns
        currencyEntries.forEach(entry => {
            if (entry && typeof entry === 'object') {
                Object.keys(entry).forEach(key => {
                    if (key !== 'date') {
                        allColumns.add(key);
                    }
                });
            }
        });
        
        const result = Array.from(allColumns).sort();
        config.debug('🔍 getAllCurrencyColumns found:', result.length, 'unique currencies from', currencyEntries.length, 'entries');
        config.debug('🔍 All currencies:', result.join(', '));
        
        return result;
    }

    /**
     * Calculate portfolio-aware currency quality scores
     * Uses portfolio timeline if available, otherwise falls back to full dataset calculation
     * @param {Array} currencyEntries - Currency data entries
     * @param {Array} currencyColumns - Available currency columns
     * @param {Object} portfolioData - Portfolio data with entries (optional)
     * @returns {Object} Quality scores object
     */
    calculatePortfolioAwareScores(currencyEntries, currencyColumns, portfolioData = null) {
        config.debug('🎯 Calculating portfolio-aware currency quality scores');
        
        if (!currencyEntries || currencyEntries.length === 0) {
            config.debug('No currency entries provided, returning empty scores');
            return {};
        }

        // Use provided currency columns (from Excel headers) or fallback to discovering from data entries  
        const allCurrencies = currencyColumns || this.getAllColumns(currencyEntries);
        const allCurrenciesWithEUR = ['EUR', ...allCurrencies.filter(c => c !== 'EUR')];
        const qualityScores = {};

        // If portfolio data is available, use portfolio timeline for calculation
        if (portfolioData && portfolioData.entries) {
            config.debug('Using portfolio timeline for quality calculation');
            
            // Get portfolio timeline from allocation dates
            const allocationDates = portfolioData.entries
                .filter(entry => entry.allocationDate)
                .map(entry => entry.allocationDate)
                .sort();
                
            if (allocationDates.length > 0) {
                const portfolioStartDate = allocationDates[0];
                const portfolioEndDate = allocationDates[allocationDates.length - 1];
                config.debug(`Portfolio timeline: ${portfolioStartDate} to ${portfolioEndDate}`);
                
                // Get trading days in portfolio timeline from currency data
                const portfolioTradingDates = new Set();
                currencyEntries.forEach(entry => {
                    if (entry.date >= portfolioStartDate && entry.date <= portfolioEndDate) {
                        portfolioTradingDates.add(entry.date);
                    }
                });
                
                config.debug(`Found ${portfolioTradingDates.size} trading days in portfolio timeline`);
                
                // Calculate portfolio-specific quality scores
                allCurrenciesWithEUR.forEach(currency => {
                    if (currency === 'EUR') {
                        qualityScores[currency] = {
                            score: 100,
                            availableDays: portfolioTradingDates.size,
                            totalTimelineDays: portfolioTradingDates.size,
                            name: 'Euro (Base Currency)',
                            calculationMethod: 'portfolio-specific',
                            dataRange: {
                                start: currencyEntries[0]?.date,
                                end: currencyEntries[currencyEntries.length - 1]?.date
                            }
                        };
                    } else {
                        const currencyMatchingDates = currencyEntries.filter(entry => 
                            portfolioTradingDates.has(entry.date) && 
                            entry[currency] && !isNaN(entry[currency]) && entry[currency] > 0
                        );
                        
                        const qualityScore = portfolioTradingDates.size > 0 ? 
                            Math.round((currencyMatchingDates.length / portfolioTradingDates.size) * 100) : 0;
                        
                        qualityScores[currency] = {
                            score: qualityScore,
                            availableDays: currencyMatchingDates.length,
                            totalTimelineDays: portfolioTradingDates.size,
                            calculationMethod: 'portfolio-specific',
                            dataRange: {
                                start: currencyEntries[0]?.date,
                                end: currencyEntries[currencyEntries.length - 1]?.date
                            }
                        };
                        
                        config.debug(`${currency}: ${currencyMatchingDates.length}/${portfolioTradingDates.size} = ${qualityScore}%`);
                    }
                });
                
                return qualityScores;
            }
        }

        // V2: Return empty object - quality scores not needed with good data quality
        config.debug('V2: Skipping quality score calculation (not needed with enhanced reference points)');
        return {};
    }

    // =============================================
    // DIAGNOSTIC METHODS
    // =============================================

    /**
     * Get service status and diagnostic information
     * @returns {Object} Service status information
     */
    getStatus() {
        return {
            detectedCurrency: this.detectedCurrency,
            isChangingCurrency: this.isChangingCurrency,
            hasCachedData: !!this.cachedCurrencyData,
            cachedDataLength: this.cachedCurrencyData ? this.cachedCurrencyData.length : 0,
            dependencies: {
                database: !!this.database,
                converter: !!this.converter,
                mappings: !!this.mappings,
                translator: !!this.translator,
                calculator: !!this.calculator,
                eventBus: !!this.eventBus,
                domManager: !!this.dom
            }
        };
    }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CurrencyService;
} else if (typeof window !== 'undefined') {
    window.CurrencyService = CurrencyService;
}