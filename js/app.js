/**
 * Main Application Logic for Equate Portfolio Analysis
 * Coordinates all components and handles user interactions
 */

class EquateApp {
    constructor() {
        this.isInitialized = false;
        this.hasData = false;
        this.currentCalculations = null;
        this.elements = {};
    }

    /**
     * Show user-friendly error message
     * @param {string} message - Error message to display
     * @param {string} details - Optional technical details
     */
    showError(message, details = '') {
        const errorDiv = document.createElement('div');
        errorDiv.style.cssText = `
            position: fixed; top: 20px; right: 20px; z-index: 10000;
            background: #ef4444; color: white; padding: 15px 20px;
            border-radius: 8px; max-width: 400px; box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        `;
        errorDiv.innerHTML = `
            <strong>⚠️ ${message}</strong>
            ${details ? `<br><small style="opacity: 0.8;">${details}</small>` : ''}
            <button onclick="this.parentElement.remove()" style="
                background: none; border: none; color: white; float: right; 
                margin-left: 10px; cursor: pointer; font-size: 18px;
            ">×</button>
        `;
        document.body.appendChild(errorDiv);
        
        // Auto-remove after 10 seconds
        setTimeout(() => {
            if (errorDiv.parentElement) {
                errorDiv.remove();
            }
        }, 10000);
    }

    /**
     * Initialize the application
     */
    async init() {
        try {
            this.showLoading('Initializing application...');
            
            // Initialize translation system
            await translationManager.init();
            config.debug('✅ Translation system initialized');
            
            // Initialize database with stability checks
            await equateDB.init();
            config.debug('✅ Database initialized');
            
            // Run stability checks
            await this.runStabilityChecks();

            // Get DOM elements
            this.cacheElements();
            
            // Set up event listeners
            this.setupEventListeners();
            this.setupSectionReordering();
            
            
            // Load user preferences
            await this.loadUserPreferences();
            
            // Initialize Results tab state
            toggleResultsDisplay(false);
            
            // Check for cached data
            await this.checkCachedData();
            
            // Load cached manual price if available
            await this.loadCachedManualPrice();
            
            this.isInitialized = true;
            this.hideLoading();
            
            config.debug('Application initialized successfully');
            
        } catch (error) {
            config.error('Failed to initialize application:', error);
            this.hideLoading();
            this.showError('Failed to initialize application: ' + error.message);
        }
    }

    /**
     * Run comprehensive stability checks
     */
    async runStabilityChecks() {
        config.debug('🔧 Running stability checks...');
        
        try {
            // Test basic database operations
            config.debug('1. Testing database write/read...');
            await equateDB.savePreference('test_key', 'test_value');
            const testValue = await equateDB.getPreference('test_key');
            if (testValue !== 'test_value') {
                throw new Error('Database read/write test failed');
            }
            config.debug('✅ Database read/write working');
            
            // Check existing data
            config.debug('2. Checking existing portfolio data...');
            const portfolioData = await equateDB.getPortfolioData();
            if (portfolioData) {
                config.debug('✅ Found existing portfolio data:', {
                    userId: portfolioData.userId,
                    uploadDate: portfolioData.uploadDate,
                    hasTransactions: portfolioData.hasTransactions
                });
            } else {
                config.debug('ℹ️ No existing portfolio data');
            }
            
            // Check historical prices
            config.debug('3. Checking historical prices...');
            const pricesCount = await equateDB.getHistoricalPricesCount();
            config.debug(`ℹ️ Historical prices in database: ${pricesCount}`);
            
            // Check browser compatibility
            config.debug('4. Checking browser compatibility...');
            const compatibility = {
                indexedDB: !!window.indexedDB,
                fetch: !!window.fetch,
                xlsx: !!window.XLSX,
                plotly: !!window.Plotly
            };
            config.debug('Browser compatibility:', compatibility);
            
            const missingFeatures = Object.entries(compatibility)
                .filter(([key, value]) => !value)
                .map(([key]) => key);
                
            if (missingFeatures.length > 0) {
                config.warn('⚠️ Missing browser features:', missingFeatures);
            } else {
                config.debug('✅ All required browser features available');
            }
            
        } catch (error) {
            config.error('❌ Stability check failed:', error);
            this.showError('Application stability check failed: ' + error.message);
        }
    }

    /**
     * Cache DOM elements
     */
    cacheElements() {
        try {
            this.elements = {
                // Theme and Format
                themeSelect: document.getElementById('theme'),
                formatSelect: document.getElementById('formatSelect'),
                languageSelect: document.getElementById('languageSelect'),
                
                // Price controls (moved to Results tab)
                manualPrice: document.getElementById('manualPriceInput'),
                manualPriceInput: document.getElementById('manualPriceInput'), // New location in Results tab
                resultsExportBtn: document.getElementById('resultsExportBtn'),
                
                // Upload
                portfolioInput: document.getElementById('portfolio'),
                transactionsInput: document.getElementById('transactions'),
                portfolioZone: document.getElementById('portfolioZone'),
                transactionsZone: document.getElementById('transactionsZone'),
                portfolioInfo: document.getElementById('portfolioInfo'),
                transactionsInfo: document.getElementById('transactionsInfo'),
                analyzeBtn: document.getElementById('analyzeBtn'),
                
                // Settings - removed, now hardcoded
                
                // Results
                priceSource: document.getElementById('priceSource'),
                
                // Metrics
                userInvestment: document.getElementById('userInvestment'),
                companyMatch: document.getElementById('companyMatch'),
                freeShares: document.getElementById('freeShares'),
                dividendIncome: document.getElementById('dividendIncome'),
                totalInvestment: document.getElementById('totalInvestment'),
                totalSold: document.getElementById('totalSold'),
                currentValue: document.getElementById('currentValue'),
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

            // Check for critical missing elements
            const criticalElements = ['portfolioZone', 'transactionsZone', 'analyzeBtn'];
            const missingElements = criticalElements.filter(elementKey => !this.elements[elementKey]);
            
            if (missingElements.length > 0) {
                config.warn('Missing critical elements:', missingElements);
            }

        } catch (error) {
            config.error('Error caching DOM elements:', error);
            this.elements = {};
        }
    }

    /**
     * Set up event listeners
     */
    setupEventListeners() {
        // Theme switching
        if (this.elements.themeSelect) {
            this.elements.themeSelect.addEventListener('change', (e) => {
                this.switchTheme(e.target.value);
            });
        }

        // Language switching
        if (this.elements.languageSelect) {
            this.elements.languageSelect.addEventListener('change', (e) => {
                this.setLanguagePreference(e.target.value);
            });
        }

        // Drag and drop
        this.setupDragAndDrop();

        // Manual price input
        if (this.elements.manualPrice) {
            // Manual price input - just save the value, don't calculate yet
            this.elements.manualPrice.addEventListener('input', (e) => {
                // Just cache the value when user types, don't calculate yet
                if (parseFloat(e.target.value) > 0) {
                    equateDB.saveManualPrice(parseFloat(e.target.value));
                }
            });
        }

        // Settings removed - now using auto-detected currency, hardcoded 2 decimals, and Plotly built-in chart controls

        // Analyze button
        if (this.elements.analyzeBtn) {
            this.elements.analyzeBtn.addEventListener('click', () => {
                this.generateAnalysis();
            });
        }
    }

    /**
     * Set up drag and drop functionality
     */
    setupDragAndDrop() {
        ['portfolioZone', 'transactionsZone'].forEach(zoneId => {
            const zone = document.getElementById(zoneId);
            const input = zone.querySelector('input[type="file"]');
            
            // Prevent default behaviors
            zone.addEventListener('dragenter', (e) => {
                e.preventDefault();
                e.stopPropagation();
                zone.classList.add('drag-hover');
            });
            
            zone.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.stopPropagation();
                zone.classList.add('drag-hover');
            });
            
            zone.addEventListener('dragleave', (e) => {
                e.preventDefault();
                e.stopPropagation();
                // Only remove class if we're leaving the zone entirely
                if (!zone.contains(e.relatedTarget)) {
                    zone.classList.remove('drag-hover');
                }
            });
            
            zone.addEventListener('drop', (e) => {
                e.preventDefault();
                e.stopPropagation();
                zone.classList.remove('drag-hover');
                
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    input.files = files;
                    this.handleFileSelect(input, zoneId);
                }
            });
            
            // Click to select file
            zone.addEventListener('click', () => {
                input.click();
            });
            
            // Handle file selection
            input.addEventListener('change', (e) => {
                this.handleFileSelect(e.target, zoneId);
            });
        });
    }

    /**
     * Set up section reordering functionality
     */
    setupSectionReordering() {
        const sectionIcons = document.getElementById('sectionIcons');
        if (!sectionIcons) return;

        const icons = sectionIcons.querySelectorAll('.section-icon');
        
        // Load saved section order from localStorage or use default order
        const savedOrder = localStorage.getItem('sectionOrder');
        if (savedOrder) {
            const order = JSON.parse(savedOrder);
            this.applySectionOrder(order);
        } else {
            // Apply default order to ensure sections are properly organized
            const defaultOrder = ['cards', 'timeline', 'bar', 'pie'];
            this.applySectionOrder(defaultOrder);
        }

        // Add drag and drop event listeners to each icon
        icons.forEach(icon => {
            icon.addEventListener('dragstart', (e) => {
                icon.classList.add('dragging');
                e.dataTransfer.setData('text/plain', icon.dataset.section);
                e.dataTransfer.effectAllowed = 'move';
            });

            icon.addEventListener('dragend', () => {
                icon.classList.remove('dragging');
                document.querySelectorAll('.section-icon').forEach(i => i.classList.remove('drag-over'));
            });

            icon.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
            });

            icon.addEventListener('dragenter', (e) => {
                e.preventDefault();
                if (!icon.classList.contains('dragging')) {
                    icon.classList.add('drag-over');
                }
            });

            icon.addEventListener('dragleave', (e) => {
                if (!icon.contains(e.relatedTarget)) {
                    icon.classList.remove('drag-over');
                }
            });

            icon.addEventListener('drop', (e) => {
                e.preventDefault();
                icon.classList.remove('drag-over');
                
                const draggedSection = e.dataTransfer.getData('text/plain');
                const targetSection = icon.dataset.section;
                
                if (draggedSection !== targetSection) {
                    this.reorderSections(draggedSection, targetSection);
                }
            });
        });
    }

    /**
     * Reorder sections based on drag and drop
     */
    reorderSections(draggedSection, targetSection) {
        const sectionIcons = document.getElementById('sectionIcons');
        const draggedIcon = sectionIcons.querySelector(`[data-section="${draggedSection}"]`);
        const targetIcon = sectionIcons.querySelector(`[data-section="${targetSection}"]`);
        
        if (draggedIcon && targetIcon) {
            // Get current order
            const currentOrder = Array.from(sectionIcons.children).map(icon => icon.dataset.section);
            
            // Find positions
            const draggedIndex = currentOrder.indexOf(draggedSection);
            const targetIndex = currentOrder.indexOf(targetSection);
            
            // Remove dragged item and insert at target position
            currentOrder.splice(draggedIndex, 1);
            currentOrder.splice(targetIndex, 0, draggedSection);
            
            // Apply new order
            this.applySectionOrder(currentOrder);
            
            // Save to localStorage
            localStorage.setItem('sectionOrder', JSON.stringify(currentOrder));
            
            config.debug('Section order updated:', currentOrder);
        }
    }

    /**
     * Apply section order to both icons and content sections
     */
    applySectionOrder(order) {
        const sectionIcons = document.getElementById('sectionIcons');
        const resultsContent = document.getElementById('resultsContent');
        
        if (!sectionIcons || !resultsContent) return;

        // Reorder icons
        order.forEach(sectionName => {
            const icon = sectionIcons.querySelector(`[data-section="${sectionName}"]`);
            if (icon) {
                sectionIcons.appendChild(icon);
            }
        });

        // Reorder content sections by finding chart containers through their specific chart IDs
        const sectionMapping = {
            'cards': resultsContent.querySelector('.metrics-section'),
            'timeline': document.querySelector('#portfolioChart')?.closest('.chart-container'),
            'bar': document.querySelector('#performanceBarChart')?.closest('.chart-container'),
            'pie': document.querySelector('#investmentPieChart')?.closest('.chart-container')
        };

        // Debug: Log what sections were found
        config.debug('Section mapping found:', {
            cards: !!sectionMapping.cards,
            timeline: !!sectionMapping.timeline,
            bar: !!sectionMapping.bar,
            pie: !!sectionMapping.pie
        });

        // Create a temporary container to preserve the results-controls
        const tempContainer = document.createElement('div');
        const resultsControls = resultsContent.querySelector('.results-controls');
        
        // Move results-controls to temp container to preserve it
        if (resultsControls) {
            tempContainer.appendChild(resultsControls);
        }

        // Reorder sections according to the order array
        order.forEach(sectionName => {
            const section = sectionMapping[sectionName];
            if (section) {
                config.debug(`Moving section: ${sectionName}`);
                tempContainer.appendChild(section);
            } else {
                config.warn(`Section not found: ${sectionName}`);
            }
        });

        // Clear results content and add everything back in order
        resultsContent.innerHTML = '';
        while (tempContainer.firstChild) {
            resultsContent.appendChild(tempContainer.firstChild);
        }
    }

    /**
     * Apply section order after charts are created (timing-safe version)
     */
    applySectionOrderAfterChartsCreated() {
        const savedOrder = localStorage.getItem('sectionOrder');
        if (savedOrder) {
            const order = JSON.parse(savedOrder);
            this.applySectionOrder(order);
        } else {
            // Apply default order to ensure sections are properly organized
            const defaultOrder = ['cards', 'timeline', 'bar', 'pie'];
            this.applySectionOrder(defaultOrder);
        }
    }

    /**
     * Handle file selection
     */
    async handleFileSelect(input, zoneId) {
        const zone = document.getElementById(zoneId);
        const infoDiv = document.getElementById(zoneId.replace('Zone', 'Info'));
        
        if (input.files.length > 0) {
            const file = input.files[0];
            zone.classList.add('file-selected');
            infoDiv.innerHTML = `✓ ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
            infoDiv.classList.add('show');
        } else {
            zone.classList.remove('file-selected');
            infoDiv.classList.remove('show');
        }
        
        this.checkFilesReady();
    }

    /**
     * Check if required files are ready for analysis
     */
    checkFilesReady() {
        const portfolioFile = this.elements.portfolioInput.files.length > 0;
        
        if (portfolioFile || this.hasData) {
            this.elements.analyzeBtn.disabled = false;
            if (this.hasData) {
                this.elements.analyzeBtn.textContent = 'Regenerate Analysis';
                // Enable recalculate button in Quick Actions
                if (this.elements.recalculateBtn) {
                    this.elements.recalculateBtn.disabled = false;
                }
            } else {
                this.elements.analyzeBtn.textContent = 'Generate Analysis';
                // Disable recalculate button until first analysis
                if (this.elements.recalculateBtn) {
                    this.elements.recalculateBtn.disabled = true;
                }
            }
        } else {
            this.elements.analyzeBtn.disabled = true;
            this.elements.analyzeBtn.textContent = 'Please upload Portfolio Details';
            // Disable recalculate button
            if (this.elements.recalculateBtn) {
                this.elements.recalculateBtn.disabled = true;
            }
        }
    }

    /**
     * Generate portfolio analysis
     */
    async generateAnalysis() {
        try {
            this.showLoading('Processing your portfolio data...');

            let portfolioData, transactionData, company = 'Other', isEnglish = true; // Default to Other (safer)
            let analysisResult = null;

            // Check if we have files to parse or should use cached data
            const portfolioFile = this.elements.portfolioInput.files[0];
            
            if (portfolioFile) {
                // Parse uploaded files (new upload) - CACHE-AWARE ANALYZER APPROACH
                config.debug('🚨 TAKING NEW UPLOAD PATH - using FileAnalyzer for unified analysis');
                let transactionFile = this.elements.transactionsInput.files[0] || null;
                const originalTransactionFile = transactionFile; // Remember original state for UI clearing

                // Get cached data for FileAnalyzer to handle all validation
                const cachedData = await equateDB.getPortfolioData();

                // FileAnalyzer handles all validation, cache checks, and compatibility
                config.debug('🔍 Starting cache-aware unified file set analysis...');
                analysisResult = await fileAnalyzer.analyzeFileSet(portfolioFile, transactionFile, cachedData);
                
                // Handle FileAnalyzer results
                if (analysisResult.shouldClearCachedTransactions) {
                    config.debug('🗑️ Clearing cached transaction data due to currency mismatch');
                    await equateDB.clearTransactionData();
                    
                    // Immediately update drop zone UI to reflect cache clearing
                    const clearedDbInfo = await equateDB.getDatabaseInfo();
                    this.updateDropZonesForCachedData(clearedDbInfo);
                }
                
                // Clear transaction file input if it was rejected by FileAnalyzer
                if (originalTransactionFile && analysisResult.errors.some(e => e.type === 'CURRENCY_MISMATCH')) {
                    config.debug('🗑️ Clearing rejected transaction file from UI');
                    this.elements.transactionsInput.value = ''; // Clear the file input
                    this.handleFileSelect(this.elements.transactionsInput, 'transactionsZone'); // Update visual styling
                }
                
                config.debug('📊 File Analysis Results:', fileAnalyzer.getAnalysisSummary(analysisResult));

                // Handle analysis errors (currency mismatches, etc.)
                if (analysisResult.errors.length > 0) {
                    this.handleAnalysisErrors(analysisResult.errors);
                    // Continue with portfolio-only analysis if transaction file was rejected
                    transactionFile = null;
                }

                // Display warnings to user
                if (analysisResult.warnings.length > 0) {
                    this.displayAnalysisWarnings(analysisResult.warnings);
                }

                // Step 2: Apply UI settings from analysis
                this.applyUISettings(analysisResult.unified);
                
                // Legacy compatibility
                isEnglish = (analysisResult.unified.language === 'english');

                // Step 3: Detect date format before parsing any data
                const detectedFormat = await fileParser.detectDateFormat(portfolioFile, transactionFile);
                
                // If ambiguous, ask user to choose
                if (detectedFormat === null) {
                    config.debug('📅 All string dates are ambiguous, asking user...');
                    await fileParser.promptUserForDateFormat();
                } else if (detectedFormat === 'SKIP') {
                    config.debug('📅 No string dates found, skipping format detection');
                }

                // Step 4: Parse files with analysis context
                config.debug('📄 Parsing files with pre-analysis context...');
                portfolioData = await fileParser.parsePortfolioFile(portfolioFile, analysisResult);
                config.debug('Portfolio data parsed:', portfolioData);

                // Parse transaction data if compatible
                if (transactionFile && analysisResult.errors.length === 0) {
                    transactionData = await fileParser.parseTransactionFile(transactionFile, analysisResult);
                    fileParser.validateUserIds();
                    config.debug('Transaction data parsed:', transactionData);
                }

                // Step 5: Enhanced company detection using both files
                company = fileParser.detectCompanyFromBothFiles(portfolioData, transactionData);
                
                config.debug('🔍 Final Detection Results:', {
                    language: analysisResult.unified.language,
                    currency: analysisResult.unified.currency, 
                    company: company,
                    hasTransactions: !!transactionData
                });

                // Save to database with detection results including currency
                await equateDB.savePortfolioData(portfolioData, transactionData, portfolioData.userId, isEnglish, company, this.detectedCurrency);
                
                // Update drop zone indicators to reflect current cache state (especially after transaction rejection)
                const updatedDbInfo = await equateDB.getDatabaseInfo();
                this.updateDropZonesForCachedData(updatedDbInfo);
            } else if (this.hasData) {
                // Use cached data (regenerate with existing data)
                config.debug('🚨 TAKING CACHED DATA PATH - using existing data');
                config.debug('🔄 Using cached data for regeneration...');
                const cachedData = await equateDB.getPortfolioData();
                if (!cachedData) {
                    throw new Error('No cached data found. Please upload portfolio files.');
                }
                portfolioData = cachedData.portfolioData;
                transactionData = cachedData.transactionData;
                
                // Use cached detection results (with fallback for older cached data)
                isEnglish = cachedData.isEnglish !== undefined ? cachedData.isEnglish : true; // Update outer scope variable
                company = cachedData.company || 'Other'; // Update outer scope variable (default to Other)
                
                config.debug('🔍 Cached Detection Results - English:', isEnglish, ', Company:', company);
            } else {
                throw new Error('No files uploaded and no cached data available.');
            }

            // Set up calculator
            portfolioCalculator.setPortfolioData(portfolioData, transactionData, this.detectedCurrency);
            
            // Debug: Check company value before calling loadHistoricalPrices
            config.debug('🔍 About to call loadHistoricalPrices with company:', company);
            
            // Ensure historical prices are loaded (conditionally based on company)
            await this.loadHistoricalPrices(company);
            
            // Use historical prices only for Allianz companies
            if (company === 'Allianz') {
                const historicalPrices = await equateDB.getHistoricalPrices();
                config.debug('Retrieved historical prices from DB:', {
                    count: historicalPrices.length,
                    firstFew: historicalPrices.slice(0, 3),
                    lastFew: historicalPrices.slice(-3)
                });
                
                if (historicalPrices.length > 0) {
                    config.debug('Setting historical prices on calculator');
                    await portfolioCalculator.setHistoricalPrices(historicalPrices);
                } else {
                    // Fallback: use market price from portfolio data if no historical prices
                    config.warn('No historical prices available, using portfolio market price');
                    const marketPrice = portfolioData.entries[0]?.marketPrice || 355.30;
                    await portfolioCalculator.setManualPrice(marketPrice);
                }
            } else {
                // For non-Allianz companies, only use market price from portfolio
                config.debug(`🏢 Company is ${company}, using only portfolio market price (no historical prices)`);
                const marketPrice = portfolioData.entries[0]?.marketPrice || 355.30;
                await portfolioCalculator.setManualPrice(marketPrice);
            }

            // Calculate metrics (after AsOfDate is saved to IndexedDB)
            this.currentCalculations = await portfolioCalculator.calculate();
            config.debug('Calculations completed:', this.currentCalculations);

            // Update UI
            this.displayResults();
            config.debug('🔄 Creating portfolio chart with transaction data:', {
                hasTransactionData: !!transactionData,
                transactionCount: transactionData ? transactionData.entries.length : 0
            });
            this.createPortfolioChart();
            this.createBreakdownCharts();
            this.hasData = true;
            
            // Show results in Results tab
            toggleResultsDisplay(true);
            switchTab('results');

            // Update button state
            this.checkFilesReady();

            this.hideLoading();

        } catch (error) {
            config.error('Analysis failed:', error);
            this.hideLoading();
            this.showError('Analysis failed: ' + error.message);
        }
    }

    /**
     * Apply UI settings from file analysis result
     * @param {Object} unifiedAnalysis - Unified analysis result
     */
    applyUISettings(unifiedAnalysis) {
        config.debug('🎨 Applying UI settings from analysis:', unifiedAnalysis);
        
        // Set UI language (user can still override via language selector)
        if (unifiedAnalysis.language !== translationManager.getCurrentLanguage()) {
            config.debug('🌐 Auto-switching UI language to:', unifiedAnalysis.language);
            this.setLanguagePreference(unifiedAnalysis.language);
        }
        
        // Store detected currency for display formatting
        this.detectedCurrency = unifiedAnalysis.currency;
        
        config.debug('✅ UI settings applied:', {
            language: translationManager.getCurrentLanguage(),
            currency: this.detectedCurrency
        });
    }

    /**
     * Handle analysis errors (e.g., currency mismatches)
     * @param {Array} errors - Array of error objects
     */
    handleAnalysisErrors(errors) {
        config.warn('⚠️ Analysis errors detected:', errors);
        
        errors.forEach(error => {
            if (error.type === 'CURRENCY_MISMATCH') {
                this.showWarning(
                    'Currency Mismatch',
                    `${error.message}\n\n${error.resolution}`,
                    'warning'
                );
            } else {
                this.showWarning(
                    'File Analysis Error',
                    `${error.message}\n\n${error.resolution}`,
                    'error'
                );
            }
        });
    }

    /**
     * Display analysis warnings to user
     * @param {Array} warnings - Array of warning objects
     */
    displayAnalysisWarnings(warnings) {
        config.debug('ℹ️ Analysis warnings:', warnings);
        
        warnings.forEach(warning => {
            if (warning.type === 'LANGUAGE_MISMATCH') {
                this.showInfo(
                    'Language Difference Detected',
                    `${warning.message}\n\n${warning.resolution}`,
                    'info'
                );
            }
        });
    }

    /**
     * Show warning message to user
     */
    showWarning(title, message, type = 'warning') {
        // Create a simple alert for now - can be enhanced with custom UI later
        const icon = type === 'error' ? '❌' : type === 'info' ? 'ℹ️' : '⚠️';
        alert(`${icon} ${title}\n\n${message}`);
    }

    /**
     * Show info message to user
     */
    showInfo(title, message, type = 'info') {
        this.showWarning(title, message, type);
    }

    /**
     * Get currency symbol for a given currency code
     * @param {String} currency - Currency code (e.g., 'EUR', 'USD', 'GBP')
     * @returns {String} Currency symbol
     */
    getCurrencySymbol(currency) {
        if (!currency) return '';
        
        config.debug('🔍 getCurrencySymbol called with:', currency, typeof currency);
        
        if (typeof CurrencyMappings !== 'undefined') {
            const mappings = new CurrencyMappings();
            const currencyInfo = mappings.getByCode(currency);
            
            config.debug('🔍 Currency lookup result:', {
                currency,
                found: !!currencyInfo,
                symbol: currencyInfo?.symbol,
                name: currencyInfo?.name
            });
            
            if (currencyInfo && currencyInfo.symbol) {
                return currencyInfo.symbol;
            } else {
                config.debug(`⚠️ Unknown currency in getCurrencySymbol: ${currency} - using currency code as display`);
                return currency; // Use the currency code itself (e.g., "THB", "XYZ")
            }
        } else {
            console.warn('⚠️ CurrencyMappings not available in getCurrencySymbol');
            return '';
        }
    }

    /**
     * Load cached data
     */
    async loadCachedData() {
        try {
            this.showLoading('Loading cached data...');

            const cachedData = await equateDB.getPortfolioData();
            if (!cachedData) {
                throw new Error('No cached data found');
            }

            // Get currency from cached data (if it was stored)
            this.detectedCurrency = cachedData.currency || null;
            config.debug('💰 Retrieved currency from cached data:', this.detectedCurrency);
            
            // Set up calculator with cached data
            portfolioCalculator.setPortfolioData(cachedData.portfolioData, cachedData.transactionData, this.detectedCurrency);
            
            // For cached data, detect company again or default to Allianz
            const company = fileParser.detectCompany(cachedData.portfolioData);
            
            // Ensure historical prices are loaded (conditionally based on company)
            await this.loadHistoricalPrices(company);
            
            // Use historical prices only for Allianz companies (cached data)
            if (company === 'Allianz') {
                const historicalPrices = await equateDB.getHistoricalPrices();
                config.debug('Retrieved historical prices from DB (cached):', {
                    count: historicalPrices.length,
                    firstFew: historicalPrices.slice(0, 3),
                    lastFew: historicalPrices.slice(-3)
                });
                
                if (historicalPrices.length > 0) {
                    config.debug('Setting historical prices on calculator (cached)');
                    await portfolioCalculator.setHistoricalPrices(historicalPrices);
                } else {
                    // Fallback: use market price from portfolio data if no historical prices
                    config.warn('No historical prices available, using portfolio market price');
                    const marketPrice = cachedData.portfolioData.entries[0]?.marketPrice || 355.30;
                    await portfolioCalculator.setManualPrice(marketPrice);
                }
            } else {
                // For non-Allianz companies, only use market price from portfolio (cached)
                config.debug(`🏢 Company is ${company}, using only portfolio market price (no historical prices)`);
                const marketPrice = cachedData.portfolioData.entries[0]?.marketPrice || 355.30;
                await portfolioCalculator.setManualPrice(marketPrice);
            }

            // Calculate metrics (after AsOfDate is saved to IndexedDB)
            this.currentCalculations = await portfolioCalculator.calculate();

            // Update UI
            this.displayResults();
            this.createPortfolioChart();
            this.createBreakdownCharts();
            this.updateManualPriceFieldStyle(false); // Reset to historical prices styling
            this.hasData = true;
            
            // Show results in Results tab
            toggleResultsDisplay(true);
            switchTab('results');

            // Update button state
            this.checkFilesReady();

            this.hideLoading();

        } catch (error) {
            config.error('Failed to load cached data:', error);
            this.hideLoading();
            this.showError('Failed to load cached data: ' + error.message);
        }
    }

    /**
     * Load cached manual price and pre-fill input field
     */
    async loadCachedManualPrice() {
        try {
            const cachedPrice = await equateDB.getLatestManualPrice();
            if (cachedPrice && this.elements.manualPriceInput) {
                this.elements.manualPriceInput.value = cachedPrice.price;
                config.debug(`💰 Pre-filled manual price: €${cachedPrice.price}`);
            }
        } catch (error) {
            config.warn('Could not load cached manual price:', error);
        }
    }

    /**
     * Handle manual price input
     */
    async handleManualPriceInput(value) {
        const price = parseFloat(value);
        config.debug('🎯 Manual price input:', price, 'hasData:', this.hasData);
        if (price > 0 && this.hasData) {
            config.debug('🎯 Setting manual price on calculator');
            await portfolioCalculator.setManualPrice(price);
            config.debug('🎯 Calculating with new manual price...');
            this.currentCalculations = await portfolioCalculator.calculate(); // Use calculate() instead of recalculate()
            config.debug('🎯 New calculations with manual price:', this.currentCalculations);
            
            // Update UI components
            this.updateMetricsDisplay();
            this.updatePriceSource();
            this.updateManualPriceFieldStyle(true); // Mark field as active
            
            config.debug('🎯 Recreating chart with manual price data...');
            this.createPortfolioChart(); // Regenerate chart with new price
            this.createBreakdownCharts(); // Regenerate breakdown charts with new price
            
            // Auto-trigger chart autoscale after manual price recalculation
            setTimeout(() => {
                this.triggerChartAutoscale();
            }, 200);
            
            // Save manual price for future sessions
            equateDB.saveManualPrice(price);
        }
    }

    /**
     * Update manual price field styling based on active state
     */
    updateManualPriceFieldStyle(isActive) {
        if (this.elements.manualPriceInput) {
            if (isActive) {
                this.elements.manualPriceInput.style.backgroundColor = 'rgba(34, 197, 94, 0.1)';
                this.elements.manualPriceInput.style.borderColor = 'rgba(34, 197, 94, 0.5)';
                this.elements.manualPriceInput.style.boxShadow = '0 0 0 3px rgba(34, 197, 94, 0.1)';
            } else {
                this.elements.manualPriceInput.style.backgroundColor = '';
                this.elements.manualPriceInput.style.borderColor = '';
                this.elements.manualPriceInput.style.boxShadow = '';
            }
        }
    }

    /**
     * Display analysis results
     */
    displayResults() {
        if (!this.currentCalculations) return;

        this.updateMetricsDisplay();
        this.updatePriceSource();
        this.showCurrencyWarning();
        
        // Update UI language (labels only, after values are set)
        if (typeof translationManager !== 'undefined') {
            translationManager.updateUILanguage();
        }
        
        // Enable export buttons
        const exportButtons = document.querySelectorAll('.export-btn');
        exportButtons.forEach(btn => btn.disabled = false);
    }

    /**
     * Diagnostic function to check element mappings
     */
    diagnoseElements() {
        console.log('🔍 DIAGNOSTIC: Full Element to Card Mapping');
        console.log('🔍 HTML Source check:', document.documentElement.outerHTML.includes('TEST MARKER 123') ? '✅ Updated HTML loaded' : '❌ Old HTML still cached');
        
        const elements = ['userInvestment', 'companyMatch', 'freeShares', 'dividendIncome', 'totalInvestment', 
                         'totalReturn', 'currentValue', 'totalSold', 'totalValue', 'returnOnTotalInvestment',
                         'returnPercentage', 'xirrUserInvestment', 'availableShares', 'returnPercentageOnTotalInvestment', 'xirrTotalInvestment'];
        elements.forEach(elementId => {
            const element = document.getElementById(elementId);
            if (element) {
                const cardTitle = element.parentElement?.querySelector('h3')?.textContent;
                console.log(`${elementId} -> "${cardTitle}" (Value: ${element.textContent})`);
            } else {
                console.log(`${elementId} -> NOT FOUND`);
            }
        });
    }

    /**
     * Update metrics display
     */
    updateMetricsDisplay() {
        if (!this.currentCalculations) return;

        const currency = this.detectedCurrency; // Single source of truth - no fallback needed
        const decimals = 2; // Hardcoded to 2 decimals

        this.elements.userInvestment.textContent = this.formatCurrency(this.currentCalculations.userInvestment, currency, decimals);
        this.elements.companyMatch.textContent = this.formatCurrency(this.currentCalculations.companyMatch, currency, decimals);
        this.elements.freeShares.textContent = this.formatCurrency(this.currentCalculations.freeShares, currency, decimals);
        this.elements.dividendIncome.textContent = this.formatCurrency(this.currentCalculations.dividendIncome, currency, decimals);
        this.elements.totalInvestment.textContent = this.formatCurrency(this.currentCalculations.totalInvestment, currency, decimals);
        
        this.elements.currentValue.textContent = this.formatCurrency(this.currentCalculations.currentValue, currency, decimals);
        this.elements.totalSold.textContent = this.formatCurrency(this.currentCalculations.totalSold, currency, decimals);
        this.elements.totalValue.textContent = this.formatCurrency(this.currentCalculations.totalValue, currency, decimals);
        
        const totalReturn = this.currentCalculations.totalReturn;
        this.elements.totalReturn.textContent = this.formatCurrency(totalReturn, currency, decimals);
        this.elements.totalReturn.className = `metric-value ${totalReturn >= 0 ? 'positive' : 'negative'}`;
        
        const returnOnTotalInvestment = this.currentCalculations.returnOnTotalInvestment;
        this.elements.returnOnTotalInvestment.textContent = this.formatCurrency(returnOnTotalInvestment, currency, decimals);
        this.elements.returnOnTotalInvestment.className = `metric-value ${returnOnTotalInvestment >= 0 ? 'positive' : 'negative'}`;
        
        const returnPercentage = this.currentCalculations.returnPercentage;
        this.elements.returnPercentage.textContent = this.formatPercentage(returnPercentage, decimals);
        this.elements.returnPercentage.className = `metric-value ${returnPercentage >= 0 ? 'positive' : 'negative'}`;
        
        const returnPercentageOnTotalInvestment = this.currentCalculations.returnPercentageOnTotalInvestment;
        this.elements.returnPercentageOnTotalInvestment.textContent = this.formatPercentage(returnPercentageOnTotalInvestment, decimals);
        this.elements.returnPercentageOnTotalInvestment.className = `metric-value ${returnPercentageOnTotalInvestment >= 0 ? 'positive' : 'negative'}`;
        
        // XIRR metrics with styling
        this.elements.xirrUserInvestment.textContent = this.formatPercentage(this.currentCalculations.xirrUserInvestment, decimals);
        this.elements.xirrUserInvestment.className = `metric-value ${this.currentCalculations.xirrUserInvestment >= 0 ? 'positive' : 'negative'}`;
        
        this.elements.xirrTotalInvestment.textContent = this.formatPercentage(this.currentCalculations.xirrTotalInvestment, decimals);
        this.elements.xirrTotalInvestment.className = `metric-value ${this.currentCalculations.xirrTotalInvestment >= 0 ? 'positive' : 'negative'}`;
        
        this.elements.availableShares.textContent = this.currentCalculations.availableShares.toString();
    }

    /**
     * Update price source indicator
     */
    updatePriceSource() {
        if (!this.currentCalculations) return;

        const indicator = this.elements.priceSource;
        const text = indicator.querySelector('.indicator-text');
        
        const currency = this.detectedCurrency;
        const currencySymbol = this.getCurrencySymbol(currency);
        
        if (this.currentCalculations.priceSource === 'manual') {
            indicator.classList.add('manual');
            text.innerHTML = `<strong>${translationManager.t('price_label')}:</strong> ${currencySymbol}${this.currentCalculations.currentPrice.toFixed(2)} (manual input)`;
        } else {
            indicator.classList.remove('manual');
            text.innerHTML = `<strong>${translationManager.t('price_label')}:</strong> ${currencySymbol}${this.currentCalculations.currentPrice.toFixed(2)} ${translationManager.t('as_of_label')} ${this.currentCalculations.priceDate}`;
        }
    }

    /**
     * Show currency warning if there's a mismatch
     */
    showCurrencyWarning() {
        if (!this.currentCalculations || !this.currentCalculations.currencyWarning) {
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
                        <p>${this.currentCalculations.currencyWarning}</p>
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
                warningText.textContent = this.currentCalculations.currencyWarning;
            }
        }
    }


    /**
     * Collapse a specific section
     */
    collapseSection(sectionId) {
        const content = document.getElementById(sectionId);
        const header = content?.parentElement.querySelector('.section-header');
        
        if (content && header) {
            content.classList.add('collapsed');
            header.classList.add('collapsed');
        }
    }

    /**
     * Expand a specific section
     */
    expandSection(sectionId) {
        const content = document.getElementById(sectionId);
        const header = content?.parentElement.querySelector('.section-header');
        
        if (content && header) {
            content.classList.remove('collapsed');
            header.classList.remove('collapsed');
        }
    }


    /**
     * Check for cached data on startup
     */
    async checkCachedData() {
        try {
            config.debug('🔍 Checking for cached data...');
            
            // Check if IndexedDB is available
            if (!window.indexedDB) {
                config.error('❌ IndexedDB not supported in this browser');
                this.elements.cacheInfo.textContent = 'IndexedDB not supported';
                return;
            }
            
            const dbInfo = await equateDB.getDatabaseInfo();
            config.debug('📊 Database info:', dbInfo);
            
            if (dbInfo && dbInfo.hasPortfolioData) {
                config.debug('✅ Found cached portfolio data');
                const lastUpload = new Date(dbInfo.lastUpload);
                const daysAgo = Math.floor((Date.now() - lastUpload.getTime()) / (1000 * 60 * 60 * 24));
                config.debug(`📊 Cache status: Last updated ${daysAgo === 0 ? 'today' : daysAgo + ' days ago'}`);
                
                // Update drop zones to show cached data status
                this.updateDropZonesForCachedData(dbInfo);
                
                // Smart tab management - collapse tabs and show results directly
                manageTabVisibility(true);
                
                // Also check historical prices
                const historicalCount = await equateDB.getHistoricalPricesCount();
                config.debug(`📈 Historical prices in cache: ${historicalCount}`);
                
                // AUTO-LOAD cached data immediately
                config.debug('🚀 Auto-loading cached data...');
                await this.loadCachedData();
            } else {
                config.debug('ℹ️ No cached data found');
                
                // Clear any cached data indicators from drop zones
                this.clearDropZoneIndicators();
                
                // Smart tab management - show tabs in normal state
                manageTabVisibility(false);
            }
        } catch (error) {
            config.error('❌ Error checking cached data:', error);
            // Show tabs in normal state if error occurs
            manageTabVisibility(false);
        }
    }

    /**
     * Load historical prices using environment-aware approach (like Tourploeg)
     * Only loads hist.xlsx for Allianz company, skips for others
     */
    async loadHistoricalPrices(company = 'Other') {
        config.debug(`🔍 loadHistoricalPrices called with company: "${company}"`);
        try {
            // Skip loading hist.xlsx for non-Allianz companies
            if (company !== 'Allianz') {
                config.debug(`🏢 Company is "${company}" (not "Allianz"), skipping hist.xlsx loading`);
                return;
            }

            // Check if we already have historical prices
            const existingPrices = await equateDB.getHistoricalPrices();
            if (existingPrices.length > 0) {
                config.debug(`Historical prices loaded from cache: ${existingPrices.length} entries`);
                return;
            }

            config.debug('🏢 Company is Allianz, loading historical prices from hist.xlsx...');

            // Environment-aware URL construction (like Tourploeg project)
            let baseUrl = '';
            const currentUrl = window.location.href;
            
            config.debug('Current URL:', currentUrl);
            
            if (currentUrl.includes('github.io')) {
                // Production: GitHub Pages - use raw GitHub URL
                const parts = currentUrl.split('.github.io')[0].split('//')[1];
                const repoName = currentUrl.split('/')[3] || 'Equate';
                baseUrl = `https://raw.githubusercontent.com/${parts}/${repoName}/main/`;
                config.debug('Production mode detected, base URL:', baseUrl);
            } else {
                // Development: Local - use relative path
                baseUrl = './';
                config.debug('Development mode detected, base URL:', baseUrl);
            
                // Check if we're using file:// protocol
                if (window.location.protocol === 'file:') {
                    config.warn('🚨 CORS Issue Detected: You are running via file:// protocol');
                    config.warn('📋 To test historical prices locally, please run:');
                    config.warn('   1. Open terminal/command prompt in project folder');
                    config.warn('   2. Run: python -m http.server 8000');
                    config.warn('   3. Open: http://localhost:8000');
                    config.warn('💡 This is normal - GitHub Pages deployment will work fine');
                }
            }

            // Try to load historical prices with cache buster
            const cacheBuster = `?v=${Date.now()}`;
            const fullUrl = baseUrl + 'hist.xlsx' + cacheBuster;
            config.debug('Attempting to fetch:', fullUrl);
            
            const response = await fetch(fullUrl);
            
            config.debug('Fetch response:', {
                status: response.status,
                statusText: response.statusText,
                ok: response.ok,
                url: response.url
            });
            
            if (!response.ok) {
                config.warn(`Historical price file fetch failed: ${response.status} ${response.statusText} - calculations will use portfolio market price`);
                return;
            }

            const arrayBuffer = await response.arrayBuffer();
            const data = XLSX.read(arrayBuffer, { type: 'array' });
            const worksheet = data.Sheets[data.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Parse historical data
            const priceEntries = [];
            for (let i = 1; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row[0] || !row[1]) continue;

                // Convert Excel date number to JavaScript date
                const excelDate = row[0];
                const price = parseFloat(row[1]);

                if (typeof excelDate === 'number' && !isNaN(price)) {
                    const jsDate = this.excelDateToJSDate(excelDate);
                    priceEntries.push({
                        date: jsDate.toISOString().split('T')[0],
                        price: price
                    });
                } else {
                    config.warn(`Skipping invalid data row ${i}:`, { excelDate, price, row });
                }
            }

            if (priceEntries.length === 0) {
                config.warn('No valid price entries found in hist.xlsx');
                return;
            }

            // Sort by date and save to database
            priceEntries.sort((a, b) => new Date(a.date) - new Date(b.date));
            
            // Add asOfDate from portfolio if available and more recent
            if (portfolioCalculator.portfolioData) {
                const asOfDate = portfolioCalculator.extractAsOfDate();
                const marketPrice = portfolioCalculator.extractMarketPrice();
                
                if (asOfDate && marketPrice) {
                    const mostRecentHistorical = priceEntries.length > 0 
                        ? priceEntries[priceEntries.length - 1].date 
                        : null;
                    
                    if (!mostRecentHistorical || asOfDate > mostRecentHistorical) {
                        priceEntries.push({
                            date: asOfDate,
                            price: marketPrice
                        });
                        config.debug(`📈 Added asOfDate ${asOfDate} (€${marketPrice}) to historical prices before saving to cache`);
                        
                        // Re-sort after adding asOfDate
                        priceEntries.sort((a, b) => new Date(a.date) - new Date(b.date));
                    }
                }
            }
            
            config.debug(`About to save ${priceEntries.length} price entries to database`);
            config.debug('Sample entries before saving:', priceEntries.slice(0, 3));
            
            await equateDB.saveHistoricalPrices(priceEntries);
            
            // Verify they were saved
            const savedPrices = await equateDB.getHistoricalPrices();
            config.debug(`✅ Historical prices loaded and cached: ${priceEntries.length} entries`);
            config.debug(`✅ Verified in database: ${savedPrices.length} entries`);
            config.debug(`Date range: ${priceEntries[0]?.date} to ${priceEntries[priceEntries.length - 1]?.date}`);
            
        } catch (error) {
            config.warn('Could not load historical prices:', error.message);
            config.debug('Calculations will use portfolio market price as fallback');
        }
    }

    /**
     * Convert Excel date serial number to JavaScript Date
     */
    excelDateToJSDate(excelDate) {
        // Excel's epoch is 1900-01-01, but it incorrectly treats 1900 as a leap year
        const excelEpoch = new Date(1899, 11, 30); // December 30, 1899
        const jsDate = new Date(excelEpoch.getTime() + (excelDate * 24 * 60 * 60 * 1000));
        return jsDate;
    }

    /**
     * Load user preferences
     */
    async loadUserPreferences() {
        try {
            const theme = await equateDB.getPreference('theme') || 'light';
            this.elements.themeSelect.value = theme;
            this.switchTheme(theme);

            const numberFormat = await equateDB.getPreference('numberFormat') || 'us';
            this.elements.formatSelect.value = numberFormat;
            this.setNumberFormat(numberFormat);

            const language = await equateDB.getPreference('language') || 'english';
            this.elements.languageSelect.value = language;
            translationManager.setLanguage(language);
            config.debug('🌐 Restored language preference:', language);
        } catch (error) {
            config.error('Error loading preferences:', error);
        }
    }

    /**
     * Set language and save preference
     */
    setLanguagePreference(language) {
        try {
            this.elements.languageSelect.value = language;
            translationManager.setLanguage(language);
            equateDB.savePreference('language', language);
            config.debug('🌐 Language preference saved:', language);
        } catch (error) {
            config.error('Error setting language preference:', error);
        }
    }

    /**
     * Switch theme
     */
    switchTheme(theme) {
        document.body.setAttribute('data-theme', theme);
        equateDB.savePreference('theme', theme);
    }

    /**
     * Set number format and refresh all displays
     */
    setNumberFormat(format) {
        try {
            console.log('App.setNumberFormat called with:', format);
            this.numberFormat = format;
            equateDB.savePreference('numberFormat', format);
            
            // Refresh all displays if we have calculation data
            if (this.currentCalculations) {
                console.log('Refreshing displays with new format:', format);
                this.updateMetricsDisplay();
                this.refreshCharts();
                console.log('Format updated successfully to:', format);
            } else {
                console.log('No calculations to refresh yet - format will apply when data is loaded');
            }
        } catch (error) {
            console.error('Error in App.setNumberFormat:', error);
        }
    }


    /**
     * Refresh all charts with current number format
     */
    refreshCharts() {
        // Only refresh charts if we have valid calculation data
        if (!this.currentCalculations || !this.currentCalculations.userInvestment) {
            console.log('No valid calculations available - skipping chart refresh');
            console.log('currentCalculations:', this.currentCalculations);
            return;
        }
        
        // Get the same parameters as the original chart creation
        const calc = this.currentCalculations;
        const currencySymbol = this.getCurrencySymbol(calc.currency);
        const fontColor = getComputedStyle(document.documentElement).getPropertyValue('--text-primary').trim();
        const bgColor = getComputedStyle(document.documentElement).getPropertyValue('--bg-secondary').trim();
        
        // Recreate all charts to apply new number formatting
        this.createPortfolioChart();
        this.createPerformanceBarChart(calc, currencySymbol, fontColor, bgColor);
        this.createInvestmentPieChart(calc, currencySymbol, fontColor, bgColor);
    }



    /**
     * Core number formatting utility - handles both US and EU formats
     * @param {number} value - Number to format
     * @param {number} decimals - Number of decimal places
     * @returns {string} Formatted number string
     */
    formatNumber(value, decimals = 2) {
        if (value == null || isNaN(value)) return '-';
        
        // Get user's preferred format (default to 'us' if not set)
        const format = this.numberFormat || 'us';
        
        if (format === 'eu') {
            // European format: 10.000,00
            return value.toLocaleString('de-DE', {
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals
            });
        } else {
            // US format: 10,000.00
            return value.toLocaleString('en-US', {
                minimumFractionDigits: decimals,
                maximumFractionDigits: decimals
            });
        }
    }

    /**
     * Format currency
     */
    formatCurrency(value, currency = null, decimals = 2) {
        // Use getCurrencySymbol which handles unknown currencies properly
        const symbol = this.getCurrencySymbol(currency);
        
        const formattedNumber = this.formatNumber(value, decimals);
        
        // Only add symbol if we found a valid one, otherwise just show the number
        return symbol ? `${symbol} ${formattedNumber}` : formattedNumber;
    }

    /**
     * Format percentage
     */
    formatPercentage(value, decimals = 2) {
        if (value == null || isNaN(value)) return '-';
        
        const formattedNumber = this.formatNumber(value, decimals);
        return `${formattedNumber}%`;
    }

    /**
     * Show loading overlay
     */
    showLoading(message = 'Loading...') {
        if (this.elements.loadingOverlay) {
            this.elements.loadingOverlay.style.display = 'flex';
            const text = this.elements.loadingOverlay.querySelector('p');
            if (text) text.textContent = message;
        }
    }

    /**
     * Hide loading overlay
     */
    hideLoading() {
        if (this.elements.loadingOverlay) {
            this.elements.loadingOverlay.style.display = 'none';
        }
    }

    /**
     * Show error message
     */
    showError(message) {
        alert('Error: ' + message); // Simple error handling - could be improved with a modal
    }


    /**
     * Create interactive portfolio timeline chart
     */
    async createPortfolioChart() {
        if (!this.currentCalculations || !portfolioCalculator.portfolioData) {
            return;
        }

        try {
            // Get detected currency - NO FALLBACKS, if we don't have it, we don't use one
            let currency = this.detectedCurrency;
            
            // Only use calculations currency if it's not 'UNKNOWN'
            if (!currency && this.currentCalculations && this.currentCalculations.currency !== 'UNKNOWN') {
                currency = this.currentCalculations.currency;
            }
            
            config.debug('🔍 Chart currency resolution:', {
                detectedCurrency: this.detectedCurrency,
                calculationsOriginCurrency: this.currentCalculations?.currency,
                finalCurrency: currency
            });
            
            const currencySymbol = currency ? this.getCurrencySymbol(currency) : '';
            
            // Get timeline data from calculator
            const timelineData = await portfolioCalculator.getPortfolioTimeline();
            
            if (!timelineData || timelineData.length === 0) {
                // Show a placeholder message instead of empty chart
                document.getElementById('portfolioChart').innerHTML = 
                    '<div style="display: flex; align-items: center; justify-content: center; height: 100%; color: #666; flex-direction: column;">' +
                    '<div style="font-size: 18px; margin-bottom: 10px;">📊</div>' +
                    '<div>No timeline data available for chart</div>' +
                    '<div style="font-size: 12px; margin-top: 5px;">This may be due to missing allocation dates or historical price data</div>' +
                    '</div>';
                return;
            }

            // Get historical prices
            const historicalPrices = await equateDB.getHistoricalPrices();
            
            // Use daily data by default - Plotly controls handle filtering
            const filteredData = timelineData;

            // Prepare data for Plotly
            const dates = filteredData.map(d => d.date);
            const portfolioValues = filteredData.map(d => d.portfolioValue);
            const profitLoss = filteredData.map(d => d.profitLoss);
            const stockPrices = filteredData.map(d => d.currentPrice);
            
            // Find transaction points with reasons
            const transactionPoints = filteredData.filter(d => d.hasTransaction && d.reason);
            const transactionDates = transactionPoints.map(d => d.date);
            const transactionValues = transactionPoints.map(d => d.portfolioValue);
            const transactionReasons = transactionPoints.map(d => d.reason);
            
            // Find manual price points - ONLY the specific manual price date, not all matching prices
            const manualPricePoints = [];
            if (this.currentCalculations.priceSource === 'manual') {
                config.debug('🎯 Looking for manual price points. Manual price:', this.currentCalculations.currentPrice);
                
                // Find the latest point(s) that match the manual price - these are the scenario points
                const latestPoint = filteredData[filteredData.length - 1];
                
                if (latestPoint && Math.abs(latestPoint.currentPrice - this.currentCalculations.currentPrice) < 0.01) {
                    manualPricePoints.push(latestPoint);
                    config.debug('🎯 Manual price point found at date:', latestPoint.date, 'price:', latestPoint.currentPrice);
                } else {
                    config.debug('🎯 No manual price point - matches existing reality, no scenario needed');
                }
            }

            // Create traces - using single Y-axis for both Portfolio Value and Profit/Loss
            const traces = [
                {
                    x: dates,
                    y: portfolioValues,
                    type: 'scatter',
                    mode: 'lines',
                    name: translationManager.t('portfolio_value_legend'),
                    line: { color: '#3498db', width: 3 },
                    hovertemplate: `<b>${translationManager.t('portfolio_value_legend')}</b><br>` +
                                 `${translationManager.t('chart_date_label')}: %{x|%d-%m-%Y}<br>` +
                                 `${translationManager.t('chart_value_label')}: ${currencySymbol}%{y:,.2f}<br>` +
                                 '<extra></extra>'
                },
                {
                    x: dates,
                    y: profitLoss,
                    type: 'scatter',
                    mode: 'lines',
                    name: translationManager.t('profit_loss_legend'),
                    line: { color: '#27ae60', width: 2, dash: 'dot' },
                    hovertemplate: `<b>${translationManager.t('profit_loss_legend')}</b><br>` +
                                 `${translationManager.t('chart_date_label')}: %{x|%d-%m-%Y}<br>` +
                                 `${translationManager.t('chart_value_label')}: ${currencySymbol}%{y:,.2f}<br>` +
                                 '<extra></extra>'
                },
                {
                    x: dates,
                    y: stockPrices,
                    type: 'scatter',
                    mode: 'lines',
                    name: translationManager.t('stock_price_legend'),
                    line: { color: '#95a5a6', width: 1 },
                    yaxis: 'y2',
                    hovertemplate: `<b>${translationManager.t('stock_price_legend')}</b><br>` +
                                 `${translationManager.t('chart_date_label')}: %{x|%d-%m-%Y}<br>` +
                                 `${translationManager.t('chart_share_price_label')}: ${currencySymbol}%{y:,.2f}<br>` +
                                 '<extra></extra>'
                }
            ];

            // Add transaction points with detailed hover info
            if (transactionPoints.length > 0) {
                traces.push({
                    x: transactionDates,
                    y: transactionValues,
                    type: 'scatter',
                    mode: 'markers',
                    name: translationManager.t('transaction_legend'),
                    marker: {
                        color: '#e74c3c',
                        size: 12,
                        symbol: 'diamond',
                        line: { color: 'white', width: 2 }
                    },
                    text: transactionReasons,
                    hovertemplate: `<b>${translationManager.t('transaction_legend')}</b><br>` +
                                 `${translationManager.t('chart_date_label')}: %{x|%d-%m-%Y}<br>` +
                                 `${translationManager.t('chart_value_label')}: ${currencySymbol}%{y:,.0f}<br>` +
                                 '%{text}<br>' +
                                 '<extra></extra>'
                });
            }

            // Add manual price points with enhanced hover info
            if (manualPricePoints.length > 0) {
                // Manual price marker on Portfolio Value line
                traces.push({
                    x: manualPricePoints.map(d => d.date),
                    y: manualPricePoints.map(d => d.portfolioValue),
                    type: 'scatter',
                    mode: 'markers',
                    name: translationManager.t('manual_price_scenario_legend'),
                    marker: {
                        color: '#f39c12',
                        size: 15,
                        symbol: 'star',
                        line: { color: 'white', width: 2 }
                    },
                    hovertemplate: `<b>${translationManager.t('manual_price_scenario_legend')}</b><br>` +
                                 `${translationManager.t('chart_date_label')}: %{x|%d-%m-%Y}<br>` +
                                 `${translationManager.t('portfolio_value_legend')}: ${currencySymbol}%{y:,.2f}<br>` +
                                 `${translationManager.t('chart_share_price_label')}: ${currencySymbol}${this.currentCalculations.currentPrice.toFixed(2)}<br>` +
                                 '<extra></extra>'
                });

                // Manual price marker on Profit/Loss line
                traces.push({
                    x: manualPricePoints.map(d => d.date),
                    y: manualPricePoints.map(d => d.profitLoss),
                    type: 'scatter',
                    mode: 'markers',
                    name: `${translationManager.t('manual_price_scenario_legend')} (${translationManager.t('profit_loss_legend')})`,
                    marker: {
                        color: '#f39c12',
                        size: 15,
                        symbol: 'star',
                        line: { color: 'white', width: 2 }
                    },
                    hovertemplate: `<b>${translationManager.t('manual_price_scenario_legend')} ${translationManager.t('profit_loss_legend')}</b><br>` +
                                 `${translationManager.t('chart_date_label')}: %{x|%d-%m-%Y}<br>` +
                                 `${translationManager.t('profit_loss_legend')}: ${currencySymbol}%{y:,.2f}<br>` +
                                 `${translationManager.t('chart_share_price_label')}: ${currencySymbol}${this.currentCalculations.currentPrice.toFixed(2)}<br>` +
                                 '<extra></extra>'
                });
            }

            // Layout configuration - dual Y-axis (left: portfolio values, right: stock price)
            const layout = {
                separators: this.numberFormat === 'eu' ? ',.' : '.,',
                xaxis: {
                    title: translationManager.t('chart_date_axis_label'),
                    type: 'date',
                    tickangle: -45,
                    rangeselector: {
                        buttons: [
                            {count: 1, label: '1M', step: 'month', stepmode: 'backward'},
                            {count: 3, label: '3M', step: 'month', stepmode: 'backward'},
                            {count: 6, label: '6M', step: 'month', stepmode: 'backward'},
                            {count: 1, label: '1Y', step: 'year', stepmode: 'backward'},
                            {step: 'all', label: 'All'}
                        ],
                        y: 1.02,
                        yanchor: 'bottom'
                    }
                },
                yaxis: {
                    title: `${translationManager.t('chart_value_axis_label')} (${currencySymbol})`,
                    titlefont: { color: '#3498db' },
                    tickfont: { color: '#3498db' },
                    tickformat: ',.0f',
                    zeroline: true,
                    zerolinecolor: '#666',
                    zerolinewidth: 1,
                    fixedrange: false,
                    scaleanchor: null,
                    scaleratio: null,
                    side: 'left'
                },
                yaxis2: {
                    title: `${translationManager.t('stock_price_legend')} (${currencySymbol})`,
                    titlefont: { color: '#95a5a6' },
                    tickfont: { color: '#95a5a6' },
                    tickformat: ',.2f',
                    overlaying: 'y',
                    side: 'right',
                    fixedrange: false,
                    showgrid: false
                },
                legend: {
                    x: 0,
                    y: 1,
                    bgcolor: 'rgba(255,255,255,0.8)'
                },
                hovermode: 'closest',
                dragmode: 'zoom',
                margin: { l: 80, r: 120, t: 50, b: 80 },
                plot_bgcolor: 'rgba(0,0,0,0)',
                paper_bgcolor: 'rgba(0,0,0,0)',
                autosize: true,
                height: 450
            };

            // Apply theme
            const isDark = document.body.getAttribute('data-theme') === 'dark';
            if (isDark) {
                layout.font = { color: '#e2e8f0' };
                layout.plot_bgcolor = 'rgba(30, 41, 59, 0.5)';
                layout.paper_bgcolor = 'rgba(30, 41, 59, 0.5)';
            }

            // Configuration
            const plotlyConfig = {
                responsive: true,
                displayModeBar: true,
                modeBarButtonsToRemove: ['pan2d', 'lasso2d', 'select2d'],
                displaylogo: false,
                scrollZoom: true
            };

            // Create the chart - let it fill the container naturally
            Plotly.newPlot('portfolioChart', traces, layout, plotlyConfig);

            // Store timeline data and currency for language change updates
            this.lastTimelineData = timelineData;
            this.lastCurrencyUsed = currency;
            
            // Show disclaimer if total sold is 0 but portfolio value differs from timeline end
            this.updateTimelineDisclaimer(timelineData, currencySymbol);
            
            // Show legend interaction hint overlay for 4 seconds
            this.showLegendInteractionHint();

        } catch (error) {
            config.error('Error creating portfolio chart:', error);
            document.getElementById('portfolioChart').innerHTML = 
                '<div style="display: flex; align-items: center; justify-content: center; height: 100%; color: #666;">' +
                'Error loading chart: ' + error.message + '</div>';
        }
    }

    /**
     * Show legend interaction hint overlay for 4 seconds
     */
    showLegendInteractionHint() {
        // Check if we should show the hint (only for first time or when stock price line is newly added)
        const chartContainer = document.getElementById('portfolioChart');
        if (!chartContainer) return;

        // Create overlay element
        const overlay = document.createElement('div');
        overlay.id = 'legendHintOverlay';
        overlay.innerHTML = translationManager.t('legend_interaction_hint');
        
        // Style the overlay
        overlay.style.cssText = `
            position: absolute;
            left: 50%;
            top: 180px;
            transform: translateX(-50%);
            background: rgba(120, 120, 120, 0.9);
            color: white;
            padding: 12px 20px;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            white-space: nowrap;
            z-index: 1000;
            opacity: 0;
            transition: opacity 0.3s ease-in-out;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
        `;
        
        // Position relative to chart container
        const chartWrapper = chartContainer.closest('.chart-container');
        if (chartWrapper) {
            chartWrapper.style.position = 'relative';
            chartWrapper.appendChild(overlay);
        } else {
            chartContainer.style.position = 'relative';
            chartContainer.appendChild(overlay);
        }
        
        // Fade in
        setTimeout(() => {
            overlay.style.opacity = '1';
        }, 100);
        
        // Fade out and remove after 4 seconds
        setTimeout(() => {
            overlay.style.opacity = '0';
            setTimeout(() => {
                if (overlay.parentNode) {
                    overlay.parentNode.removeChild(overlay);
                }
            }, 300);
        }, 4000);
    }

    /**
     * Refresh timeline disclaimer with current language (called when language changes)
     */
    refreshTimelineDisclaimer() {
        config.debug('🌐 refreshTimelineDisclaimer called', {
            hasCalculations: !!this.currentCalculations,
            hasTimelineData: !!this.lastTimelineData,
            detectedCurrency: this.detectedCurrency,
            calculationsCurrency: this.currentCalculations?.currency,
            totalSold: this.currentCalculations?.totalSold
        });
        
        if (this.currentCalculations && this.lastTimelineData) {
            // NO FALLBACKS - if we don't have currency, we don't use one
            let currency = this.detectedCurrency;
            
            // Only use calculations currency if it's not 'UNKNOWN'
            if (!currency && this.currentCalculations && this.currentCalculations.currency !== 'UNKNOWN') {
                currency = this.currentCalculations.currency;
            }
            
            // Try the stored currency from timeline creation (but still no fallbacks like EUR)
            if (!currency || currency === 'UNKNOWN') {
                currency = this.lastCurrencyUsed;
            }
            
            config.debug('🔍 Currency sources for disclaimer refresh:', {
                detectedCurrency: this.detectedCurrency,
                calculationsCurrency: this.currentCalculations.currency,
                lastCurrencyUsed: this.lastCurrencyUsed,
                finalCurrency: currency
            });
            
            // If no valid currency found, use empty string (no symbol)
            const currencySymbol = (currency && currency !== 'UNKNOWN') ? this.getCurrencySymbol(currency) : '';
            this.updateTimelineDisclaimer(this.lastTimelineData, currencySymbol);
            config.debug('🌐 Refreshed timeline disclaimer for language change with currency:', currency);
        } else {
            config.debug('⚠️ Missing calculations or timeline data for disclaimer refresh');
        }
    }

    /**
     * Update timeline disclaimer AND total sold card disclaimer based on totalSold and value discrepancies
     */
    updateTimelineDisclaimer(timelineData, currencySymbol) {
        const timelineDisclaimerElement = document.getElementById('timelineDisclaimer');
        const totalSoldDisclaimerElement = document.getElementById('totalSoldDisclaimer');
        
        if (!timelineDisclaimerElement || !timelineData || timelineData.length === 0) {
            return;
        }

        const totalSold = this.currentCalculations.totalSold;
        
        // Show disclaimer only if totalSold = 0 (no transaction data)
        if (totalSold === 0) {
            const timelineEndValue = timelineData[timelineData.length - 1].portfolioValue;
            const currentPortfolioValue = this.currentCalculations.currentValue;
            const estimatedSales = timelineEndValue - currentPortfolioValue;
            
            if (estimatedSales > 0.01) { // Show only if significant difference (> 1 cent)
                // Format amount with currency symbol and proper locale
                const formattedAmount = `${currencySymbol}${estimatedSales.toLocaleString('en-US', { 
                    minimumFractionDigits: 2, 
                    maximumFractionDigits: 2 
                })}`;
                
                // Build complete disclaimer text using translations for timeline
                const timelineDisclaimerText = `${translationManager.t('timeline_disclaimer_text')} ${formattedAmount} ${translationManager.t('timeline_disclaimer_suffix')}`;
                
                // Update the timeline disclaimer content
                timelineDisclaimerElement.querySelector('.disclaimer-content').innerHTML = 
                    `<strong>⚠️ </strong>${timelineDisclaimerText}`;
                
                timelineDisclaimerElement.style.display = 'block';
                
                // Update the Total Sold card disclaimer (small and subtle)
                if (totalSoldDisclaimerElement) {
                    const cardDisclaimerText = `(${formattedAmount} ${translationManager.t('card_disclaimer_suffix')})`;
                    totalSoldDisclaimerElement.textContent = cardDisclaimerText;
                    totalSoldDisclaimerElement.style.display = 'block';
                    config.debug('📢 Total Sold card disclaimer updated:', cardDisclaimerText);
                } else {
                    config.debug('⚠️ totalSoldDisclaimerElement not found!');
                }
                
                config.debug('📢 Both disclaimers shown:', {
                    timelineEndValue,
                    currentPortfolioValue,
                    estimatedSales: formattedAmount,
                    timelineText: timelineDisclaimerText,
                    cardText: totalSoldDisclaimerElement?.textContent
                });
                return;
            }
        }
        
        // Hide both disclaimers in all other cases
        timelineDisclaimerElement.style.display = 'none';
        if (totalSoldDisclaimerElement) {
            totalSoldDisclaimerElement.style.display = 'none';
        }
        config.debug('📢 Both disclaimers hidden (totalSold > 0 or no discrepancy)');
    }

    /**
     * Create interactive breakdown charts (pie chart and performance bar)
     */
    createBreakdownCharts() {
        if (!this.currentCalculations) {
            return;
        }

        try {
            const calc = this.currentCalculations;
            const currency = this.detectedCurrency;
            const currencySymbol = this.getCurrencySymbol(currency);
            
            // Apply theme
            const isDark = document.body.getAttribute('data-theme') === 'dark';
            const fontColor = isDark ? '#e2e8f0' : '#2c3e50';
            const bgColor = isDark ? 'rgba(30, 41, 59, 0.5)' : 'rgba(255, 255, 255, 0.5)';

            // 1. Investment Sources Pie Chart
            this.createInvestmentPieChart(calc, currencySymbol, fontColor, bgColor);
            
            // 2. Performance Bar Chart
            this.createPerformanceBarChart(calc, currencySymbol, fontColor, bgColor);
            
            // 3. Apply section reordering after charts are created
            setTimeout(() => {
                this.applySectionOrderAfterChartsCreated();
            }, 100);
            
        } catch (error) {
            config.error('Error creating breakdown charts:', error);
        }
    }

    /**
     * Create pie chart showing investment sources
     */
    createInvestmentPieChart(calc, currencySymbol, fontColor, bgColor) {
        const pieData = [{
            values: [calc.userInvestment, calc.companyMatch, calc.freeShares, calc.dividendIncome],
            labels: [translationManager.t('your_investment_legend'), translationManager.t('company_match_legend'), translationManager.t('free_shares_legend'), translationManager.t('dividend_income_legend')],
            type: 'pie',
            hole: 0.4, // Donut chart
            hovertemplate: '<b>%{label}</b><br>' +
                           'Amount: ' + currencySymbol + '%{value:,.2f}<br>' +
                           'Percentage: %{percent}<br>' +
                           '<extra></extra>',
            textinfo: 'label+percent+value',
            textposition: 'auto',
            texttemplate: '%{label}<br>' + currencySymbol + '%{value:,.0f}<br>(%{percent})',
            marker: {
                colors: ['#27ae60', '#3498db', '#9b59b6', '#f39c12'], // Green, Blue, Purple, Orange
                line: {
                    color: fontColor,
                    width: 1
                }
            }
        }];

        const pieLayout = {
            separators: this.numberFormat === 'eu' ? ',.' : '.,',
            font: { color: fontColor, size: 12 },
            paper_bgcolor: bgColor,
            plot_bgcolor: bgColor,
            showlegend: true,
            legend: {
                orientation: 'h',
                y: -0.1,
                x: 0.5,
                xanchor: 'center',
                font: { size: 12 }
            },
            margin: { t: 20, b: 60, l: 20, r: 20 },
            annotations: [{
                font: { size: 18, color: fontColor },
                showarrow: false,
                text: `Total<br>${currencySymbol}${calc.totalInvestment.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`,
                x: 0.5,
                y: 0.5
            }]
        };

        const pieConfig = {
            responsive: true,
            displayModeBar: true,
            modeBarButtonsToRemove: ['pan2d', 'lasso2d', 'select2d', 'autoScale2d'],
            displaylogo: false,
            toImageButtonOptions: {
                format: 'png',
                filename: 'investment_sources',
                height: 600,
                width: 800,
                scale: 2
            }
        };

        Plotly.newPlot('investmentPieChart', pieData, pieLayout, pieConfig);

        // Add click handler for interactive summary
        const pieChartElement = document.getElementById('investmentPieChart');
        pieChartElement.removeAllListeners('plotly_click'); // Remove existing listeners
        pieChartElement.on('plotly_click', (data) => {
            this.updatePieChartSummary(data.points[0], calc, currencySymbol);
        });
    }

    /**
     * Create stacked bar chart showing performance overview
     */
    createPerformanceBarChart(calc, currencySymbol, fontColor, bgColor) {
        const barData = [
            // Your Investment (base layer)
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [calc.userInvestment, 0, calc.userInvestment, 0],
                name: translationManager.t('your_investment_legend'),
                type: 'bar',
                marker: { color: '#27ae60' },
                text: [currencySymbol + calc.userInvestment.toLocaleString('en-US', { maximumFractionDigits: 0 }), '', currencySymbol + calc.userInvestment.toLocaleString('en-US', { maximumFractionDigits: 0 }), ''],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: `<b>${translationManager.t('your_investment_legend')}</b><br>` +
                               'Amount: ' + currencySymbol + '%{y:,.2f}<br>' +
                               '<extra></extra>'
            },
            // Company Match
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [calc.companyMatch, 0, 0, 0],
                name: translationManager.t('company_match_legend'),
                type: 'bar',
                marker: { color: '#3498db' },
                text: [calc.companyMatch > 2000 ? currencySymbol + calc.companyMatch.toLocaleString('en-US', { maximumFractionDigits: 0 }) : '', '', '', ''],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: `<b>${translationManager.t('company_match_legend')}</b><br>` +
                               'Amount: ' + currencySymbol + '%{y:,.2f}<br>' +
                               '<extra></extra>'
            },
            // Free Shares
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [calc.freeShares, 0, 0, 0],
                name: translationManager.t('free_shares_legend'),
                type: 'bar',
                marker: { color: '#9b59b6' },
                text: [calc.freeShares > 1500 ? currencySymbol + calc.freeShares.toLocaleString('en-US', { maximumFractionDigits: 0 }) : '', '', '', ''],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: '<b>Free Shares</b><br>' +
                               'Amount: ' + currencySymbol + '%{y:,.2f}<br>' +
                               '<extra></extra>'
            },
            // Dividends
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [calc.dividendIncome, 0, 0, 0],
                name: translationManager.t('dividend_income_legend'),
                type: 'bar',
                marker: { color: '#f39c12' },
                text: [calc.dividendIncome > 2000 ? currencySymbol + calc.dividendIncome.toLocaleString('en-US', { maximumFractionDigits: 0 }) : '', '', '', ''],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: '<b>Dividend Income</b><br>' +
                               'Amount: ' + currencySymbol + '%{y:,.2f}<br>' +
                               '<extra></extra>'
            },
            // Total Investment (base for Total Return bar)
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [0, 0, 0, calc.totalInvestment],
                name: translationManager.t('total_investment_base_legend'),
                type: 'bar',
                marker: { color: '#bdc3c7' },
                text: ['', '', '', currencySymbol + calc.totalInvestment.toLocaleString('en-US', { maximumFractionDigits: 0 })],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: '<b>Total Investment</b><br>' +
                               'Amount: ' + currencySymbol + '%{y:,.2f}<br>' +
                               'Your Investment + Company Benefits<br>' +
                               '<extra></extra>'
            },
            // Current Portfolio
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [0, calc.currentValue, 0, 0],
                name: translationManager.t('current_portfolio_legend'),
                type: 'bar',
                marker: { color: '#2ecc71' },
                text: ['', currencySymbol + calc.currentValue.toLocaleString('en-US', { maximumFractionDigits: 0 }), '', ''],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: `<b>${translationManager.t('current_portfolio_legend')}</b><br>` +
                               'Amount: ' + currencySymbol + '%{y:,.2f}<br>' +
                               '<extra></extra>'
            },
            // Total Sold
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [0, calc.totalSold, 0, 0],
                name: translationManager.t('total_sold_legend'),
                type: 'bar',
                marker: { color: '#95a5a6' },
                text: ['', calc.totalSold > 2000 ? currencySymbol + calc.totalSold.toLocaleString('en-US', { maximumFractionDigits: 0 }) : '', '', ''],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: `<b>${translationManager.t('total_sold_legend')}</b><br>` +
                               'Amount: ' + currencySymbol + '%{y:,.2f}<br>' +
                               '<extra></extra>'
            },
            // Your Return (profit)
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [0, 0, calc.totalReturn, 0],
                name: translationManager.t('return_on_your_investment_legend'),
                type: 'bar',
                marker: { color: '#e74c3c' },
                text: ['', '', currencySymbol + calc.totalReturn.toLocaleString('en-US', { maximumFractionDigits: 0 }) + ' (' + calc.returnPercentage.toFixed(1) + '%)', ''],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: `<b>${translationManager.t('return_on_your_investment_legend')}</b><br>` +
                               'Amount: ' + currencySymbol + '%{y:,.2f} (' + calc.returnPercentage.toFixed(2) + '%)<br>' +
                               'Total Value - Your Investment<br>' +
                               '<extra></extra>'
            },
            // Total Return (share price gains)
            {
                x: ['Total Investment', 'Total Value', 'Return on Your Investment', 'Return on Total Investment'],
                y: [0, 0, 0, calc.returnOnTotalInvestment],
                name: translationManager.t('return_on_total_investment_legend'),
                type: 'bar',
                marker: { color: '#e67e22' },
                text: ['', '', '', currencySymbol + calc.returnOnTotalInvestment.toLocaleString('en-US', { maximumFractionDigits: 0 }) + ' (' + calc.returnPercentageOnTotalInvestment.toFixed(1) + '%)'],
                textposition: 'inside',
                textfont: { color: 'white', size: 12 },
                hovertemplate: '<b>Return on Total Investment</b><br>' +
                               'Amount: ' + currencySymbol + '%{y:,.2f} (' + calc.returnPercentageOnTotalInvestment.toFixed(2) + '%)<br>' +
                               'Total Value - Total Investment<br>' +
                               '<extra></extra>'
            }
        ];

        const barLayout = {
            separators: this.numberFormat === 'eu' ? ',.' : '.,',
            barmode: 'stack',
            font: { color: fontColor, size: 12 },
            paper_bgcolor: bgColor,
            plot_bgcolor: bgColor,
            xaxis: {
                title: { text: '', font: { color: fontColor } },
                tickfont: { color: fontColor, size: 11 },
                gridcolor: 'rgba(128, 128, 128, 0.2)'
            },
            yaxis: {
                title: { text: `${translationManager.t('chart_value_axis_label')} (${currencySymbol})`, font: { color: fontColor, size: 12 } },
                tickfont: { color: fontColor, size: 11 },
                tickformat: ',.0f',
                gridcolor: 'rgba(128, 128, 128, 0.2)'
            },
            margin: { t: 20, b: 100, l: 60, r: 20 },
            showlegend: true,
            legend: {
                orientation: 'h',
                y: -0.2,
                x: 0.5,
                xanchor: 'center',
                font: { size: 11 }
            }
        };

        const barConfig = {
            responsive: true,
            displayModeBar: true,
            modeBarButtonsToRemove: ['pan2d', 'lasso2d', 'select2d'],
            displaylogo: false,
            toImageButtonOptions: {
                format: 'png',
                filename: 'performance_overview',
                height: 600,
                width: 1000,
                scale: 2
            }
        };

        Plotly.newPlot('performanceBarChart', barData, barLayout, barConfig);

        // Add click handler for interactive summary
        const chartElement = document.getElementById('performanceBarChart');
        chartElement.removeAllListeners('plotly_click'); // Remove existing listeners
        chartElement.on('plotly_click', (data) => {
            this.updateBreakdownSummary(data.points[0], calc, currencySymbol);
        });
    }

    /**
     * Update breakdown summary with calculation details
     */
    updateBreakdownSummary(point, calc, currencySymbol) {
        const summaryElement = document.getElementById('barChartSummary');
        let summaryText = '';

        // Show the summary when clicked
        summaryElement.style.display = 'block';

        switch (point.x) {
            case 'Total Investment':
                if (point.data.name === translationManager.t('your_investment_legend')) {
                    summaryText = `💰 <strong>${translationManager.t('your_investment_legend')}:</strong> ${currencySymbol}${calc.userInvestment.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - Money you personally invested in shares`;
                } else if (point.data.name === translationManager.t('company_match_legend')) {
                    summaryText = `🤝 <strong>${translationManager.t('company_match_legend')}:</strong> ${currencySymbol}${calc.companyMatch.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - Company matching contributions for Employee Share Purchase Plan`;
                } else if (point.data.name === translationManager.t('free_shares_legend')) {
                    summaryText = `🎁 <strong>${translationManager.t('free_shares_legend')}:</strong> ${currencySymbol}${calc.freeShares.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - Award shares given by company for free`;
                } else if (point.data.name === translationManager.t('dividend_income_legend')) {
                    summaryText = `💎 <strong>${translationManager.t('dividend_income_legend')}:</strong> ${currencySymbol}${calc.dividendIncome.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - Dividends automatically reinvested into more shares`;
                }
                break;
            case 'Total Value':
                if (point.data.name === translationManager.t('current_portfolio_legend')) {
                    summaryText = `📊 <strong>${translationManager.t('current_portfolio_legend')}:</strong> ${currencySymbol}${calc.currentValue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - Value of shares you still own`;
                } else if (point.data.name === translationManager.t('total_sold_legend')) {
                    summaryText = `💸 <strong>${translationManager.t('total_sold_legend')}:</strong> ${currencySymbol}${calc.totalSold.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - Money received from shares you already sold`;
                }
                break;
            case 'Return on Your Investment':
                if (point.data.name === translationManager.t('your_investment_legend')) {
                    summaryText = `💰 <strong>${translationManager.t('your_investment_legend')} Base:</strong> ${currencySymbol}${calc.userInvestment.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - Your original investment (the base)`;
                } else if (point.data.name === translationManager.t('return_on_your_investment_legend')) {
                    summaryText = `🎯 <strong>${translationManager.t('return_on_your_investment_legend')}:</strong> ${currencySymbol}${calc.totalReturn.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${calc.returnPercentage.toFixed(2)}%) = Total Value (${currencySymbol}${calc.totalValue.toFixed(2)}) - Your Investment (${currencySymbol}${calc.userInvestment.toFixed(2)})`;
                }
                break;
            case 'Return on Total Investment':
                if (point.data.name === translationManager.t('total_investment_base_legend')) {
                    summaryText = `🏦 <strong>${translationManager.t('total_investment_base_legend')}:</strong> ${currencySymbol}${calc.totalInvestment.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} - Your Investment + All Company Benefits (the base)`;
                } else if (point.data.name === translationManager.t('return_on_total_investment_legend')) {
                    summaryText = `📈 <strong>${translationManager.t('return_on_total_investment_legend')}:</strong> ${currencySymbol}${calc.returnOnTotalInvestment.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${calc.returnPercentageOnTotalInvestment.toFixed(2)}%) = Total Value (${currencySymbol}${calc.totalValue.toFixed(2)}) - Total Investment (${currencySymbol}${calc.totalInvestment.toFixed(2)})`;
                }
                break;
            default:
                summaryText = '💡 <strong>Key Insight:</strong> Click on any chart element to see detailed calculations';
        }

        summaryElement.innerHTML = `<p>${summaryText}</p>`;
    }

    /**
     * Update breakdown summary for pie chart clicks
     */
    updatePieChartSummary(point, calc, currencySymbol) {
        const summaryElement = document.getElementById('pieChartSummary');
        let summaryText = '';

        // Show the summary when clicked
        summaryElement.style.display = 'block';

        switch (point.label) {
            case translationManager.t('your_investment_legend'):
                summaryText = `💰 <strong>${translationManager.t('your_investment_legend')}:</strong> ${currencySymbol}${calc.userInvestment.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${(calc.userInvestment / calc.totalInvestment * 100).toFixed(1)}%) - Money you personally invested in shares`;
                break;
            case translationManager.t('company_match_legend'):
                summaryText = `🤝 <strong>${translationManager.t('company_match_legend')}:</strong> ${currencySymbol}${calc.companyMatch.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${(calc.companyMatch / calc.totalInvestment * 100).toFixed(1)}%) - Company matching contributions for Employee Share Purchase Plan`;
                break;
            case translationManager.t('free_shares_legend'):
                summaryText = `🎁 <strong>${translationManager.t('free_shares_legend')}:</strong> ${currencySymbol}${calc.freeShares.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${(calc.freeShares / calc.totalInvestment * 100).toFixed(1)}%) - Award shares given by company for free`;
                break;
            case translationManager.t('dividend_income_legend'):
                summaryText = `💎 <strong>${translationManager.t('dividend_income_legend')}:</strong> ${currencySymbol}${calc.dividendIncome.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} (${(calc.dividendIncome / calc.totalInvestment * 100).toFixed(1)}%) - Dividends automatically reinvested into more shares`;
                break;
            default:
                summaryText = '💡 <strong>Key Insight:</strong> Click on any chart element to see detailed calculations';
        }

        summaryElement.innerHTML = `<p>${summaryText}</p>`;
    }

    /**
     * Filter timeline data based on resolution
     */
    filterTimelineData(timelineData, resolution) {
        if (resolution === 'daily') {
            return timelineData;
        }

        const filtered = [];
        let interval;
        
        if (resolution === 'weekly') {
            interval = 7;
        } else if (resolution === 'monthly') {
            interval = 30;
        } else {
            interval = 7; // default to weekly
        }
        

        for (let i = 0; i < timelineData.length; i += interval) {
            // Take the last day of each interval
            const chunk = timelineData.slice(i, i + interval);
            const lastDay = chunk[chunk.length - 1];
            
            // Always include transaction days within this chunk
            const transactionDays = chunk.filter(d => d.hasTransaction);
            
            filtered.push(lastDay);
            transactionDays.forEach(t => {
                if (t.date !== lastDay.date) {
                    filtered.push(t);
                }
            });
        }

        // ALWAYS include the very last day (most recent, potentially with asOfDate or manual price)
        const lastDay = timelineData[timelineData.length - 1];
        if (lastDay) {
            // Remove any existing entry for the same date to avoid duplicates
            const existingIndex = filtered.findIndex(f => f.date === lastDay.date);
            if (existingIndex >= 0) {
                filtered[existingIndex] = lastDay; // Replace with the last day data
            } else {
                filtered.push(lastDay); // Add if not exists
            }
        }

        // Sort by date and remove duplicates
        const uniqueFiltered = filtered
            .sort((a, b) => new Date(a.date) - new Date(b.date))
            .filter((item, index, arr) => index === 0 || item.date !== arr[index - 1].date);
            
        return uniqueFiltered;
    }

    /**
     * Clear all cached data
     */
    async clearAllData() {
        try {
            this.showLoading('Clearing cached data...');
            
            // Clear database
            await equateDB.clearAllData();
            
            // Reset application state
            this.hasData = false;
            this.currentCalculations = null;
            portfolioCalculator.clearData();
            
            // Reset Results tab to no data state
            toggleResultsDisplay(false);
            
            // Show tabs in normal state for new data entry
            manageTabVisibility(false);
            
            // Reset file inputs
            ['portfolio', 'transactions'].forEach(inputId => {
                const input = document.getElementById(inputId);
                if (input) {
                    input.value = '';
                }
            });
            
            // Reset upload zones
            ['portfolioZone', 'transactionsZone'].forEach(zoneId => {
                const zone = document.getElementById(zoneId);
                const info = document.getElementById(zoneId.replace('Zone', 'Info'));
                if (zone) zone.classList.remove('file-selected');
                if (info) {
                    info.classList.remove('show');
                    info.textContent = ''; // Clear cached data text
                }
            });
            
            // Reset analyze button
            this.checkFilesReady();
            
            // Expand sections when data is cleared
            this.expandSection('helpContent');
            this.expandSection('uploadContent');
            
            // Remove focus styling from results section
            const resultsSection = document.getElementById('resultsSection');
            if (resultsSection) {
                resultsSection.classList.remove('focused');
            }
            
            this.hideLoading();
            alert('All cached data has been cleared successfully.');
            
        } catch (error) {
            config.error('Error clearing data:', error);
            this.hideLoading();
            this.showError('Failed to clear cached data: ' + error.message);
        }
    }

    /**
     * Trigger Plotly chart autoscale functionality
     */
    triggerChartAutoscale() {
        const chartElement = document.getElementById('portfolioChart');
        if (chartElement && window.Plotly) {
            try {
                // Trigger Plotly's autoscale functionality - same as clicking the autoscale button
                Plotly.relayout('portfolioChart', {
                    'xaxis.autorange': true,
                    'yaxis.autorange': true
                });
                config.debug('✅ Auto-triggered chart autoscale');
            } catch (error) {
                config.debug('⚠️ Chart autoscale failed:', error.message);
            }
        }
    }

    /**
     * Update drop zones to show cached data status
     */
    updateDropZonesForCachedData(dbInfo) {
        // Portfolio drop zone
        const portfolioZone = document.getElementById('portfolioZone');
        const portfolioInfo = document.getElementById('portfolioInfo');
        
        if (portfolioZone && portfolioInfo) {
            if (dbInfo.hasPortfolioData) {
                portfolioZone.classList.add('file-selected');
                portfolioInfo.textContent = 'Cached Portfolio Data available';
                portfolioInfo.classList.add('show');
            } else {
                portfolioZone.classList.remove('file-selected');
                portfolioInfo.textContent = '';
                portfolioInfo.classList.remove('show');
            }
        }
        
        // Transaction drop zone - handle both cached and non-cached states
        const transactionsZone = document.getElementById('transactionsZone');
        const transactionsInfo = document.getElementById('transactionsInfo');
        
        if (transactionsZone && transactionsInfo) {
            if (dbInfo.hasTransactions) {
                transactionsZone.classList.add('file-selected');
                transactionsInfo.textContent = 'Cached Transaction Data available';
                transactionsInfo.classList.add('show');
            } else {
                transactionsZone.classList.remove('file-selected');
                transactionsInfo.textContent = '';
                transactionsInfo.classList.remove('show');
            }
        }
    }

    /**
     * Clear cached data indicators from drop zones
     */
    clearDropZoneIndicators() {
        ['portfolioZone', 'transactionsZone'].forEach(zoneId => {
            const zone = document.getElementById(zoneId);
            const info = document.getElementById(zoneId.replace('Zone', 'Info'));
            if (zone) zone.classList.remove('file-selected');
            if (info) {
                info.classList.remove('show');
                info.textContent = '';
            }
        });
    }
}

// Global functions for HTML event handlers

/**
 * Toggle section visibility (called from HTML onclick)
 */
function toggleSection(sectionId) {
    const content = document.getElementById(sectionId);
    const header = content?.parentElement.querySelector('.section-header');
    const section = content?.parentElement; // Get the parent section element
    
    if (content && header && section) {
        const isCollapsed = content.classList.contains('collapsed');
        
        if (isCollapsed) {
            content.classList.remove('collapsed');
            header.classList.remove('collapsed');
            section.classList.remove('collapsed');
        } else {
            content.classList.add('collapsed');
            header.classList.add('collapsed');
            section.classList.add('collapsed');
        }
    }
}

/**
 * Switch between tabs in the tabbed interface
 */
function switchTab(tabName) {
    // Remove active class from all tabs and panes
    document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-pane').forEach(pane => pane.classList.remove('active'));
    
    // Activate the selected tab and pane
    const tabButton = document.getElementById(tabName + 'Tab');
    const tabPane = document.getElementById(tabName + 'Pane');
    
    if (tabButton && tabPane) {
        tabButton.classList.add('active');
        tabPane.classList.add('active');
    }

    // Auto-trigger Plotly autoscale when Results tab gets focus
    if (tabName === 'results') {
        setTimeout(() => {
            const chartElement = document.getElementById('portfolioChart');
            if (chartElement && window.Plotly) {
                try {
                    // Trigger Plotly's autoscale functionality
                    Plotly.relayout('portfolioChart', {
                        'xaxis.autorange': true,
                        'yaxis.autorange': true
                    });
                    config.debug('✅ Auto-triggered chart autoscale for Results tab');
                } catch (error) {
                    config.debug('⚠️ Chart autoscale failed (chart may not be loaded yet):', error.message);
                }
            }
        }, 100); // Small delay to ensure tab switch is complete
    }
}

/**
 * Smart tab management based on data state
 */
function manageTabVisibility(hasCachedData) {
    const header = document.querySelector('.header');
    const tabContent = document.querySelector('.tab-content');
    
    if (hasCachedData) {
        // Collapse header and tabs when cached data exists - user goes straight to results
        header.classList.add('collapsed');
        tabContent.classList.add('collapsed');
        switchTab('results'); // Switch to Results tab
    } else {
        // Expand header and tabs for new users - default to Upload tab
        header.classList.remove('collapsed');
        tabContent.classList.remove('collapsed');
        switchTab('upload'); // Default to Upload tab
    }
}

/**
 * Show/hide results content based on data availability
 */
function toggleResultsDisplay(hasData) {
    const noDataMessage = document.getElementById('noDataMessage');
    const resultsContent = document.getElementById('resultsContent');
    
    if (hasData) {
        noDataMessage.style.display = 'none';
        resultsContent.style.display = 'block';
    } else {
        noDataMessage.style.display = 'block';
        resultsContent.style.display = 'none';
    }
}

function switchTheme(theme) {
    app.switchTheme(theme);
}

function setNumberFormat(format) {
    try {
        console.log('setNumberFormat called with:', format);
        if (app) {
            app.setNumberFormat(format);
        } else {
            console.error('App not initialized yet');
        }
    } catch (error) {
        console.error('Error in setNumberFormat:', error);
    }
}

function loadCachedData() {
    app.loadCachedData();
}

function generateAnalysis() {
    return app.generateAnalysis();
}

function applyManualPrice() {
    const manualPriceInput = document.getElementById('manualPrice');
    if (manualPriceInput && manualPriceInput.value) {
        app.handleManualPriceInput(manualPriceInput.value);
    }
}

function updateManualPrice() {
    const manualPriceInput = document.getElementById('manualPriceInput');
    if (manualPriceInput && manualPriceInput.value) {
        app.handleManualPriceInput(manualPriceInput.value);
    }
}

function recalculateWithManualPrice() {
    const manualPriceInput = document.getElementById('manualPriceInput');
    if (manualPriceInput && manualPriceInput.value) {
        app.handleManualPriceInput(manualPriceInput.value);
    }
}

function handleFileSelect(input, zoneId) {
    app.handleFileSelect(input, zoneId);
}

function exportCSV() {
    if (app.currentCalculations && app.hasData) {
        portfolioExporter.setData(app.currentCalculations, portfolioCalculator.portfolioData);
        portfolioExporter.exportCSV();
    }
}

function exportPDF() {
    if (app.currentCalculations && app.hasData) {
        portfolioExporter.setData(app.currentCalculations, portfolioCalculator.portfolioData);
        portfolioExporter.exportPDF(); // Use actual PDF export with chart capture
    }
}


function clearCachedData() {
    if (confirm('Are you sure you want to clear all cached portfolio data? You will need to re-upload your files.')) {
        app.clearAllData();
    }
}

// Initialize the application when DOM is loaded
const app = new EquateApp();

document.addEventListener('DOMContentLoaded', () => {
    app.init();
});