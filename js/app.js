/**
 * Main Application Logic for Equate Portfolio Analysis
 * Coordinates all components and handles user interactions
 */

class EquateApp {
    constructor() {
        this.isInitialized = false;
        this.hasData = false;
        this.currentCalculations = null;
        this.justProcessedFreshUpload = false; // Flag to prevent duplicate processing
        
        // =============================================
        // NEW SERVICE ARCHITECTURE
        // =============================================
        
        // Initialize basic services that don't require external dependencies
        this.eventBus = new EventBus();
        if (typeof config !== 'undefined' && !config.isProduction) {
            this.eventBus.setDebugMode(true);
        }
        
        this.domManager = new DOMManager();
        
        // Keep legacy elements reference for backwards compatibility during transition
        this.elements = this.domManager.elements;
        
        // Services requiring external dependencies will be initialized in init()
        this.currencyService = null;
        
        // Reference Points Builder for unified extraction
        this.referencePointsBuilder = null;
        
        // Company selection for Sprint 2
        this.selectedCompany = null;
        this.currentCompanyFile = null;
        
    }

    /**
     * Set up event listeners for service coordination
     * Handles communication between services without tight coupling
     */
    setupServiceEventListeners() {
        if (!this.eventBus) return;

        // Listen for calculation updates from CurrencyService
        this.eventBus.on(EventBus.Events.CALCULATIONS_UPDATED, (calculations) => {
            this.currentCalculations = calculations;
            this.createPortfolioChart();
            this.createBreakdownCharts();
        });

        // Listen for loading events
        this.eventBus.on(EventBus.Events.LOADING_STARTED, (message) => {
            this.showLoading(message);
        });

        this.eventBus.on(EventBus.Events.LOADING_FINISHED, () => {
            this.hideLoading();
        });

        // Listen for currency changes to update legacy properties (temporary compatibility)
        this.eventBus.on(EventBus.Events.CURRENCY_CHANGED, (currency) => {
            this.detectedCurrency = currency; // Keep legacy property in sync
        });
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

    // =============================================
    // INITIALIZATION METHODS
    // =============================================

    /**
     * Initialize the application
     */
    async init() {
        try {
            this.showLoading('Initializing application...');
            
            // Initialize translation system
            await translationManager.init();
            config.debug('✅ Translation system initialized');
            
            // Load hist_base.xlsx metadata for instrument-to-company mapping
            await this.loadBaseMetadata();
            
            // Initialize database with stability checks
            await equateDB.init();
            config.debug('✅ Database initialized');
            
            // Initialize CurrencyService now that dependencies are ready
            if (!this.currencyService) {
                try {
                    this.currencyService = new CurrencyService({
                        database: equateDB,
                        converter: currencyConverter,
                        mappings: new CurrencyMappings(),
                        translator: translationManager,
                        calculator: portfolioCalculator,
                        eventBus: this.eventBus,
                        domManager: this.domManager
                    });
                    
                    // Set up cross-service event listeners
                    this.setupServiceEventListeners();
                    
                    config.debug('✅ CurrencyService initialized with dependencies');
                } catch (error) {
                    config.error('❌ Failed to initialize CurrencyService:', error);
                    // Continue with legacy functionality
                }
            }
            
            // Run stability checks
            await this.runStabilityChecks();

            // Get DOM elements
            this.cacheElements();
            
            // Set up event listeners
            this.setupEventListeners();
            this.setupSectionReordering();
            
            // Initialize export customizer
            if (window.exportCustomizer) {
                window.exportCustomizer.init();
                config.debug('✅ Export customizer initialized');
            }
            
            // Initialize Reference Points Builder
            this.referencePointsBuilder = new ReferencePointsBuilder();
            config.debug('✅ Reference Points Builder initialized');
            
            
            // Load user preferences
            await this.loadUserPreferences();
            
            // Initialize currency selector
            await this.currencyService.initializeSelector();
            
            // TEMPORARY: Show currency selector for testing (remove after testing)
            if (!config.isProduction) {
                this.currencyService.showSelector();
                config.debug('💱 Currency selector shown for testing');
                
                // Check database version
                config.debug('🗄️ Database version:', equateDB.version);
                config.debug('🗄️ Object stores:', equateDB.db ? Array.from(equateDB.db.objectStoreNames) : 'Not available');
            }
            
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

    // =============================================
    // DOM ELEMENT CACHING
    // =============================================

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

    // =============================================
    // EVENT LISTENERS & UI SETUP
    // =============================================

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
            
            zone.addEventListener('drop', async (e) => {
                e.preventDefault();
                e.stopPropagation();
                zone.classList.remove('drag-hover');
                
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    input.files = files;
                    await this.handleFileSelect(input, zoneId);
                }
            });
            
            // Click to select file
            zone.addEventListener('click', () => {
                input.click();
            });
            
            // Handle file selection
            input.addEventListener('change', async (e) => {
                await this.handleFileSelect(e.target, zoneId);
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
            
            // Track section reorder event
            if (typeof gtag !== 'undefined') {
                gtag('event', 'section_reorder', {
                    'from_section': draggedSection,
                    'to_position': targetIndex,
                    'new_order': currentOrder.join(','),
                    'event_category': 'user_interaction'
                });
            }
            
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
            } else if (sectionName !== 'calculations') {
                // Ignore 'calculations' section (removed feature), warn about other missing sections
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

    // =============================================
    // FILE PROCESSING PIPELINE
    // =============================================

    /**
     * Handle file selection
     */
    async handleFileSelect(input, zoneId) {
        const zone = document.getElementById(zoneId);
        const infoDiv = document.getElementById(zoneId.replace('Zone', 'Info'));
        
        if (input.files.length > 0) {
            const file = input.files[0];
            
            // Track file upload event
            if (typeof gtag !== 'undefined') {
                gtag('event', 'file_upload', {
                    'file_count': 1,
                    'event_category': 'user_interaction'
                });
            }
            
            zone.classList.add('file-selected');
            infoDiv.innerHTML = `✓ ${file.name} (${(file.size / 1024 / 1024).toFixed(2)} MB)`;
            infoDiv.classList.add('show');
        } else {
            // Only remove styling if there's no cached data to display
            const hasShowClass = infoDiv.classList.contains('show');
            const hasContent = infoDiv.innerHTML.trim() !== '';
            
            if (!hasShowClass || !hasContent || !infoDiv.innerHTML.includes('Cached')) {
                zone.classList.remove('file-selected');
                infoDiv.classList.remove('show');
                infoDiv.innerHTML = '';
                
                // Ensure instruction text is visible when no files and no cached data
                const instructionP = zone.querySelector('p');
                if (instructionP) {
                    instructionP.style.display = 'block';
                }
            }
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
            // Reset flag for new upload session
            this.justProcessedFreshUpload = false;
            
            // First, perform file analysis for modal and processing (single analysis)
            let analysisResult = null;
            const portfolioFile = this.elements.portfolioInput.files[0];
            let transactionFile = this.elements.transactionsInput.files[0] || null;

            if (portfolioFile || transactionFile) {
                // Clear any stored data from previous analysis when new files are uploaded
                this.storedAnalysisResult = null;
                this.storedPortfolioFile = null;
                this.storedTransactionFile = null;
                this.storedCachedData = null;
                this.storedPortfolioData = null;
                this.storedTransactionData = null;
                
                // Get cached data for FileAnalyzer to handle all validation
                const cachedData = await equateDB.getPortfolioData();
                
                // Single FileAnalyzer call for both modal and processing
                config.debug('🔍 Starting unified file set analysis for modal and processing...');
                analysisResult = await fileAnalyzer.analyzeFileSet(portfolioFile, transactionFile, cachedData);
                
                // Store analysis result for modal to use
                this.storedAnalysisResult = analysisResult;
                this.storedPortfolioFile = portfolioFile;
                this.storedTransactionFile = transactionFile;
                this.storedCachedData = cachedData;
            } else {
                // Clear stored results if no new files
                this.storedAnalysisResult = null;
                this.storedPortfolioFile = null;
                this.storedTransactionFile = null;
                this.storedCachedData = null;
            }

            this.showLoading('Processing your portfolio data...');

            let portfolioData, transactionData, company = 'Other', isEnglish = true; // Default to Other (safer)
            
            if (portfolioFile || transactionFile) {
                // Parse uploaded files (new upload) - using pre-computed analysis result
                config.debug('🚨 TAKING NEW UPLOAD PATH - using stored analysis result');
                const originalTransactionFile = transactionFile; // Remember original state for UI clearing

                // Clear cached data when new files are uploaded to ensure fresh company detection
                if (portfolioFile) {
                    config.debug('🗑️ Clearing cached portfolio data for fresh company detection');
                    await equateDB.clearPortfolioData();
                }
                if (transactionFile) {
                    config.debug('🗑️ Clearing cached transaction data for fresh company detection');
                    await equateDB.clearTransactionData();
                }

                // Use the stored analysis result from earlier (no duplicate analysis)
                analysisResult = this.storedAnalysisResult;
                const cachedData = this.storedCachedData;
                
                // Handle FileAnalyzer results
                if (analysisResult.shouldClearCachedTransactions) {
                    config.debug('🗑️ Clearing cached transaction data due to currency mismatch');
                    await equateDB.clearTransactionData();
                    
                    // Immediately update drop zone UI to reflect cache clearing
                    const clearedDbInfo = await equateDB.getDatabaseInfo();
                    await this.updateDropZonesForCachedData(clearedDbInfo);
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
                await this.applyUISettings(analysisResult.unified);
                
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

                // Step 4: Use pre-parsed data from modal or parse if needed
                config.debug('📄 Using stored parsed data from modal or parsing if needed...');
                if (portfolioFile) {
                    // Use stored parsed data from modal if available
                    portfolioData = this.storedPortfolioData || await fileParser.parsePortfolioFile(portfolioFile, analysisResult);
                    config.debug('Portfolio data (stored or parsed):', portfolioData);
                } else {
                    // Use cached portfolio data
                    const cachedData = await equateDB.getPortfolioData();
                    portfolioData = cachedData.portfolioData;
                    config.debug('Using cached portfolio data:', portfolioData);
                }

                // Parse transaction data if compatible
                if (transactionFile && analysisResult.errors.length === 0) {
                    // Use stored parsed data from modal if available
                    transactionData = this.storedTransactionData || await fileParser.parseTransactionFile(transactionFile, analysisResult);
                    if (!this.storedTransactionData) {
                        fileParser.validateUserIds();
                    }
                    config.debug('Transaction data (stored or parsed):', transactionData);
                }

                // Step 5: Detect company from parsed data using FileParser
                if (portfolioData?.entries) {
                    company = fileParser.detectCompanyFromBothFiles(portfolioData, transactionData);
                    config.debug('🏢 Detected company from parsed data:', company);
                    
                    // Map to clean company name for file loading
                    this.selectedCompany = await this.mapInstrumentToCompany(company);
                    if (this.selectedCompany) {
                        config.debug('🏢 Mapped to company name:', this.selectedCompany);
                    }
                    
                    // Update stored analysis result with detected company for modal display
                    if (this.storedAnalysisResult) {
                        this.storedAnalysisResult.unified.company = company;
                        if (this.storedAnalysisResult.portfolio) {
                            this.storedAnalysisResult.portfolio.company = company;
                        }
                        // Ensure transaction company is set for modal display
                        if (transactionData) {
                            // Create transaction object if it doesn't exist
                            if (!this.storedAnalysisResult.transaction) {
                                this.storedAnalysisResult.transaction = {};
                            }
                            this.storedAnalysisResult.transaction.company = company;
                        }
                        config.debug('🏢 Updated stored analysis result with company:', company);
                    } else {
                        // If storedAnalysisResult is null, create a minimal structure for modal
                        config.debug('🏢 Creating minimal storedAnalysisResult for modal with company:', company);
                        this.storedAnalysisResult = {
                            unified: { company: company },
                            portfolio: { company: company },
                            transaction: transactionData ? { company: company } : null
                        };
                    }
                } else {
                    company = 'Unknown';
                    config.warn('⚠️ No portfolio data available for company detection');
                }
                
                // Validate company and provide user feedback
                if (['Mixed', 'Mismatch'].includes(company)) {
                    const errorMsg = company === 'Mixed' 
                        ? 'Portfolio contains multiple different companies. Please ensure all entries have the same instrument.'
                        : 'Portfolio and transaction files contain different companies. Please ensure both files are for the same company.';
                    
                    this.hideLoading(); // Hide spinner before showing error
                    this.showError(`Company Validation Error: ${errorMsg}`);
                    return;
                }
                
                config.debug('🔍 Final Detection Results:', {
                    language: analysisResult.unified.language,
                    currency: analysisResult.unified.currency, 
                    company: company,
                    hasTransactions: !!transactionData
                });

                // Save to database with detection results including currency and company
                await equateDB.savePortfolioData(portfolioData, transactionData, portfolioData.userId, isEnglish, company, this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency, analysisResult?.transaction?.company || company);
                
                // Set flag to prevent auto-cache-load of just-saved data
                this.justProcessedFreshUpload = true;
                
                // Update drop zone indicators to reflect current cache state (especially after transaction rejection)
                const updatedDbInfo = await equateDB.getDatabaseInfo();
                await this.updateDropZonesForCachedData(updatedDbInfo);
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

            // Add startDate to portfolioData for optimization (single source of truth)
            if (portfolioData && portfolioData.entries && portfolioData.entries.length > 0) {
                // Find earliest transaction date from portfolio entries
                const dates = portfolioData.entries
                    .map(entry => entry.allocationDate)
                    .filter(date => date && date.trim() !== '')
                    .sort();
                
                if (dates.length > 0) {
                    portfolioData.startDate = dates[0];
                    config.debug(`🚀 OPTIMIZATION: Added startDate to portfolioData: ${portfolioData.startDate} (reducing data loading from ~200K to ~60K entries)`);
                } else {
                    config.warn('No valid dates found in portfolio entries - no optimization possible');
                }
            }

            // V2 TARGETED LOADING: Load historical currency data (precise filtering handled by ReferencePointsBuilder)
            const historicalCurrencyData = await this.currencyService.loadHistoricalData(portfolioData);

            // Set up calculator
            await portfolioCalculator.setPortfolioData(portfolioData, transactionData, this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency, {
                currencyData: historicalCurrencyData
            });
            
            // Debug: Check company value before calling loadHistoricalPrices  
            const companyForLoading = this.selectedCompany || company;
            config.debug('🔍 About to call loadHistoricalPrices with company:', companyForLoading);
            
            // Ensure historical prices are loaded (conditionally based on company)
            try {
                await this.loadHistoricalPrices(companyForLoading, portfolioData);
            } catch (error) {
                if (error.message.includes('not supported')) {
                    // Show user-friendly error for unsupported companies
                    this.hideLoading(); // Hide spinner before showing error
                    this.showError(`Company Not Supported: ${error.message}`);
                    return;
                } else {
                    // Re-throw other errors
                    throw error;
                }
            }
            
            // IMPORTANT: Initialize currency selector with updated quality scores
            await this.currencyService.initializeSelector();
            
            // Use historical prices for supported companies
            if (company && !['Unknown', 'Mixed', 'Mismatch', 'Other'].includes(company)) {
                const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                const historicalPrices = await equateDB.getHistoricalPrices(currentCurrency);
                config.debug(`Retrieved historical prices from DB for currency ${currentCurrency}:`, {
                    count: historicalPrices.length,
                    firstFew: historicalPrices.slice(0, 3),
                    lastFew: historicalPrices.slice(-3)
                });
                
                if (historicalPrices.length > 0) {
                    // V2 UNIFIED APPROACH: All currencies handled by enhanced reference points
                    const detectedCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                    config.debug(`🚀 V2 System: Setting historical prices for ${detectedCurrency} currency (unified approach)`);
                    
                    // 🎯 PORTFOLIO-AWARE FILTERING using enhanced reference points
                    let filteredHistoricalPrices = historicalPrices;
                    if (portfolioCalculator.enhancedReferencePoints && portfolioCalculator.enhancedReferencePoints.length > 0) {
                        const firstPurchaseDate = portfolioCalculator.enhancedReferencePoints[0].date;
                        const originalCount = historicalPrices.length;
                        
                        filteredHistoricalPrices = historicalPrices.filter(price => price.date >= firstPurchaseDate);
                        
                        const filteredCount = filteredHistoricalPrices.length;
                        const reductionPercent = originalCount > 0 ? ((originalCount - filteredCount) / originalCount * 100).toFixed(1) : 0;
                        
                        config.debug(`🎯 Portfolio-aware filtering: ${originalCount} → ${filteredCount} historical prices (${reductionPercent}% reduction, starting from ${firstPurchaseDate})`);
                    }
                    
                    // Set historical prices (currency switching handled by V2 instant O(1) system)
                    await portfolioCalculator.setHistoricalPrices(filteredHistoricalPrices);
                    this.timelineConversionStatus = { success: true, currency: detectedCurrency };
                } else {
                    // No historical prices available - generate synthetic timeline from user data
                    config.warn('No historical prices available, generating synthetic timeline using transaction reference points');
                    
                    const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                    const syntheticPrices = await this.createSyntheticTimelineFromUserData(portfolioData, transactionData, currentCurrency);
                    
                    if (syntheticPrices.length > 0) {
                        await portfolioCalculator.setHistoricalPrices(syntheticPrices, currentCurrency);
                        this.timelineConversionStatus = { success: true, currency: currentCurrency, synthetic: true };
                        
                        // BUGFIX: In synthetic mode, set currentPrice from enhanced reference points asOfDate, not external historical data
                        const asOfDatePoint = portfolioCalculator.enhancedReferencePoints?.find(point => point.type === 'asOfDate');
                        if (asOfDatePoint && asOfDatePoint.prices && asOfDatePoint.prices[currentCurrency]) {
                            portfolioCalculator.currentPrice = asOfDatePoint.prices[currentCurrency];
                            config.debug(`🔧 SYNTHETIC MODE: Set currentPrice to ${asOfDatePoint.prices[currentCurrency]} from reference points (not external historical data)`);
                        }
                    } else {
                        // Synthetic generation failed - this should not happen with valid transaction data
                        throw new Error(`Synthetic timeline generation failed for currency ${currentCurrency}. This indicates invalid transaction data.`);
                    }
                }
            } else {
                // For Non-Share-companies, generate synthetic timeline from transaction data
                config.debug(`🏢 Non-Share-company detected: ${company}, generating synthetic timeline from transaction data`);
                
                const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                const syntheticPrices = await this.createSyntheticTimelineFromUserData(portfolioData, transactionData, currentCurrency);
                
                if (syntheticPrices.length > 0) {
                    await portfolioCalculator.setHistoricalPrices(syntheticPrices, currentCurrency);
                    this.timelineConversionStatus = { success: true, currency: currentCurrency, synthetic: true };
                    
                    // BUGFIX: In synthetic mode, set currentPrice from enhanced reference points asOfDate, not external historical data
                    const asOfDatePoint = portfolioCalculator.enhancedReferencePoints?.find(point => point.type === 'asOfDate');
                    if (asOfDatePoint && asOfDatePoint.prices && asOfDatePoint.prices[currentCurrency]) {
                        portfolioCalculator.currentPrice = asOfDatePoint.prices[currentCurrency];
                        config.debug(`🔧 SYNTHETIC MODE: Set currentPrice to ${asOfDatePoint.prices[currentCurrency]} from reference points (not external historical data)`);
                    }
                } else {
                    // Synthetic generation failed - this should not happen with valid transaction data
                    throw new Error(`Synthetic timeline generation failed for currency ${currentCurrency}. This indicates invalid transaction data.`);
                }
            }

            // Calculate metrics (after AsOfDate is saved to IndexedDB)
            this.currentCalculations = await portfolioCalculator.calculate();
            config.debug('Calculations completed:', this.currentCalculations);

            // Show Data Analysis Summary modal with complete data for user confirmation
            const shouldContinue = await this.showDataAnalysisSummary();
            if (!shouldContinue) {
                config.debug('🚫 User cancelled analysis from Data Analysis Summary');
                this.hideLoading(); // Hide spinner before returning
                return; // User clicked Cancel, exit early
            }

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

            // Show currency selector now that we have portfolio data (unless in synthetic mode)
            const isInSyntheticMode = this.timelineConversionStatus && this.timelineConversionStatus.synthetic;
            if (!isInSyntheticMode) {
                this.currencyService.showSelector();
                config.debug('💱 Currency selector shown (multi-currency mode)');
            } else {
                this.currencyService.hideSelector();
                config.debug('🚫 Currency selector hidden (synthetic mode - no multi-currency conversion available)');
            }

            // Update button state
            this.checkFilesReady();

            // Clean up stored analysis data - only needed for single call optimization
            // Regeneration uses IndexedDB cache, not these in-memory variables
            this.storedAnalysisResult = null;
            this.storedPortfolioFile = null;
            this.storedTransactionFile = null;
            this.storedCachedData = null;
            this.storedPortfolioData = null;
            this.storedTransactionData = null;

            this.hideLoading();

        } catch (error) {
            config.error('Analysis failed:', error);
            
            // Clean up stored analysis data on error to prevent inconsistent state
            this.storedAnalysisResult = null;
            this.storedPortfolioFile = null;
            this.storedTransactionFile = null;
            this.storedCachedData = null;
            this.storedPortfolioData = null;
            this.storedTransactionData = null;
            
            this.hideLoading();
            this.showError('Analysis failed: ' + error.message);
        }
    }

    /**
     * Apply UI settings from file analysis result
     * @param {Object} unifiedAnalysis - Unified analysis result
     */
    async applyUISettings(unifiedAnalysis) {
        config.debug('🎨 Applying UI settings from analysis:', unifiedAnalysis);
        
        // Set UI language (user can still override via language selector)
        if (unifiedAnalysis.language !== translationManager.getCurrentLanguage()) {
            config.debug('🌐 Auto-switching UI language to:', unifiedAnalysis.language);
            this.setLanguagePreference(unifiedAnalysis.language);
        }
        
        // Store detected currency for display formatting
        if (this.currencyService) {
            this.currencyService.setDetectedCurrency(unifiedAnalysis.currency);
        }
        
        // Update user preference to match the uploaded file's currency
        await equateDB.savePreference('selectedCurrency', unifiedAnalysis.currency);
        
        config.debug('✅ UI settings applied:', {
            language: translationManager.getCurrentLanguage(),
            currency: this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency
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

            // Only use cached currency if no new file currency was detected
            const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
            if (!currentCurrency) {
                const cachedCurrency = cachedData.currency || null;
                if (this.currencyService) {
                    this.currencyService.setDetectedCurrency(cachedCurrency);
                }
            }
            const finalCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
            config.debug('💰 Retrieved currency from cached data:', finalCurrency);
            
            // V2 TARGETED LOADING: Load historical currency data (precise filtering handled by ReferencePointsBuilder)
            const historicalCurrencyData = await this.currencyService.loadHistoricalData(cachedData.portfolioData);
            
            // Set up calculator with cached data
            await portfolioCalculator.setPortfolioData(cachedData.portfolioData, cachedData.transactionData, finalCurrency, {
                currencyData: historicalCurrencyData
            });
            
            // For cached data, use stored company (already validated during storage)
            const company = cachedData.company || 'Other';
            
            // Ensure historical prices are loaded (conditionally based on company)
            try {
                await this.loadHistoricalPrices(company, cachedData.portfolioData);
            } catch (error) {
                if (error.message.includes('not supported')) {
                    // Show user-friendly error for unsupported companies
                    this.hideLoading(); // Hide spinner before showing error
                    this.showError(`Company Not Supported: ${error.message}`);
                    return;
                } else {
                    // Re-throw other errors
                    throw error;
                }
            }
            
            // IMPORTANT: Initialize currency selector even with cached data
            await this.currencyService.initializeSelector();
            
            // Use historical prices for supported companies (cached data)
            if (company && !['Unknown', 'Mixed', 'Mismatch', 'Other'].includes(company)) {
                const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                const historicalPrices = await equateDB.getHistoricalPrices(currentCurrency);
                config.debug(`Retrieved historical prices from DB for currency ${currentCurrency} (cached):`, {
                    count: historicalPrices.length,
                    firstFew: historicalPrices.slice(0, 3),
                    lastFew: historicalPrices.slice(-3)
                });
                
                if (historicalPrices.length > 0) {
                    // V2 UNIFIED APPROACH: All currencies handled by enhanced reference points (cached data)
                    const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                    config.debug(`🚀 V2 System: Setting historical prices for ${currentCurrency} currency (cached, unified approach)`);
                    
                    // 🎯 PORTFOLIO-AWARE FILTERING using cached data enhanced reference points
                    let filteredHistoricalPrices = historicalPrices;
                    if (portfolioCalculator.enhancedReferencePoints && portfolioCalculator.enhancedReferencePoints.length > 0) {
                        const firstPurchaseDate = portfolioCalculator.enhancedReferencePoints[0].date;
                        const originalCount = historicalPrices.length;
                        
                        filteredHistoricalPrices = historicalPrices.filter(price => price.date >= firstPurchaseDate);
                        
                        const filteredCount = filteredHistoricalPrices.length;
                        const reductionPercent = originalCount > 0 ? ((originalCount - filteredCount) / originalCount * 100).toFixed(1) : 0;
                        
                        config.debug(`🎯 Portfolio-aware filtering (cached): ${originalCount} → ${filteredCount} historical prices (${reductionPercent}% reduction, starting from ${firstPurchaseDate})`);
                    }
                    
                    // Set historical prices (currency switching handled by V2 instant O(1) system)
                    await portfolioCalculator.setHistoricalPrices(filteredHistoricalPrices);
                    this.timelineConversionStatus = { success: true, currency: currentCurrency };
                } else {
                    // No historical prices available - generate synthetic timeline from user data (cached)
                    config.warn('No historical prices available, generating synthetic timeline using transaction reference points (cached)');
                    
                    const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                    const syntheticPrices = await this.createSyntheticTimelineFromUserData(cachedData.portfolioData, cachedData.transactionData, currentCurrency);
                    
                    if (syntheticPrices.length > 0) {
                        await portfolioCalculator.setHistoricalPrices(syntheticPrices, currentCurrency);
                        this.timelineConversionStatus = { success: true, currency: currentCurrency, synthetic: true };
                        
                        // BUGFIX: In synthetic mode, set currentPrice from enhanced reference points asOfDate, not external historical data
                        const asOfDatePoint = portfolioCalculator.enhancedReferencePoints?.find(point => point.type === 'asOfDate');
                        if (asOfDatePoint && asOfDatePoint.prices && asOfDatePoint.prices[currentCurrency]) {
                            portfolioCalculator.currentPrice = asOfDatePoint.prices[currentCurrency];
                            config.debug(`🔧 SYNTHETIC MODE: Set currentPrice to ${asOfDatePoint.prices[currentCurrency]} from reference points (not external historical data)`);
                        }
                    } else {
                        // Synthetic generation failed - this should not happen with valid transaction data
                        throw new Error(`Synthetic timeline generation failed for currency ${currentCurrency}. This indicates invalid transaction data.`);
                    }
                }
            } else {
                // For Non-Share-companies, generate synthetic timeline from transaction data (cached)
                config.debug(`🏢 Non-Share-company detected: ${company}, generating synthetic timeline from transaction data (cached)`);
                
                const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                const syntheticPrices = await this.createSyntheticTimelineFromUserData(cachedData.portfolioData, cachedData.transactionData, currentCurrency);
                
                if (syntheticPrices.length > 0) {
                    await portfolioCalculator.setHistoricalPrices(syntheticPrices, currentCurrency);
                    this.timelineConversionStatus = { success: true, currency: currentCurrency, synthetic: true };
                    
                    // BUGFIX: In synthetic mode, set currentPrice from enhanced reference points asOfDate, not external historical data
                    const asOfDatePoint = portfolioCalculator.enhancedReferencePoints?.find(point => point.type === 'asOfDate');
                    if (asOfDatePoint && asOfDatePoint.prices && asOfDatePoint.prices[currentCurrency]) {
                        portfolioCalculator.currentPrice = asOfDatePoint.prices[currentCurrency];
                        config.debug(`🔧 SYNTHETIC MODE: Set currentPrice to ${asOfDatePoint.prices[currentCurrency]} from reference points (not external historical data)`);
                    }
                } else {
                    // Synthetic generation failed - this should not happen with valid transaction data
                    throw new Error(`Synthetic timeline generation failed for currency ${currentCurrency}. This indicates invalid transaction data.`);
                }
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

            // Show currency selector now that we have cached portfolio data (unless in synthetic mode)
            const isInSyntheticMode = this.timelineConversionStatus && this.timelineConversionStatus.synthetic;
            if (!isInSyntheticMode) {
                this.currencyService.showSelector();
                config.debug('💱 Currency selector shown (multi-currency mode, cached data)');
            } else {
                this.currencyService.hideSelector(); 
                config.debug('🚫 Currency selector hidden (synthetic mode - no multi-currency conversion available, cached data)');
            }

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
        if (price > 0) {
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

    // =============================================
    // RESULTS DISPLAY & METRICS
    // =============================================

    /**
     * Display analysis results
     */
    displayResults() {
        if (!this.currentCalculations) return;

        this.updateMetricsDisplay();
        this.updatePriceSource();
        this.currencyService.showWarning();
        
        // Update calculations tab with new data
        if (typeof calculationsTab !== 'undefined') {
            calculationsTab.toggleCalculationsDisplay(true);
            // Only render if calculations tab is currently active
            const calculationsPane = document.getElementById('calculationsPane');
            if (calculationsPane && calculationsPane.classList.contains('active')) {
                calculationsTab.render(this.currentCalculations);
            }
        }
        
        // Update UI language (labels only, after values are set)
        if (typeof translationManager !== 'undefined') {
            // Fire and forget - don't await to avoid blocking UI updates
            translationManager.updateUILanguage().catch(e => config.debug('Error updating UI language:', e));
            
            // Update export customizer translations
            if (window.exportCustomizer) {
                window.exportCustomizer.applyTranslations();
            }
        }
        
        // Enable export buttons
        const exportButtons = document.querySelectorAll('.export-btn');
        exportButtons.forEach(btn => btn.disabled = false);
    }

    /**
     * Diagnostic function to check element mappings
     */
    diagnoseElements() {
        
        const elements = ['userInvestment', 'companyMatch', 'freeShares', 'dividendIncome', 'totalInvestment', 
                         'totalReturn', 'currentValue', 'totalSold', 'totalValue', 'returnOnTotalInvestment',
                         'returnPercentage', 'xirrUserInvestment', 'availableShares', 'returnPercentageOnTotalInvestment', 'xirrTotalInvestment'];
        elements.forEach(elementId => {
            const element = document.getElementById(elementId);
            if (element) {
                const cardTitle = element.parentElement?.querySelector('h3')?.textContent;
            } else {
            }
        });
    }

    /**
     * Update metrics display
     */
    updateMetricsDisplay() {
        if (!this.currentCalculations) return;

        const currency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency; // Use service when available
        const decimals = 2; // Hardcoded to 2 decimals

        this.elements.userInvestment.textContent = currencyConverter.formatCurrency(this.currentCalculations.userInvestment, currency, decimals);
        this.elements.companyMatch.textContent = currencyConverter.formatCurrency(this.currentCalculations.companyMatch, currency, decimals);
        this.elements.freeShares.textContent = currencyConverter.formatCurrency(this.currentCalculations.freeShares, currency, decimals);
        this.elements.dividendIncome.textContent = currencyConverter.formatCurrency(this.currentCalculations.dividendIncome, currency, decimals);
        this.elements.totalInvestment.textContent = currencyConverter.formatCurrency(this.currentCalculations.totalInvestment, currency, decimals);
        
        this.elements.currentValue.textContent = currencyConverter.formatCurrency(this.currentCalculations.currentValue, currency, decimals);
        
        // Add profit/loss since download disclaimer
        if (portfolioCalculator.calculationBreakdowns && portfolioCalculator.calculationBreakdowns.currentPortfolio) {
            const breakdown = portfolioCalculator.calculationBreakdowns.currentPortfolio;
            const profitLossItem = breakdown.find(item => item.isProfitLoss);
            
            if (profitLossItem && this.elements.currentPortfolioDisclaimer) {
                const profitLoss = profitLossItem.currentValue;
                const isProfit = profitLoss >= 0;
                const formattedAmount = currencyConverter.formatCurrency(profitLoss, currency, decimals);
                
                this.elements.currentPortfolioDisclaimer.textContent = `Incl. ${formattedAmount} since download`;
                this.elements.currentPortfolioDisclaimer.className = `metric-disclaimer ${isProfit ? 'profit' : 'loss'}`;
                this.elements.currentPortfolioDisclaimer.style.display = 'block';
            }
        }
        this.elements.totalSold.textContent = currencyConverter.formatCurrency(this.currentCalculations.totalSold, currency, decimals);
        this.elements.totalValue.textContent = currencyConverter.formatCurrency(this.currentCalculations.totalValue, currency, decimals);
        
        const totalReturn = this.currentCalculations.totalReturn;
        this.elements.totalReturn.textContent = currencyConverter.formatCurrency(totalReturn, currency, decimals);
        this.elements.totalReturn.className = `metric-value ${totalReturn >= 0 ? 'positive' : 'negative'}`;
        
        const returnOnTotalInvestment = this.currentCalculations.returnOnTotalInvestment;
        this.elements.returnOnTotalInvestment.textContent = currencyConverter.formatCurrency(returnOnTotalInvestment, currency, decimals);
        this.elements.returnOnTotalInvestment.className = `metric-value ${returnOnTotalInvestment >= 0 ? 'positive' : 'negative'}`;
        
        const returnPercentage = this.currentCalculations.returnPercentage;
        this.elements.returnPercentage.textContent = currencyConverter.formatPercentage(returnPercentage, decimals);
        this.elements.returnPercentage.className = `metric-value ${returnPercentage >= 0 ? 'positive' : 'negative'}`;
        
        const returnPercentageOnTotalInvestment = this.currentCalculations.returnPercentageOnTotalInvestment;
        this.elements.returnPercentageOnTotalInvestment.textContent = currencyConverter.formatPercentage(returnPercentageOnTotalInvestment, decimals);
        this.elements.returnPercentageOnTotalInvestment.className = `metric-value ${returnPercentageOnTotalInvestment >= 0 ? 'positive' : 'negative'}`;
        
        // XIRR metrics with styling
        this.elements.xirrUserInvestment.textContent = currencyConverter.formatPercentage(this.currentCalculations.xirrUserInvestment, decimals);
        this.elements.xirrUserInvestment.className = `metric-value ${this.currentCalculations.xirrUserInvestment >= 0 ? 'positive' : 'negative'}`;
        
        this.elements.xirrTotalInvestment.textContent = currencyConverter.formatPercentage(this.currentCalculations.xirrTotalInvestment, decimals);
        this.elements.xirrTotalInvestment.className = `metric-value ${this.currentCalculations.xirrTotalInvestment >= 0 ? 'positive' : 'negative'}`;
        
        this.elements.availableShares.textContent = this.formatNumber(this.currentCalculations.availableShares, 3);
    }

    /**
     * Update price source indicator
     */
    updatePriceSource() {
        if (!this.currentCalculations) return;

        const indicator = this.elements.priceSource;
        const text = indicator.querySelector('.indicator-text');
        
        const currency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
        const currencySymbol = currencyConverter.getCurrencySymbol(currency);
        
        if (this.currentCalculations.priceSource === 'manual') {
            indicator.classList.add('manual');
            text.innerHTML = `<strong>${translationManager.t('price_label')}:</strong> ${currencySymbol}${this.currentCalculations.currentPrice.toFixed(2)} (manual input)`;
        } else {
            indicator.classList.remove('manual');
            text.innerHTML = `<strong>${translationManager.t('price_label')}:</strong> ${currencySymbol}${this.currentCalculations.currentPrice.toFixed(2)} ${translationManager.t('as_of_label')} ${this.currentCalculations.priceDate}`;
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
            // Skip auto-cache-load if we just processed fresh upload
            if (this.justProcessedFreshUpload) {
                config.debug('⚡ Skipping auto-cache-load - just processed fresh upload');
                return;
            }
            
            // Check if IndexedDB is available
            if (!window.indexedDB) {
                config.error('❌ IndexedDB not supported in this browser');
                this.elements.cacheInfo.textContent = 'IndexedDB not supported';
                return;
            }
            
            const dbInfo = await equateDB.getDatabaseInfo();
            
            if (dbInfo && dbInfo.hasPortfolioData) {
                const lastUpload = new Date(dbInfo.lastUpload);
                const daysAgo = Math.floor((Date.now() - lastUpload.getTime()) / (1000 * 60 * 60 * 24));
                
                // Update drop zones to show cached data status
                await this.updateDropZonesForCachedData(dbInfo);
                
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

    // =============================================
    // HISTORICAL DATA MANAGEMENT
    // =============================================


    /**
     * Load historical prices using environment-aware approach (like Tourploeg)
     * Dynamically loads hist.xlsx for supported Shares-companies (with historical price data)
     * @param {String} company - Instrument name from portfolio (e.g., "IBM shares")
     * @param {Object} portfolioData - Portfolio data with optional startDate for optimization
     */
    async loadHistoricalPrices(company = 'Unknown', portfolioData = null) {
        const startDate = portfolioData?.startDate || null;
        config.debug(`🔍 loadHistoricalPrices called with company: "${company}", startDate: ${startDate}`);
        try {
            // Skip loading for invalid companies
            if (!company || ['Unknown', 'Mixed', 'Mismatch', 'Other'].includes(company)) {
                config.debug(`🏢 Company is "${company}", skipping hist.xlsx loading (no historical data available)`);
                return;
            }

            // Check if we already have historical prices for this company
            const existingPrices = await equateDB.getHistoricalPrices();
            if (existingPrices.length > 0) {
                config.debug(`Historical prices loaded from cache: ${existingPrices.length} entries`);
                return;
            }

            config.debug(`🏢 Loading historical prices for "${company}" from hist.xlsx...`);

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

            // Determine which file to load based on detected company (Sprint 3)
            let histFile;
            let useCompanyFile = false;
            
            if (this.selectedCompany) {
                histFile = `hist_${this.selectedCompany}.xlsx`;
                useCompanyFile = true;
                config.debug(`🏢 Attempting to load company-specific file: ${histFile}`);
            } else {
                histFile = 'hist_base.xlsx'; // Always use base file as fallback
                config.debug('🏢 Loading base historical file: hist_base.xlsx');
            }
            
            // Try to load historical prices with cache buster
            const cacheBuster = `?v=${Date.now()}`;
            const fullUrl = baseUrl + histFile + cacheBuster;
            config.debug('Attempting to fetch:', fullUrl);
            
            const response = await fetch(fullUrl);
            
            config.debug('Fetch response:', {
                status: response.status,
                statusText: response.statusText,
                ok: response.ok,
                url: response.url
            });
            
            if (!response.ok) {
                // If company-specific file failed, try fallback to hist_base.xlsx
                if (useCompanyFile) {
                    config.warn(`Company-specific file ${histFile} not found (${response.status}), falling back to hist_base.xlsx`);
                    
                    // Retry with hist_base.xlsx
                    const fallbackUrl = baseUrl + 'hist_base.xlsx' + cacheBuster;
                    const fallbackResponse = await fetch(fallbackUrl);
                    
                    if (!fallbackResponse.ok) {
                        config.warn(`Fallback to hist_base.xlsx also failed: ${fallbackResponse.status} ${fallbackResponse.statusText} - calculations will use portfolio market price`);
                        return;
                    }
                    
                    histFile = 'hist_base.xlsx';
                    useCompanyFile = false;
                    
                    config.debug('✅ Successfully fell back to hist_base.xlsx');
                    
                    const arrayBuffer = await fallbackResponse.arrayBuffer();
                    const data = XLSX.read(arrayBuffer, { type: 'array' });
                    
                    return this.processHistoricalData(data, histFile, useCompanyFile, company, portfolioData);
                } else {
                    config.warn(`Historical price file fetch failed: ${response.status} ${response.statusText} - calculations will use portfolio market price`);
                    return;
                }
            }

            const arrayBuffer = await response.arrayBuffer();
            const data = XLSX.read(arrayBuffer, { type: 'array' });
            
            return this.processHistoricalData(data, histFile, useCompanyFile, company, portfolioData);
        } catch (error) {
            config.error('❌ Error in loadHistoricalPrices:', error);
            throw error;
        }
    }

    /**
     * Process historical data from Excel file (extracted for reuse by fallback logic)
     */
    async processHistoricalData(data, histFile, useCompanyFile, detectedInstrument = null, portfolioData = null) {
        try {
            const startDate = portfolioData?.startDate || null;
            config.debug(`🔍 processHistoricalData: startDate extracted = ${startDate}`);
            
            // Determine sheet name based on file type (Sprint 3)
            let worksheet, sheetName;
            if (useCompanyFile) {
                // Company-specific files use 'Share' sheet (singular)
                sheetName = 'Share';
                worksheet = data.Sheets[sheetName];
                if (!worksheet) {
                    config.error(`No "${sheetName}" sheet found in ${histFile}`);
                    return;
                }
            } else {
                // Base file uses 'Shares' sheet (plural)
                sheetName = 'Shares';
                worksheet = data.Sheets[sheetName];
                if (!worksheet) {
                    config.error(`No "${sheetName}" sheet found in ${histFile}`);
                    return;
                }
            }
            
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (rawData.length === 0) {
                config.error(`No data found in ${histFile} ${sheetName} tab`);
                return;
            }

            // Handle different data formats (Sprint 3)
            let companyColumnIndex = -1;
            let allCurrencyPrices = null;
            
            if (useCompanyFile) {
                // 🚀 REVOLUTIONARY: Company-specific multi-currency format extraction
                config.debug(`🏢 Processing company-specific multi-currency data for ${this.selectedCompany || 'detected company'}`);
                
                const headerRow = rawData[0]; // ['Date', 'EUR', 'USD', 'GBP', 'ARS', ...]
                const currencies = headerRow.slice(1).filter(currency => currency && currency.toString().trim());
                
                config.debug(`💰 Extracting ${startDate ? `filtered (from ${startDate})` : 'ALL'} currencies from company file: ${currencies.join(', ')}`);
                
                // Extract multi-currency entries - FILTERED during extraction for optimal performance
                let multiCurrencyEntries = [];
                let skippedCount = 0;
                for (let i = 1; i < rawData.length; i++) {
                    const row = rawData[i];
                    if (!row[0]) continue; // Skip rows without date
                    
                    // Parse date
                    let isoDate;
                    const dateStr = row[0];
                    if (typeof dateStr === 'string' && dateStr.includes('-')) {
                        const parts = dateStr.split('-');
                        if (parts.length === 3) {
                            const day = parts[0].padStart(2, '0');
                            const month = parts[1].padStart(2, '0');
                            const year = parts[2];
                            isoDate = `${year}-${month}-${day}`;
                        }
                    } else if (typeof dateStr === 'number') {
                        const jsDate = this.excelDateToJSDate(dateStr);
                        isoDate = jsDate.toISOString().split('T')[0];
                    }
                    
                    if (isoDate) {
                        // OPTIMIZATION: Skip dates before portfolio startDate during extraction
                        if (startDate && isoDate < startDate) {
                            skippedCount++;
                            continue;
                        }
                        
                        const entry = { date: isoDate };
                        
                        // Extract ALL currency values for this date
                        currencies.forEach((currency, index) => {
                            const value = parseFloat(row[index + 1]);
                            entry[currency] = (!isNaN(value) && value > 0) ? value : null;
                        });
                        
                        multiCurrencyEntries.push(entry);
                    }
                }
                
                if (startDate && skippedCount > 0) {
                    config.debug(`🚀 OPTIMIZATION: Extracted ${multiCurrencyEntries.length} multi-currency entries (skipped ${skippedCount} entries before ${startDate})`);
                } else {
                    config.debug(`🚀 Extracted ${multiCurrencyEntries.length} multi-currency entries with ${currencies.length} currencies each`);
                }
                
                // Add asOfDate to multi-currency database if missing
                if (portfolioCalculator.portfolioData?.asOfDate && portfolioCalculator.portfolioData?.entries?.[0]?.marketPrice) {
                    const asOfDate = portfolioCalculator.portfolioData.asOfDate;
                    const marketPrice = portfolioCalculator.portfolioData.entries[0].marketPrice;
                    const portfolioCurrency = portfolioCalculator.portfolioData.currency;
                    
                    // Check if asOfDate already exists in multi-currency data
                    const existingEntry = multiCurrencyEntries.find(entry => entry.date === asOfDate);
                    
                    if (!existingEntry) {
                        config.debug(`🎯 MULTICURRENCY: Adding missing asOfDate ${asOfDate} to multi-currency database`);
                        
                        // Create new entry with asOfDate ONLY in portfolio currency
                        const newEntry = { date: asOfDate };
                        newEntry[portfolioCurrency] = marketPrice;
                        // All other currencies remain undefined - that's perfect!
                        
                        multiCurrencyEntries.push(newEntry);
                        config.debug(`✅ MULTICURRENCY: AsOfDate ${asOfDate} added to multiCurrencyEntries with ${portfolioCurrency} ${marketPrice} (other currencies: undefined)`);
                        config.debug(`📊 MULTICURRENCY: Total entries now: ${multiCurrencyEntries.length} (before save to IndexedDB)`);
                    } else {
                        config.debug(`ℹ️ MULTICURRENCY: AsOfDate ${asOfDate} already exists in multi-currency data - no addition needed`);
                    }
                }
                
                // Save multi-currency data to database for instant currency switching
                await equateDB.saveMultiCurrencyPrices(multiCurrencyEntries);
                config.debug(`💾 MULTICURRENCY: Saved ${multiCurrencyEntries.length} entries to IndexedDB multiCurrencyPrices store`);
                
                // CRITICAL: Extract current currency historical prices for timeline generation
                const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
                config.debug(`🔍 EXTRACTION: Current currency detected as: ${currentCurrency}`);
                config.debug(`🔍 EXTRACTION: Multi-currency entries sample:`, multiCurrencyEntries.slice(0, 2));
                
                const priceEntries = multiCurrencyEntries
                    .filter(entry => entry[currentCurrency] !== null && entry[currentCurrency] !== undefined)
                    .map(entry => ({
                        date: entry.date,
                        price: entry[currentCurrency],
                        currency: currentCurrency
                    }));
                
                config.debug(`🔍 EXTRACTION: Filtered ${priceEntries.length} entries with valid ${currentCurrency} prices from ${multiCurrencyEntries.length} total entries`);
                
                if (priceEntries.length > 0) {
                    await equateDB.saveHistoricalPrices(priceEntries);
                    config.debug(`💾 HISTORICAL: Extracted and saved ${priceEntries.length} ${currentCurrency} historical prices to historicalPrices store`);
                } else {
                    config.warn(`⚠️ No ${currentCurrency} prices found in multi-currency data - timeline will use synthetic mode`);
                }
                
                config.debug(`🚀 V2: Using pure multi-currency approach with ${currentCurrency} historical prices extracted`);
                return; // Exit early - multi-currency processing complete
                
            } else {
                // Base multi-company format: Date | Allianz Share | IBM Share | ...
                config.debug(`🏢 Processing base multi-company format (single currency extraction)`);
                
                const headerRow = rawData[0];
                const targetInstrument = detectedInstrument || 'Unknown';
                
                for (let i = 1; i < headerRow.length; i++) {
                    const columnName = headerRow[i];
                    if (columnName && columnName.toString().trim() === targetInstrument) {
                        companyColumnIndex = i;
                        break;
                    }
                }
                
                if (companyColumnIndex === -1) {
                    const availableColumns = headerRow.slice(1).filter(col => col && col.toString().trim());
                    config.error(`Instrument "${targetInstrument}" not found in ${histFile}. Available instruments:`, availableColumns);
                    throw new Error(`Instrument "${targetInstrument}" not supported. Available instruments: ${availableColumns.join(', ')}`);
                }
                
                config.debug(`Found "${targetInstrument}" in column ${companyColumnIndex} of ${histFile}`);
            }

            // Parse historical data using the dynamically found column
            const priceEntries = [];
            for (let i = 1; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row[0] || row[companyColumnIndex] === undefined) continue;

                // Handle DD-MM-YYYY string dates (our new format)
                const dateStr = row[0];
                const price = parseFloat(row[companyColumnIndex]); // Dynamic company column

                if (dateStr && !isNaN(price)) {
                    // Parse DD-MM-YYYY format to ISO date
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
                    
                    if (isoDate) {
                        priceEntries.push({
                            date: isoDate,
                            price: price,
                            currency: currentCurrency // Store currency with each price entry - NO DEFAULT!
                        });
                    } else {
                        config.warn(`Skipping invalid date format in row ${i}:`, { dateStr, price, row });
                    }
                } else {
                    config.warn(`Skipping invalid data row ${i}:`, { dateStr, price, row });
                }
            }

            if (priceEntries.length === 0) {
                config.warn('No valid price entries found in hist.xlsx');
                return;
            }

            // Sort by date and save to database
            priceEntries.sort((a, b) => new Date(a.date) - new Date(b.date));
            
            // 🎯 PORTFOLIO-AWARE FILTERING: Only keep data from portfolio timeline onwards
            // For now, skip this optimization during initial load to maintain V2 consistency
            // TODO: Implement deferred filtering after enhanced reference points are available
            config.debug('📅 Skipping historical price filtering during initial load (V2 consistency)');
            config.debug('📅 Future optimization: Filter historical prices to portfolio timeline after enhanced reference points are built');
            
            // V2 UNIFIED: Add asOfDate from portfolio if available (all currencies supported)
            // V2 System: No currency contamination concerns with enhanced reference points
            const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
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
                        config.debug(`🚀 V2 System: Added asOfDate ${asOfDate} (${marketPrice}) to historical prices (${currentCurrency} unified approach)`);
                        
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
            config.warn('Could not process historical data:', error.message);
            config.debug('Calculations will use portfolio market price as fallback');
        }
    }






    /**
     * Calculate working days between two dates (Monday-Friday)
     * @param {Date} startDate - Start date
     * @param {Date} endDate - End date  
     * @returns {number} Number of working days
     */
    calculateWorkingDaysBetweenDates(startDate, endDate) {
        let workingDays = 0;
        let currentDate = new Date(startDate);

        while (currentDate <= endDate) {
            if (this.isWorkingDay(currentDate)) {
                workingDays++;
            }
            currentDate.setDate(currentDate.getDate() + 1);
        }

        return workingDays;
    }

    /**
     * Check if a date is a working day (Monday-Friday)
     * @param {Date} date - Date to check
     * @returns {boolean} True if working day
     */
    isWorkingDay(date) {
        const dayOfWeek = date.getDay();
        return dayOfWeek >= 1 && dayOfWeek <= 5; // Monday (1) to Friday (5)
    }

    /**
     * Convert Excel date serial number to JavaScript Date
     */
    excelDateToJSDate(excelDate) {
        // Excel's epoch is 1900-01-01, but it incorrectly treats 1900 as a leap year
        // Use UTC to avoid timezone shifts that create artificial weekend dates
        const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // December 30, 1899 UTC
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
            await translationManager.setLanguage(language);
            config.debug('🌐 Restored language preference:', language);

            // Restore currency preference
            const savedCurrency = await equateDB.getPreference('selectedCurrency');
            if (savedCurrency) {
                if (this.currencyService) {
                    this.currencyService.setDetectedCurrency(savedCurrency);
                }
            }
        } catch (error) {
            config.error('Error loading preferences:', error);
        }
    }

    // =============================================
    // CURRENCY MANAGEMENT
    // =============================================





    /**
     * Show temporary message with auto-dismiss
     */
    showTemporaryMessage(message, type = 'info', duration = 5000) {
        const messageDiv = document.createElement('div');
        messageDiv.style.cssText = `
            position: fixed; top: 20px; right: 20px; z-index: 10000;
            background: ${type === 'warning' ? '#f59e0b' : '#3b82f6'}; 
            color: white; padding: 15px 20px;
            border-radius: 8px; max-width: 400px; 
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        `;
        messageDiv.innerHTML = `
            ${message}
            <button onclick="this.parentElement.remove()" style="
                background: none; border: none; color: white; float: right; 
                margin-left: 10px; cursor: pointer; font-size: 18px;
            ">×</button>
        `;
        document.body.appendChild(messageDiv);
        
        // Auto-remove after duration
        setTimeout(() => {
            if (messageDiv.parentElement) {
                messageDiv.remove();
            }
        }, duration);
    }




    /**
     * Set language and save preference
     */
    setLanguagePreference(language) {
        try {
            // Track language switch event
            if (typeof gtag !== 'undefined') {
                gtag('event', 'language_switch', {
                    'language': language,
                    'event_category': 'settings'
                });
            }
            
            this.elements.languageSelect.value = language;
            // Fire and forget - don't await to avoid blocking UI
            translationManager.setLanguage(language).catch(e => config.debug('Error setting language:', e));
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
            // Track number format switch event
            if (typeof gtag !== 'undefined') {
                gtag('event', 'number_format_switch', {
                    'format': format,
                    'event_category': 'settings'
                });
            }
            
            this.numberFormat = format;
            equateDB.savePreference('numberFormat', format);
            
            // Refresh all displays if we have calculation data
            if (this.currentCalculations) {
                console.log('Refreshing displays with new format:', format);
                this.updateMetricsDisplay();
                this.refreshCharts();
                
                // Refresh calculations tab if it exists
                if (typeof calculationsTab !== 'undefined' && calculationsTab.isInitialized) {
                    calculationsTab.render(this.currentCalculations);
                    config.debug('🔢 Refreshed calculations tab for number format change');
                }
                
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
            return;
        }
        
        // Get the same parameters as the original chart creation
        const calc = this.currentCalculations;
        const currencySymbol = currencyConverter.getCurrencySymbol(calc.currency);
        const fontColor = getComputedStyle(document.documentElement).getPropertyValue('--text-primary').trim();
        const bgColor = getComputedStyle(document.documentElement).getPropertyValue('--bg-secondary').trim();
        
        // Recreate all charts to apply new number formatting
        this.createPortfolioChart();
        this.createBreakdownCharts();
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

    // =============================================
    // DYNAMIC INSTRUMENT-TO-COMPANY MAPPING (Sprint 3)
    // =============================================
    
    /**
     * Load hist_base.xlsx metadata for dynamic instrument-to-company mapping
     */
    async loadBaseMetadata() {
        try {
            config.debug('🏢 Loading hist_base.xlsx metadata for instrument mapping...');
            
            // Always load hist_base.xlsx at startup for metadata
            const baseUrl = this.getBaseUrl();
            const fullUrl = baseUrl + 'hist_base.xlsx' + this.getCacheBuster();
            
            const response = await fetch(fullUrl);
            if (!response.ok) {
                throw new Error(`Failed to load hist_base.xlsx: ${response.status}`);
            }
            
            const arrayBuffer = await response.arrayBuffer();
            const data = XLSX.read(arrayBuffer, { type: 'array' });
            const worksheet = data.Sheets['Shares'];
            
            // Store metadata for instrument-to-company mapping
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            this.instrumentToCompanyMap = this.buildInstrumentMapping(rawData);
            
            config.debug('🏢 Base metadata loaded successfully:', this.instrumentToCompanyMap);
            
        } catch (error) {
            config.error('❌ Error loading base metadata:', error);
            // Don't block app initialization, but log error
            config.warn('⚠️ Continuing without company-specific optimization');
        }
    }

    /**
     * Build instrument-to-company mapping from hist_base.xlsx metadata
     * @param {Array} rawData - Raw Excel data from hist_base.xlsx
     * @returns {Object} Mapping of instrument names to company names
     */
    buildInstrumentMapping(rawData) {
        try {
            if (rawData.length < 2) {
                config.warn('⚠️ Insufficient metadata rows in hist_base.xlsx');
                return {};
            }
            
            const instrumentRow = rawData[0]; // Row 0: ["Date", "Allianz Share", "IBM Share", ...]
            const companyRow = rawData[1];    // Row 1: ["Company", "Allianz", "IBM", ...]
            
            const mapping = {};
            
            // Skip first column (Date/Company headers) and build mapping
            for (let i = 1; i < instrumentRow.length && i < companyRow.length; i++) {
                const instrument = instrumentRow[i];
                const company = companyRow[i];
                
                if (instrument && company) {
                    mapping[instrument.toString().trim()] = company.toString().trim();
                }
            }
            
            config.debug('🏢 Built instrument-to-company mapping:', mapping);
            return mapping;
            
        } catch (error) {
            config.error('❌ Error building instrument mapping:', error);
            return {};
        }
    }

    /**
     * Map detected instrument to company name for file loading
     * @param {string} instrument - Detected instrument name (e.g., "Allianz Share")
     * @returns {string|null} Company name for file loading (e.g., "Allianz") or null if not found
     */
    async mapInstrumentToCompany(instrument) {
        try {
            if (!instrument || typeof instrument !== 'string') {
                config.warn('⚠️ Invalid instrument provided for mapping:', instrument);
                return null;
            }
            
            // Use cached mapping if available
            if (this.instrumentToCompanyMap && this.instrumentToCompanyMap[instrument]) {
                const companyName = this.instrumentToCompanyMap[instrument];
                config.debug(`🏢 Mapped instrument "${instrument}" → company "${companyName}"`);
                return companyName;
            }
            
            config.warn(`⚠️ No company mapping found for instrument: ${instrument}`);
            return null;
            
        } catch (error) {
            config.error('❌ Error mapping instrument to company:', error);
            return null;
        }
    }

    /**
     * Get base URL for file loading (reused from loadHistoricalPrices)
     */
    getBaseUrl() {
        const currentUrl = window.location.href;
        
        if (currentUrl.includes('github.io')) {
            const parts = currentUrl.split('.github.io')[0].split('//')[1];
            const repoName = currentUrl.split('/')[3] || 'Equate';
            return `https://raw.githubusercontent.com/${parts}/${repoName}/main/`;
        } else {
            return './';
        }
    }

    /**
     * Get cache buster for file loading
     */
    getCacheBuster() {
        return '?v=' + new Date().getTime();
    }

    // =============================================
    // UI UTILITY METHODS
    // =============================================

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
     * Update portfolio chart labels for language changes (efficient)
     */
    updatePortfolioChartLabels() {
        if (!this.currentCalculations || !document.getElementById('portfolioChart')) {
            return;
        }

        const currency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
        const currencySymbol = currency ? currencyConverter.getCurrencySymbol(currency) : '';
        
        // Update layout labels only
        const layoutUpdate = {
            'xaxis.title': translationManager.t('chart_date_axis_label'),
            'yaxis.title': `${translationManager.t('chart_value_axis_label')} (${currencySymbol})`,
            'yaxis2.title': `${translationManager.t('stock_price_legend')} (${currencySymbol})`
        };

        // Get current chart data to determine existing traces
        const chartElement = document.getElementById('portfolioChart');
        
        // Check if Plotly chart has been initialized (has data property and _fullLayout)
        if (!chartElement.data || !chartElement._fullLayout) {
            config.debug('📊 Chart not yet initialized, skipping label update');
            return;
        }
        
        const plotData = chartElement.data || [];
        const numTraces = plotData.length;
        
        if (numTraces === 0) {
            // No chart data yet, but chart is initialized - safe to update layout
            try {
                Plotly.relayout('portfolioChart', layoutUpdate);
            } catch (error) {
                config.debug('📊 Chart layout update failed, chart may be in transition:', error.message);
            }
            return;
        }

        // Build trace names array based on actual traces
        const traceNames = [
            translationManager.t('portfolio_value_legend'),           // Trace 0: Always exists
            translationManager.t('profit_loss_legend'),              // Trace 1: Always exists  
            translationManager.t('stock_price_legend')               // Trace 2: Always exists
        ];
        
        // Add conditional trace names if they exist
        if (numTraces > 3) traceNames.push(translationManager.t('transaction_legend')); // Trace 3: Transactions
        if (numTraces > 4) traceNames.push(translationManager.t('manual_price_scenario_legend')); // Trace 4: Manual price on portfolio
        if (numTraces > 5) traceNames.push(`${translationManager.t('manual_price_scenario_legend')} (${translationManager.t('profit_loss_legend')})`); // Trace 5: Manual price on profit/loss

        // Create trace indices array for existing traces only
        const traceIndices = Array.from({length: numTraces}, (_, i) => i);

        const traceUpdate = { 'name': traceNames };

        // Apply updates efficiently with error handling
        try {
            Plotly.relayout('portfolioChart', layoutUpdate);
            Plotly.restyle('portfolioChart', traceUpdate, traceIndices);
        } catch (error) {
            config.debug('📊 Chart update failed, chart may be in transition:', error.message);
            return;
        }
        
        config.debug('📊 Portfolio chart labels updated efficiently (no data regeneration)');
    }

    /**
     * Update breakdown chart labels for language changes (efficient) 
     */
    updateBreakdownChartLabels() {
        if (!this.currentCalculations) {
            return;
        }

        const currency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
        const currencySymbol = currency ? currencyConverter.getCurrencySymbol(currency) : '';

        // Update pie chart
        const pieChartElement = document.getElementById('investmentPieChart');
        if (pieChartElement && pieChartElement.data && pieChartElement.data.length > 0) {
            const pieTraceUpdate = {
                'labels': [
                    [translationManager.t('your_investment_legend'),
                     translationManager.t('company_match_legend'),
                     translationManager.t('free_shares_legend'), 
                     translationManager.t('dividend_income_legend')]
                ]
            };
            Plotly.restyle('investmentPieChart', pieTraceUpdate, [0]);
        }

        // Update bar chart  
        const barChartElement = document.getElementById('performanceBarChart');
        if (barChartElement && barChartElement.data && barChartElement.data.length > 0) {
            const barLayoutUpdate = {
                'yaxis.title': `${translationManager.t('chart_value_axis_label')} (${currencySymbol})`
            };
            
            const barTraceUpdate = {
                'name': [
                    translationManager.t('your_investment_legend'),
                    translationManager.t('company_match_legend'),
                    translationManager.t('free_shares_legend'),
                    translationManager.t('dividend_income_legend'),
                    translationManager.t('total_investment_base_legend'),
                    translationManager.t('current_portfolio_legend'),
                    translationManager.t('total_sold_legend'),
                    translationManager.t('return_on_your_investment_legend'),
                    translationManager.t('return_on_total_investment_legend')
                ]
            };

            // Only update traces that actually exist
            const actualTraceCount = barChartElement.data.length;
            const traceIndices = Array.from({length: Math.min(actualTraceCount, 9)}, (_, i) => i);

            Plotly.relayout('performanceBarChart', barLayoutUpdate);
            Plotly.restyle('performanceBarChart', barTraceUpdate, traceIndices);
        }

        config.debug('📊 Breakdown chart labels updated efficiently (no data regeneration)');
    }

    // =============================================
    // CHART GENERATION
    // =============================================

    /**
     * Create interactive portfolio timeline chart
     */
    async createPortfolioChart() {
        if (!this.currentCalculations || !portfolioCalculator.portfolioData) {
            return;
        }

        // Check for timeline conversion failures and show error message
        if (this.timelineConversionStatus && !this.timelineConversionStatus.success) {
            const errorMessage = this.timelineConversionStatus.error || 'Unknown conversion error';
            const currency = this.timelineConversionStatus.currency || 'non-EUR';
            
            document.getElementById('portfolioChart').innerHTML = 
                '<div style="display: flex; align-items: center; justify-content: center; height: 100%; color: #e74c3c; flex-direction: column; padding: 20px; text-align: center;">' +
                '<div style="font-size: 24px; margin-bottom: 15px;">⚠️</div>' +
                '<div style="font-size: 16px; font-weight: bold; margin-bottom: 10px;">Timeline Chart Unavailable</div>' +
                '<div style="font-size: 14px; margin-bottom: 15px; line-height: 1.4;">' +
                `Unable to generate timeline chart for ${currency} portfolio.<br>` +
                'Insufficient currency reference data for conversion.</div>' +
                '<div style="font-size: 12px; color: #888; line-height: 1.3;">' +
                'Portfolio cards and other metrics are still accurate and available.<br>' +
                'Timeline requires historical price conversion which needs purchase/transaction dates.</div>' +
                '</div>';
            return;
        }

        try {
            // Get detected currency - NO FALLBACKS, if we don't have it, we don't use one
            let currency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
            
            // Only use calculations currency if it's not 'UNKNOWN'
            if (!currency && this.currentCalculations && this.currentCalculations.currency !== 'UNKNOWN') {
                currency = this.currentCalculations.currency;
            }
            
            config.debug('🔍 Chart currency resolution:', {
                detectedCurrency: this.detectedCurrency,
                calculationsOriginCurrency: this.currentCalculations?.currency,
                finalCurrency: currency
            });
            
            const currencySymbol = currency ? currencyConverter.getCurrencySymbol(currency) : '';
            
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
            let filteredData = [...timelineData];

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
                
                // Find manual price points from calculator's historical prices (where they were actually added)
                const historicalPrices = portfolioCalculator.historicalPrices || [];
                const manualPriceEntries = historicalPrices.filter(entry => 
                    Math.abs(entry.price - this.currentCalculations.currentPrice) < 0.01
                );
                
                if (manualPriceEntries.length > 0) {
                    // Convert historical price entries to timeline format and add to manualPricePoints
                    manualPriceEntries.forEach(entry => {
                        // Find corresponding timeline data or create a point
                        const timelinePoint = filteredData.find(d => d.date === entry.date);
                        if (timelinePoint) {
                            manualPricePoints.push(timelinePoint);
                            config.debug('🎯 Manual price point found at date:', entry.date, 'price:', entry.price);
                        } else {
                            // Create a manual price point for the timeline
                            // Calculate portfolio value for this specific date with the manual price
                            const sharesAtDate = this.currentCalculations.availableShares;
                            const portfolioValueAtManualPrice = sharesAtDate * entry.price;
                            
                            const manualPoint = {
                                date: entry.date,
                                currentPrice: entry.price,
                                portfolioValue: portfolioValueAtManualPrice,
                                profitLoss: portfolioValueAtManualPrice - this.currentCalculations.totalInvestment,
                                hasTransaction: false,
                                reason: 'Manual Price Scenario'
                            };
                            manualPricePoints.push(manualPoint);
                            config.debug('🎯 Manual price point created for date:', entry.date, 'price:', entry.price, 'portfolioValue:', portfolioValueAtManualPrice);
                        }
                    });
                } else {
                    config.debug('🎯 No manual price points found in historical prices');
                }
            }

            // Add manual price points to main trace data for continuous line
            if (manualPricePoints.length > 0) {
                manualPricePoints.forEach(point => {
                    dates.push(point.date);
                    portfolioValues.push(point.portfolioValue);
                    profitLoss.push(point.profitLoss);
                    stockPrices.push(point.currentPrice);
                });
                config.debug('🎯 Added manual price points to main trace data for continuous line');
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
                }
            ];

            // BUGFIX: Only show stock price trendline when NOT in synthetic mode
            // In synthetic mode, we don't have reliable historical stock price data
            const isInSyntheticMode = this.timelineConversionStatus && this.timelineConversionStatus.synthetic;
            if (!isInSyntheticMode) {
                traces.push({
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
                });
                config.debug('📈 Stock price trendline added to chart (non-synthetic mode)');
            } else {
                config.debug('🚫 Stock price trendline hidden (synthetic mode)');
            }

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

            // Add AsOfDate reference points (green diamonds like transactions)
            const asOfDatePoints = [];
            const calculatorHistoricalPrices = portfolioCalculator.historicalPrices || [];
            const asOfDateEntries = calculatorHistoricalPrices.filter(entry => entry.source === 'AsOfDate');
            
            config.debug('🔍 DEBUG AsOfDate display check:', {
                calculatorHistoricalPricesCount: calculatorHistoricalPrices.length,
                asOfDateEntriesCount: asOfDateEntries.length,
                asOfDateEntries: asOfDateEntries,
                filteredDataCount: filteredData.length,
                // Show last few historical prices to see their structure
                lastFewHistoricalPrices: calculatorHistoricalPrices.slice(-5),
                // Show any entries that might be AsOfDate related
                possibleAsOfDateEntries: calculatorHistoricalPrices.filter(entry => 
                    entry.date === '2025-08-04' || entry.source || entry.isAsOfDate
                )
            });
            
            if (asOfDateEntries.length > 0) {
                asOfDateEntries.forEach(entry => {
                    const timelinePoint = filteredData.find(d => d.date === entry.date);
                    if (timelinePoint) {
                        asOfDatePoints.push({
                            date: entry.date,
                            portfolioValue: timelinePoint.portfolioValue,
                            profitLoss: timelinePoint.profitLoss,
                            price: entry.price
                        });
                    }
                });
                
                if (asOfDatePoints.length > 0) {
                    traces.push({
                        x: asOfDatePoints.map(d => d.date),
                        y: asOfDatePoints.map(d => d.portfolioValue),
                        type: 'scatter',
                        mode: 'markers',
                        name: 'Download date',
                        marker: {
                            color: '#27ae60',
                            size: 12,
                            symbol: 'diamond',
                            line: { color: 'white', width: 2 }
                        },
                        hovertemplate: `<b>Download date</b><br>` +
                                     `${translationManager.t('chart_date_label')}: %{x|%d-%m-%Y}<br>` +
                                     `${translationManager.t('portfolio_value_legend')}: ${currencySymbol}%{y:,.2f}<br>` +
                                     `Reference Price: ${currencySymbol}${asOfDatePoints[0].price.toFixed(2)}<br>` +
                                     '<extra></extra>'
                    });
                }
            }

            // Add Last Known Price reference point (black diamond) if different from download date
            const lastKnownPricePoints = [];
            if (calculatorHistoricalPrices.length > 0) {
                // Find the latest historical price (excluding AsOfDate entries)
                const regularHistoricalPrices = calculatorHistoricalPrices.filter(entry => entry.source !== 'AsOfDate');
                if (regularHistoricalPrices.length > 0) {
                    const latestHistoricalPrice = regularHistoricalPrices[regularHistoricalPrices.length - 1];
                    const downloadDate = asOfDateEntries.length > 0 ? asOfDateEntries[0].date : null;
                    
                    // Only show black diamond if last known price date differs from download date
                    if (latestHistoricalPrice.date !== downloadDate) {
                        const timelinePoint = filteredData.find(d => d.date === latestHistoricalPrice.date);
                        if (timelinePoint) {
                            lastKnownPricePoints.push({
                                date: latestHistoricalPrice.date,
                                portfolioValue: timelinePoint.portfolioValue,
                                profitLoss: timelinePoint.profitLoss,
                                price: latestHistoricalPrice.price
                            });
                        }
                    }
                }
            }

            if (lastKnownPricePoints.length > 0) {
                traces.push({
                    x: lastKnownPricePoints.map(d => d.date),
                    y: lastKnownPricePoints.map(d => d.portfolioValue),
                    type: 'scatter',
                    mode: 'markers',
                    name: 'Last known price',
                    marker: {
                        color: '#2c3e50',
                        size: 12,
                        symbol: 'diamond',
                        line: { color: 'white', width: 2 }
                    },
                    hovertemplate: `<b>Last known price</b><br>` +
                                 `${translationManager.t('chart_date_label')}: %{x|%d-%m-%Y}<br>` +
                                 `${translationManager.t('portfolio_value_legend')}: ${currencySymbol}%{y:,.2f}<br>` +
                                 `Historical Price: ${currencySymbol}${lastKnownPricePoints[0].price.toFixed(2)}<br>` +
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

            // Layout configuration - conditionally dual Y-axis (left: portfolio values, right: stock price if not synthetic)
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

            // BUGFIX: Only add secondary y-axis when NOT in synthetic mode (for stock price trendline)
            if (!isInSyntheticMode) {
                layout.yaxis2 = {
                    title: `${translationManager.t('stock_price_legend')} (${currencySymbol})`,
                    titlefont: { color: '#95a5a6' },
                    tickfont: { color: '#95a5a6' },
                    tickformat: ',.2f',
                    overlaying: 'y',
                    side: 'right',
                    fixedrange: false,
                    showgrid: false
                };
                config.debug('📈 Secondary y-axis (y2) added for stock price trendline');
            } else {
                config.debug('🚫 Secondary y-axis (y2) hidden (synthetic mode)');
            }

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
            await this.updateTimelineDisclaimer(timelineData, currencySymbol);
            
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
    async refreshTimelineDisclaimer() {
        
        if (this.currentCalculations && this.lastTimelineData) {
            // NO FALLBACKS - if we don't have currency, we don't use one
            let currency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
            
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
            const currencySymbol = (currency && currency !== 'UNKNOWN') ? currencyConverter.getCurrencySymbol(currency) : '';
            await this.updateTimelineDisclaimer(this.lastTimelineData, currencySymbol);
        } else {
        }
    }

    /**
     * Update timeline disclaimer AND total sold card disclaimer based on totalSold and value discrepancies
     */
    async updateTimelineDisclaimer(timelineData, currencySymbol) {
        const timelineDisclaimerElement = document.getElementById('timelineDisclaimer');
        const totalSoldDisclaimerElement = document.getElementById('totalSoldDisclaimer');
        
        if (!timelineDisclaimerElement || !timelineData || timelineData.length === 0) {
            return;
        }

        const totalSold = this.currentCalculations.totalSold;
        
        // Check if cached transaction data exists
        const cachedData = await equateDB.getPortfolioData();
        const hasTransactionData = cachedData?.transactionData?.entries?.length > 0;
        
        
        // Show disclaimer only if totalSold = 0 AND no cached transaction data exists
        if (totalSold === 0 && !hasTransactionData) {
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
            const currency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
            const currencySymbol = currencyConverter.getCurrencySymbol(currency);
            
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

    // =============================================
    // CACHE MANAGEMENT
    // =============================================

    /**
     * Clear all cached data (DEPRECATED - use clearEverythingConfirm for development)
     * This function is kept for backward compatibility but now properly clears everything
     */
    async clearAllData() {
        try {
            // Track clear data event
            if (typeof gtag !== 'undefined') {
                gtag('event', 'clear_data_click', {
                    'event_category': 'user_interaction'
                });
            }
            
            this.showLoading('Clearing cached data...');
            
            // NEW: Actually clear ALL data (not just portfolio data)
            const stats = window.CacheManager?.getStats() || { totalEntries: 0 };
            const allStoreCounts = await getAllIndexedDBStoreCounts();
            const localData = getLocalStorageData();
            
            const totalItems = stats.totalEntries + 
                              Object.values(allStoreCounts).reduce((sum, count) => sum + count, 0) + 
                              localData.count;
            
            if (totalItems === 0) {
                this.hideLoading();
                alert('No cached data to clear.');
                return;
            }
            
            // Clear everything comprehensively
            if (window.CacheManager) window.CacheManager.clearAll();
            await equateDB.clearAllData();
            
            // Clear ALL IndexedDB stores (nuclear option)
            const storeNames = [
                'portfolioData', 'historicalPrices', 'userPreferences', 'manualPrices',
                'historicalCurrencyData', 'currencyQualityScores'
            ];
            
            for (const storeName of storeNames) {
                try {
                    const transaction = equateDB.db.transaction([storeName], 'readwrite');
                    const store = transaction.objectStore(storeName);
                    await new Promise((resolve, reject) => {
                        const request = store.clear();
                        request.onsuccess = () => resolve();
                        request.onerror = () => reject(request.error);
                    });
                } catch (error) {
                    config.debug(`⚠️ Could not clear store ${storeName}:`, error.message);
                }
            }
            
            localStorage.clear();
            
            // Clear in-memory currency data cache
            this.currencyService.clearCache();
            
            // Clear CurrencyConverter cache (fixes stale AsOfDate issue)
            if (window.currencyConverter) {
                window.currencyConverter.clear();
            }
            
            // Reset application state
            this.hasData = false;
            this.currentCalculations = null;
            portfolioCalculator.clearData();
            
            // Reset Results tab to no data state
            toggleResultsDisplay(false);
            
            // Hide currency selector when no data
            this.currencyService.hideSelector();
            
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
    async updateDropZonesForCachedData(dbInfo) {
        
        // Portfolio drop zone
        const portfolioZone = document.getElementById('portfolioZone');
        const portfolioInfo = document.getElementById('portfolioInfo');
        
        
        if (portfolioZone && portfolioInfo) {
            if (dbInfo.hasPortfolioData) {
                portfolioZone.classList.add('file-selected');
                
                // Hide the instruction text when showing cached data
                const instructionP = portfolioZone.querySelector('p');
                if (instructionP) {
                    instructionP.style.display = 'none';
                }
                
                // Show enhanced portfolio file information
                await this.displayEnhancedPortfolioInfo(portfolioInfo);
                portfolioInfo.classList.add('show');
            } else {
                portfolioZone.classList.remove('file-selected');
                
                // Show the instruction text when no cached data
                const instructionP = portfolioZone.querySelector('p');
                if (instructionP) {
                    instructionP.style.display = 'block';
                }
                
                portfolioInfo.innerHTML = '';
                portfolioInfo.classList.remove('show');
            }
        } else {
        }
        
        // Transaction drop zone - handle both cached and non-cached states
        const transactionsZone = document.getElementById('transactionsZone');
        const transactionsInfo = document.getElementById('transactionsInfo');
        
        
        if (transactionsZone && transactionsInfo) {
            if (dbInfo.hasTransactions) {
                transactionsZone.classList.add('file-selected');
                
                // Hide the instruction text when showing cached data
                const instructionP = transactionsZone.querySelector('p');
                if (instructionP) {
                    instructionP.style.display = 'none';
                }
                
                // Show enhanced transaction file information
                await this.displayEnhancedTransactionInfo(transactionsInfo);
                transactionsInfo.classList.add('show');
            } else {
                transactionsZone.classList.remove('file-selected');
                
                // Show the instruction text when no cached data
                const instructionP = transactionsZone.querySelector('p');
                if (instructionP) {
                    instructionP.style.display = 'block';
                }
                
                transactionsInfo.innerHTML = '';
                transactionsInfo.classList.remove('show');
            }
        } else {
        }
    }

    /**
     * Display enhanced portfolio file information with detected details
     */
    async displayEnhancedPortfolioInfo(infoElement) {
        try {
            const cachedData = await equateDB.getPortfolioData();
            
            if (!cachedData) {
                infoElement.innerHTML = 'Cached Portfolio Data available';
                return;
            }

            // Reuse the same logic as the modal analysis summary
            let portfolioCompany = 'Unknown';
            try {
                if (cachedData.portfolioData) {
                    // Use cached company data instead of re-detecting
                    portfolioCompany = cachedData.company || 'Unknown';
                }
            } catch (error) {
            }

            // Get detected info from cached metadata (more reliable than re-analyzing)
            const storedDetectedInfo = cachedData.portfolioData?.detectedInfo || {};
            
            const detectedInfo = {
                portfolioCompany: portfolioCompany,
                portfolioCurrency: storedDetectedInfo.currency || 'UNKNOWN',
                portfolioLanguage: storedDetectedInfo.language || 'English', 
                portfolioDateFormat: storedDetectedInfo.dateFormat || 'DD-MM-YYYY',
                entryCount: cachedData.portfolioData?.entries?.length || 0,
                asOfDate: cachedData.portfolioData?.asOfDate || cachedData.uploadDate?.split('T')[0]
            };
            
            
            // V2 BINARY ARCHITECTURE: No conversion strategy display needed
            // Multi-currency switching is instant, synthetic mode uses original currency
            let conversionStrategyHtml = '';
            
            const htmlContent = `
                <div class="file-details">
                    <div class="file-status">✅ ${storedDetectedInfo.originalFilename || 'Portfolio Details'} (Cached)</div>
                    <div class="file-meta">
                        <div><strong>Company:</strong> <span style="color: #2c3e50;">${detectedInfo.portfolioCompany}</span></div>
                        <div><strong>Currency:</strong> <span style="color: #e74c3c;">${detectedInfo.portfolioCurrency}</span></div>
                        ${conversionStrategyHtml}
                        <div><strong>Language:</strong> <span style="color: #8e44ad;">${this.getLanguageDisplayName(detectedInfo.portfolioLanguage)}</span></div>
                        <div><strong>Date Format:</strong> <span style="color: #3498db;">${detectedInfo.portfolioDateFormat}</span></div>
                        ${detectedInfo.asOfDate ? `<div><strong>As of Date:</strong> <span style="color: #f39c12;">${detectedInfo.asOfDate}</span></div>` : ''}
                        <div><strong>Investments:</strong> <span style="color: #27ae60;">${detectedInfo.entryCount}</span></div>
                    </div>
                </div>
            `;
            
            infoElement.innerHTML = htmlContent;
        } catch (error) {
            config.error('Error displaying portfolio info:', error);
            infoElement.innerHTML = '<div class="file-status">✅ Cached Portfolio Data available</div>';
        }
    }

    /**
     * Display enhanced transaction file information with detected details
     */
    async displayEnhancedTransactionInfo(infoElement) {
        try {
            const transactionData = await equateDB.getTransactionData();
            if (!transactionData) {
                infoElement.innerHTML = 'Cached Transaction Data available';
                return;
            }

            const detectedInfo = transactionData.detectedInfo || {};
            const stats = transactionData.stats || {};
            
            infoElement.innerHTML = `
                <div class="file-details">
                    <div class="file-status">✅ ${detectedInfo.originalFilename || 'Transaction History'} (Cached)</div>
                    <div class="file-meta">
                        ${detectedInfo.company ? `<div><strong>Company:</strong> <span style="color: #2c3e50;">${detectedInfo.company}</span></div>` : ''}
                        ${detectedInfo.currency ? `<div><strong>Currency:</strong> <span style="color: #e74c3c;">${detectedInfo.currency}</span></div>` : ''}
                        ${detectedInfo.language ? `<div><strong>Language:</strong> <span style="color: #8e44ad;">${this.getLanguageDisplayName(detectedInfo.language)}</span></div>` : ''}
                        ${detectedInfo.dateFormat ? `<div><strong>Date Format:</strong> <span style="color: #3498db;">${detectedInfo.dateFormat}</span></div>` : ''}
                        ${detectedInfo.asOfDate ? `<div><strong>As of Date:</strong> <span style="color: #f39c12;">${detectedInfo.asOfDate}</span></div>` : ''}
                        ${stats.entryCount ? `<div><strong>Transactions:</strong> <span style="color: #27ae60;">${stats.entryCount}</span></div>` : ''}
                    </div>
                </div>
            `;
        } catch (error) {
            config.error('Error displaying transaction info:', error);
            infoElement.innerHTML = 'Cached Transaction Data available';
        }
    }

    /**
     * Get display name for language code
     */
    getLanguageDisplayName(languageCode) {
        const languageMap = {
            'english': 'English',
            'german': 'Deutsch', 
            'dutch': 'Nederlands',
            'french': 'Français',
            'spanish': 'Español',
            'italian': 'Italiano',
            'polish': 'Polski',
            'turkish': 'Türkçe',
            'portuguese': 'Português',
            'czech': 'Čeština',
            'romanian': 'Română',
            'croatian': 'Hrvatski',
            'indonesian': 'Bahasa Indonesia',
            'chinese': '中文'
        };
        return languageMap[languageCode] || languageCode;
    }


    /**
     * Show Data Analysis Summary modal with detected information and warnings
     * Returns a promise that resolves to true (proceed) or false (cancel)
     */
    async showDataAnalysisSummary() {
        return new Promise(async (resolve) => {
            let isResolved = false;
            
            // Get detected information
            const analysisData = await this.getAnalysisSummaryData();
            
            // Create modal content
            const modalContent = `
                <div style="
                    position: fixed;
                    top: 0;
                    left: 0;
                    width: 100vw;
                    height: 100vh;
                    background: rgba(0, 0, 0, 0.7);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 10000;
                    font-family: system-ui, -apple-system, sans-serif;
                ">
                    <div style="
                        background: white;
                        padding: 30px;
                        border-radius: 12px;
                        max-width: 600px;
                        max-height: 80vh;
                        overflow-y: auto;
                        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3);
                        color: #333;
                    ">
                        <h2 style="margin: 0 0 20px 0; color: #2c3e50; display: flex; align-items: center; gap: 10px;">
                            📊 Data Analysis Summary
                        </h2>
                        
                        ${analysisData.detectionSection}
                        ${analysisData.warningsSection}
                        
                        <div style="
                            display: flex; 
                            gap: 15px; 
                            justify-content: center; 
                            margin-top: 25px;
                            border-top: 1px solid #eee;
                            padding-top: 20px;
                        ">
                            <button class="cancel-btn" style="
                                padding: 12px 24px;
                                background: #6c757d;
                                color: white;
                                border: none;
                                border-radius: 6px;
                                cursor: pointer;
                                font-size: 14px;
                                min-width: 100px;
                            ">Cancel</button>
                            <button class="proceed-btn" style="
                                padding: 12px 24px;
                                background: #007bff;
                                color: white;
                                border: none;
                                border-radius: 6px;
                                cursor: pointer;
                                font-size: 14px;
                                min-width: 140px;
                                font-weight: 600;
                            ">Generate Analysis</button>
                        </div>
                    </div>
                </div>
            `;
            
            const modalElement = document.createElement('div');
            modalElement.innerHTML = modalContent;
            document.body.appendChild(modalElement);
            
            const handleClick = (event) => {
                if (isResolved) return;
                
                if (event.target.classList.contains('proceed-btn')) {
                    isResolved = true;
                    document.body.removeChild(modalElement);
                    resolve(true);
                } else if (event.target.classList.contains('cancel-btn')) {
                    isResolved = true;
                    document.body.removeChild(modalElement);
                    resolve(false);
                }
            };
            
            modalElement.addEventListener('click', handleClick);
        });
    }

    /**
     * Get analysis summary data including detected info and warnings
     */
    async getAnalysisSummaryData() {
        // Use stored analysis result from generateAnalysis() - no duplicate work
        const portfolioFile = this.storedPortfolioFile || this.elements.portfolioInput.files[0];
        const transactionFile = this.storedTransactionFile || this.elements.transactionsInput.files[0];
        const cachedData = this.storedCachedData || await equateDB.getPortfolioData();
        
        // Use pre-computed analysis result (performance optimization)
        const currentCurrency = this.currencyService ? this.currencyService.getDetectedCurrency() : this.detectedCurrency;
        const analysisResult = this.storedAnalysisResult || {
            portfolio: cachedData ? { currency: currentCurrency || 'Unknown', language: 'English' } : null,
            transaction: cachedData?.transactionData?.entries?.length > 0 ? { currency: currentCurrency || 'Unknown', language: 'English' } : null,
            unified: cachedData ? { currency: currentCurrency || 'Unknown', language: 'English' } : null,
            warnings: [],
            errors: []
        };
        
        // Get individual per-file companies - use same source as dropzone (cached data)
        let portfolioCompany = 'Unknown';
        let transactionCompany = 'Not uploaded';
        
        // Priority 1: Use fresh analysis results if available (for new uploads)
        if (this.storedAnalysisResult?.transaction?.company && transactionFile) {
            transactionCompany = this.storedAnalysisResult.transaction.company;
            config.debug('📊 Using fresh analysis result for transaction company:', transactionCompany);
        }
        
        // Priority 2: Fall back to cached data for company info 
        if (cachedData) {
            portfolioCompany = cachedData.company || 'Unknown';
            // Only use cached transaction company if we don't have fresh results
            if (transactionCompany === 'Not uploaded') {
                if (cachedData.transactionData?.detectedInfo?.company) {
                    transactionCompany = cachedData.transactionData.detectedInfo.company;
                } else if (cachedData.transactionData) {
                    // If we have transaction data but no company info, assume it exists
                    transactionCompany = 'Uploaded (cached)';
                }
            }
        }
        
        // Override transaction file status if we have cached transaction data
        if (cachedData?.transactionData && !transactionFile) {
            // During regeneration, no files are uploaded but we have cached transaction data
            transactionCompany = cachedData.transactionData.detectedInfo?.company || 'Uploaded (cached)';
        }
        
        if (cachedData) {
            config.debug('📊 Using cached company data (same as dropzone):', {
                portfolioCompany,
                transactionCompany
            });
        } else if (this.storedAnalysisResult) {
            // Fallback to stored analysis result if no cached data
            portfolioCompany = this.storedAnalysisResult.portfolio?.company || 'Unknown';
            transactionCompany = this.storedAnalysisResult.transaction?.company || 'Not uploaded';
            
            config.debug('📊 Fallback to stored analysis company results:', {
                portfolioCompany,
                transactionCompany
            });
        }
        
        const detectedInfo = {
            portfolioCompany: portfolioCompany,
            transactionCompany: transactionCompany,
            portfolioCurrency: analysisResult.portfolio?.currency || 'Unknown',
            transactionCurrency: analysisResult.transaction?.currency || (cachedData?.transactionData?.entries?.length > 0 ? (currentCurrency || 'Unknown') : 'Not uploaded'),
            portfolioLanguage: analysisResult.portfolio?.language || 'Unknown', 
            transactionLanguage: analysisResult.transaction?.language || (cachedData?.transactionData?.entries?.length > 0 ? 'English' : 'Not uploaded'),
            portfolioDateFormat: fileParser.detectedDateFormat || 'DD-MM-YYYY',
            transactionDateFormat: (transactionFile || cachedData?.transactionData?.entries?.length > 0) ? (fileParser.detectedDateFormat || 'DD-MM-YYYY') : 'Not uploaded'
        };

        // Build warnings from FileAnalyzer results and other conditions
        const warnings = [];
        
        // Enhanced Currency conversion strategy warning - only for poor quality currencies
        if (detectedInfo.portfolioCurrency !== 'EUR') {
            try {
                const qualityScores = await equateDB.getCurrencyQualityScores();
                const currencyScore = qualityScores[detectedInfo.portfolioCurrency]?.score || 0;
                
                // Only show interpolation warning for currencies with poor coverage (< 90%)
                if (currencyScore < 90) {
                    warnings.push({
                        icon: '⚠️',
                        title: 'Interpolated Currency Conversion',
                        message: `${detectedInfo.portfolioCurrency} timeline uses interpolated exchange rates calculated from your portfolio purchase and transaction dates. Consider using EUR format for maximum accuracy.`,
                        severity: 'warning' // Orange styling
                    });
                }
            } catch (error) {
                // Fallback: show warning if we can't check quality scores
                config.warn('Could not check currency quality scores:', error);
                warnings.push({
                    icon: '⚠️',
                    title: 'Currency Conversion Notice',
                    message: `${detectedInfo.portfolioCurrency} timeline uses available historical exchange rates. Consider using EUR format for maximum accuracy.`,
                    severity: 'info' // Blue styling for less concerning
                });
            }
        }

        // Language differences warning
        if (detectedInfo.transactionLanguage !== 'Not uploaded' && detectedInfo.portfolioLanguage !== detectedInfo.transactionLanguage) {
            warnings.push({
                icon: '🌐',
                title: 'Language Difference',
                message: `Portfolio file language (${detectedInfo.portfolioLanguage}) differs from transaction file language (${detectedInfo.transactionLanguage}). Using portfolio file language.`,
                severity: 'info' // Green styling
            });
        }

        // Missing transaction file warning  
        if (detectedInfo.transactionCurrency === 'Not uploaded') {
            warnings.push({
                icon: '📋',
                title: 'Missing Transaction History',
                message: 'Total sold amount not available without CompletedTransactions file. Upload transaction history for complete analysis.',
                severity: 'info' // Green styling
            });
        }

        // Currency mismatch warning (if we had different currencies)
        // This would be handled by FileAnalyzer validation

        const detectionSection = `
            <div style="margin-bottom: 20px;">
                <h3 style="color: #2c3e50; margin: 0 0 15px 0; font-size: 16px;">🔍 Detected Information</h3>
                <div style="background: #f8f9fa; padding: 15px; border-radius: 8px;">
                    <div style="display: grid; gap: 12px;">
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; background: #ffffff; padding: 10px; border-radius: 4px; border: 1px solid #e0e0e0;">
                            <div style="text-align: center;">
                                <div style="font-weight: 600; color: #2c3e50; margin-bottom: 8px;">📊 Portfolio File</div>
                                <div><strong>Company:</strong> ${detectedInfo.portfolioCompany}</div>
                                <div><strong>Currency:</strong> ${detectedInfo.portfolioCurrency}</div>
                                <div><strong>Language:</strong> ${this.getLanguageDisplayName(detectedInfo.portfolioLanguage)}</div>
                                <div><strong>Date Format:</strong> ${detectedInfo.portfolioDateFormat}</div>
                            </div>
                            <div style="text-align: center;">
                                <div style="font-weight: 600; color: #2c3e50; margin-bottom: 8px;">📋 Transaction File</div>
                                <div><strong>Company:</strong> ${detectedInfo.transactionCompany}</div>
                                <div><strong>Currency:</strong> ${detectedInfo.transactionCurrency}</div>
                                <div><strong>Language:</strong> ${this.getLanguageDisplayName(detectedInfo.transactionLanguage)}</div>
                                <div><strong>Date Format:</strong> ${detectedInfo.transactionDateFormat}</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;

        // Helper function to get styling based on severity
        const getSeverityStyles = (severity) => {
            switch (severity) {
                case 'error':
                    return {
                        background: '#ffebee',
                        borderColor: '#f44336',
                        textColor: '#c62828'
                    };
                case 'warning':
                    return {
                        background: '#fff8e1',
                        borderColor: '#ff9800',
                        textColor: '#e65100'
                    };
                case 'info':
                default:
                    return {
                        background: '#e8f5e8',
                        borderColor: '#4caf50',
                        textColor: '#2e7d32'
                    };
            }
        };

        const warningsSection = warnings.length > 0 ? `
            <div style="margin-bottom: 20px;">
                <h3 style="color: #2c3e50; margin: 0 0 15px 0; font-size: 16px;">ℹ️ Information & Notices</h3>
                <div style="display: flex; flex-direction: column; gap: 12px;">
                    ${warnings.map(w => {
                        const styles = getSeverityStyles(w.severity);
                        return `
                        <div style="
                            background: ${styles.background}; 
                            padding: 12px; 
                            border-radius: 6px; 
                            border-left: 4px solid ${styles.borderColor};
                            display: flex;
                            align-items: flex-start;
                            gap: 10px;
                        ">
                            <span style="font-size: 16px;">${w.icon}</span>
                            <div>
                                <div style="font-weight: 600; margin-bottom: 4px; color: ${styles.textColor};">${w.title}</div>
                                <div style="font-size: 13px; color: #555;">${w.message}</div>
                            </div>
                        </div>
                        `;
                    }).join('')}
                </div>
            </div>
        ` : '';

        return { detectionSection, warningsSection };
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


    /**
     * Create synthetic timeline directly from user transaction data
     * CRITICAL: No currency conversion - user data is already in correct currency
     * @param {Object} portfolioData - Portfolio data with reference points
     * @param {Object} transactionData - Transaction data with reference points  
     * @param {String} currency - Detected currency (can be null if detection failed)
     * @returns {Array} Synthetic timeline data in user's original currency
     */
    async createSyntheticTimelineFromUserData(portfolioData, transactionData, currency) {
        // Use ReferencePointsBuilder to get comprehensive reference points
        const enhancedReferencePoints = await this.referencePointsBuilder.buildReferencePoints(
            portfolioData, 
            transactionData
        );
        
        if (enhancedReferencePoints.length === 0) {
            config.warn('No reference points found for synthetic timeline');
            return [];
        }
        
        // Extract simple date/price pairs for flat-line timeline generation
        const timelinePoints = enhancedReferencePoints.map(point => ({
            date: point.date,
            price: point.currentPrice || point.prices[point.originalCurrency]
        }));
        
        // Create flat-line timeline between enhanced reference points  
        const timeline = this.generateFlatLineTimeline(timelinePoints);
        
        config.debug('📊 Created synthetic timeline from user data:', {
            currency: currency || 'unknown',
            enhancedReferencePoints: timelinePoints.length,
            timelinePoints: timeline.length,
            firstPoint: timelinePoints[0],
            lastPoint: timelinePoints[timelinePoints.length - 1]
        });
        
        return timeline;
    }


    /**
     * Generate flat-line timeline between timeline points (derived from enhanced reference points)
     * @param {Array} timelinePoints - Sorted timeline points with date/price
     * @returns {Array} Timeline with daily price points
     */
    generateFlatLineTimeline(timelinePoints) {
        if (timelinePoints.length === 0) return [];
        
        const timeline = [];
        const startDate = new Date(timelinePoints[0].date);
        const endDate = new Date(timelinePoints[timelinePoints.length - 1].date);
        
        // Generate daily timeline
        let currentDate = new Date(startDate);
        let currentPriceIndex = 0;
        
        while (currentDate <= endDate) {
            const dateStr = currentDate.toISOString().split('T')[0];
            
            // Find the appropriate price for this date (flat line approach)
            while (currentPriceIndex < timelinePoints.length - 1 && 
                   dateStr >= timelinePoints[currentPriceIndex + 1].date) {
                currentPriceIndex++;
            }
            
            timeline.push({
                date: dateStr,
                price: timelinePoints[currentPriceIndex].price
            });
            
            currentDate.setUTCDate(currentDate.getUTCDate() + 1);
        }
        
        return timeline;
    }

    /**
     * Check if transaction is a sell transaction
     */
    isSellTransaction(transaction) {
        return transaction.status === 'Executed' && (
            transaction.orderType === 'Sell' || 
            transaction.orderType === 'Sell at market price' || 
            transaction.orderType === 'Sell with price limit' ||
            transaction.orderType === 'Transfer'
        );
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
    // Track tab switch event
    if (typeof gtag !== 'undefined') {
        gtag('event', 'tab_switch', {
            'tab_name': tabName,
            'event_category': 'navigation'
        });
    }
    
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

    // Handle tab-specific initialization
    if (tabName === 'calculations') {
        // Initialize and render calculations tab
        if (portfolioCalculator.calculations) {
            calculationsTab.toggleCalculationsDisplay(true);
            calculationsTab.render(portfolioCalculator.calculations);
        } else {
            calculationsTab.toggleCalculationsDisplay(false);
        }
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
    
    // Also toggle calculations tab display
    if (typeof calculationsTab !== 'undefined') {
        calculationsTab.toggleCalculationsDisplay(hasData);
    }
}

function switchTheme(theme) {
    app.switchTheme(theme);
}

function setNumberFormat(format) {
    try {
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
        // Track manual price update event
        if (typeof gtag !== 'undefined') {
            gtag('event', 'manual_price_update', {
                'price': parseFloat(manualPriceInput.value),
                'event_category': 'user_interaction'
            });
        }
        
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


function exportPDF() {
    if (app.currentCalculations && app.hasData) {
        // Track PDF export event
        if (typeof gtag !== 'undefined') {
            gtag('event', 'pdf_export', {
                'event_category': 'export',
                'export_type': 'pdf'
            });
        }
        
        portfolioExporter.setData(app.currentCalculations, portfolioCalculator.portfolioData);
        portfolioExporter.exportPDF(); // Use actual PDF export with chart capture
    }
}

function showExportCustomizer() {
    if (app.currentCalculations && app.hasData) {
        portfolioExporter.setData(app.currentCalculations, portfolioCalculator.portfolioData);
        if (window.exportCustomizer) {
            window.exportCustomizer.show();
        } else {
            console.error('Export customizer not available');
        }
    }
}


function clearCachedData() {
    if (confirm('Are you sure you want to clear all cached portfolio data? You will need to re-upload your files.')) {
        app.clearAllData();
    }
}

// Cache Inspector Functions
async function openCacheInspector() {
    if (!window.CacheManager || !window.CacheManager.isAvailable()) {
        alert('Cache Inspector is only available on localhost for development purposes.');
        return;
    }
    
    const modal = document.getElementById('cacheInspectorModal');
    const statsContainer = document.getElementById('cacheStats');
    const entriesContainer = document.getElementById('cacheEntries');
    
    // Get cache data
    const stats = window.CacheManager.getStats();
    const entries = window.CacheManager.getAllEntries();
    
    // Get complete cache overview
    const dbInfo = await equateDB.getDatabaseInfo();
    const currencyData = await equateDB.getHistoricalCurrencyData();
    const qualityScores = await equateDB.getCurrencyQualityScores();
    const multiCurrencyData = await equateDB.getMultiCurrencyPrices();
    
    // Get all IndexedDB store counts
    const allStoreCounts = await getAllIndexedDBStoreCounts();
    
    // Get localStorage data
    const localStorageData = getLocalStorageData();
    
    // Calculate cache sizes
    const cacheSizes = await calculateCacheSizes();
    
    // Populate comprehensive stats
    statsContainer.innerHTML = `
        <div style="background: linear-gradient(135deg, #2a2a2a, #1a1a1a); padding: 20px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #444;">
            <h4 style="margin: 0 0 10px 0; color: #fff; font-size: 18px;">💾 Cache Overview</h4>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin-top: 10px;">
                <div style="display: flex; justify-content: space-between; padding: 12px; background: #333; border-radius: 8px; border: 1px solid ${
                    cacheSizes.totalMB > 50 ? '#e74c3c' : cacheSizes.totalMB > 10 ? '#f39c12' : '#27ae60'
                };">
                    <span style="color: #ccc;">Total Size:</span>
                    <span><strong style="color: #fff;">${cacheSizes.totalMB}MB</strong></span>
                </div>
                <div style="display: flex; justify-content: space-between; padding: 12px; background: #333; border-radius: 8px; border: 1px solid #444;">
                    <span style="color: #ccc;">Cache Systems:</span>
                    <span><strong style="color: #fff;">4 Active</strong></span>
                </div>
            </div>
        </div>
        
        <div style="background: #2a2a2a; padding: 15px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #444;">
            <h4 style="margin: 0 0 12px 0; color: #fff; font-size: 16px;">🧠 Session Cache (Development)</h4>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid ${cacheSizes.sessionMB > 20 ? '#e74c3c' : '#27ae60'};">
                <span>Entries:</span>
                <span><strong style="color: #fff;">${stats.totalEntries} (${cacheSizes.sessionMB}MB)</strong></span>
            </div>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid #444;">
                <span>Keys:</span>
                <span><strong style="color: #fff;">${stats.keys.join(', ') || 'None'}</strong></span>
            </div>
            <button onclick="clearSessionCache()" style="background: #007acc; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; margin-top: 8px;" ${stats.totalEntries === 0 ? 'disabled style="background: #666; color: #999; cursor: not-allowed;"' : ''}>Clear Session</button>
        </div>
        
        <div style="background: #2a2a2a; padding: 15px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #444;">
            <h4 style="margin: 0 0 12px 0; color: #fff; font-size: 16px;">🗄️ IndexedDB Persistent Storage</h4>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid #444;">
                <span>Portfolio Data:</span>
                <span><strong style="color: #fff;">${dbInfo.hasPortfolioData ? '✅ Yes' : '❌ No'} (${cacheSizes.portfolioMB}MB)</strong></span>
            </div>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid #444;">
                <span>Historical Prices:</span>
                <span><strong style="color: #fff;">${allStoreCounts.historicalPrices || 0} entries (${cacheSizes.historicalMB}MB)</strong></span>
            </div>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid ${multiCurrencyData.length > 0 ? '#27ae60' : '#666'};">
                <span>MultiCurrency Prices:</span>
                <span><strong style="color: #fff;">${multiCurrencyData.length || 0} entries ${multiCurrencyData.length > 0 ? '⚡ INSTANT' : ''}</strong></span>
            </div>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid #444;">
                <span>Currency Data:</span>
                <span><strong style="color: #fff;">${currencyData.length || 0} entries (${cacheSizes.currencyMB}MB)</strong></span>
            </div>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid ${Object.keys(qualityScores).length === 0 ? '#e74c3c' : '#27ae60'};">
                <span>Quality Scores:</span>
                <span><strong style="color: #fff;">${Object.keys(qualityScores).length || 0} currencies</strong></span>
            </div>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid #444;">
                <span>All DB Stores:</span>
                <span><strong style="color: #fff;">${Object.keys(allStoreCounts).length} stores</strong></span>
            </div>
            <button onclick="clearIndexedDB()" style="background: #e74c3c; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; margin-top: 8px;" ${!dbInfo.hasPortfolioData ? 'disabled style="background: #666; color: #999; cursor: not-allowed;"' : ''}>Clear Database</button>
        </div>
        
        <div style="background: #2a2a2a; padding: 15px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #444;">
            <h4 style="margin: 0 0 12px 0; color: #fff; font-size: 16px;">🏠 localStorage Preferences</h4>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid #444;">
                <span>Items:</span>
                <span><strong style="color: #fff;">${localStorageData.count} (${cacheSizes.localStorageMB}MB)</strong></span>
            </div>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid #444;">
                <span>Export Prefs:</span>
                <span><strong style="color: #fff;">${localStorageData.hasExportPrefs ? '✅ Yes' : '❌ No'}</strong></span>
            </div>
            <button onclick="clearLocalStorageCache()" style="background: #9b59b6; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; margin-top: 8px;" ${localStorageData.count === 0 ? 'disabled style="background: #666; color: #999; cursor: not-allowed;"' : ''}>Clear Preferences</button>
        </div>
        
        <div style="background: #2a2a2a; padding: 15px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #444; border-left: 4px solid #e74c3c;">
            <h4 style="margin: 0 0 12px 0; color: #e74c3c; font-size: 16px;">☢️ Nuclear Option</h4>
            <div style="display: flex; justify-content: space-between; margin: 5px 0; color: #ccc; font-size: 14px; padding: 8px 12px; border-radius: 6px; background: #1a1a1a; border: 1px solid #444;">
                <span>Total Data:</span>
                <span><strong style="color: #fff;">${
                    stats.totalEntries + Object.keys(allStoreCounts).reduce((sum, key) => sum + (allStoreCounts[key] || 0), 0) + localStorageData.count
                } items across all systems</strong></span>
            </div>
            <button onclick="clearEverythingConfirm()" style="background: #e74c3c; color: white; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; font-size: 12px; margin-top: 8px; animation: pulse 2s infinite;">🚨 Clear EVERYTHING</button>
        </div>
    `;
    
    // Build detailed entries for exploration
    const allEntries = [];
    
    // Add CacheManager entries with enhanced detail
    Object.entries(entries).forEach(([key, entry]) => {
        allEntries.push({
            key: key,
            source: 'Session Cache',
            entry: entry,
            canExpand: entry.length > 5,
            displayData: entry.sampleData || entry.data
        });
    });
    
    // Add comprehensive IndexedDB entries
    if (currencyData.length > 0) {
        const sampleCurrencies = currencyData[0] ? Object.keys(currencyData[0]).filter(k => k !== 'date').slice(0, 5) : [];
        allEntries.push({
            key: 'Historical Currency Data',
            source: 'IndexedDB',
            entry: {
                dataType: 'Exchange Rate History',
                length: currencyData.length,
                description: `Exchange rates for ${sampleCurrencies.length}+ currencies from ${currencyData[0]?.date} to ${currencyData[currencyData.length-1]?.date}`,
                size: cacheSizes.currencyMB + 'MB'
            },
            canExpand: true,
            displayData: {
                totalEntries: currencyData.length,
                availableCurrencies: sampleCurrencies,
                dateRange: currencyData.length > 0 ? `${currencyData[0]?.date} to ${currencyData[currencyData.length-1]?.date}` : 'No data',
                sampleEntry: currencyData[0] || 'No data'
            }
        });
    }
    
    // Add MultiCurrency Prices entries (for instant switching)
    if (multiCurrencyData.length > 0) {
        const sampleCurrencies = multiCurrencyData[0] ? Object.keys(multiCurrencyData[0]).filter(k => k !== 'date').slice(0, 8) : [];
        const firstFive = multiCurrencyData.slice(0, 5);
        allEntries.push({
            key: 'MultiCurrency Prices ⚡ INSTANT',
            source: 'IndexedDB',
            entry: {
                dataType: 'Pre-computed Currency Prices',
                length: multiCurrencyData.length,
                description: `Share prices in ${sampleCurrencies.length}+ currencies for instant switching (<10ms)`,
                size: 'Multi-MB'
            },
            canExpand: true,
            displayData: {
                totalEntries: multiCurrencyData.length,
                availableCurrencies: sampleCurrencies,
                dateRange: multiCurrencyData.length > 0 ? `${multiCurrencyData[0]?.date} to ${multiCurrencyData[multiCurrencyData.length-1]?.date}` : 'No data',
                firstFiveEntries: firstFive,
                sampleEntry: multiCurrencyData[0] || 'No data'
            }
        });
    }
    
    if (Object.keys(qualityScores).length > 0) {
        const sampleQualityScores = Object.keys(qualityScores).slice(0, 5).reduce((obj, key) => {
            obj[key] = qualityScores[key];
            return obj;
        }, {});
        
        allEntries.push({
            key: 'Currency Quality Scores',
            source: 'IndexedDB',
            entry: {
                dataType: 'Detection Confidence Data',
                length: Object.keys(qualityScores).length,
                description: `Quality scores for ${Object.keys(qualityScores).length} currencies based on detection confidence`,
                size: '< 1MB'
            },
            canExpand: true,
            displayData: {
                totalCurrencies: Object.keys(qualityScores).length,
                sampleScores: sampleQualityScores,
                fullData: qualityScores
            }
        });
    }
    
    // Add Portfolio Data if available
    if (dbInfo.hasPortfolioData) {
        // Get the actual cached portfolio data
        const cachedPortfolioData = await equateDB.getPortfolioData();
        allEntries.push({
            key: 'Portfolio Data',
            source: 'IndexedDB',
            entry: {
                dataType: 'Portfolio Analysis Data',
                length: 1,
                description: 'Main portfolio and transaction data with metadata',
                size: cacheSizes.portfolioMB + 'MB'
            },
            canExpand: true,
            displayData: {
                hasPortfolio: dbInfo.hasPortfolioData,
                hasTransactions: dbInfo.hasTransactions,
                lastUpload: dbInfo.lastUpload,
                userId: dbInfo.userId || 'Not set',
                portfolioData: cachedPortfolioData?.portfolioData || null,
                transactionData: cachedPortfolioData?.transactionData || null
            }
        });
    }
    
    // Add Historical Prices if available
    if (allStoreCounts.historicalPrices > 0) {
        allEntries.push({
            key: 'Historical Share Prices',
            source: 'IndexedDB',
            entry: {
                dataType: 'Shares Company Price History',
                length: allStoreCounts.historicalPrices,
                description: 'Historical share prices for Shares-companies timeline calculations',
                size: cacheSizes.historicalMB + 'MB'
            },
            canExpand: true,
            displayData: {
                priceCount: allStoreCounts.historicalPrices,
                source: 'hist.xlsx and portfolio AsOfDate',
                usage: 'Timeline chart generation'
            }
        });
    }
    
    // Add localStorage entries
    if (localStorageData.count > 0) {
        allEntries.push({
            key: 'Browser Preferences',
            source: 'localStorage',
            entry: {
                dataType: 'User Settings',
                length: localStorageData.count,
                description: 'Export preferences and UI settings',
                size: cacheSizes.localStorageMB + 'MB'
            },
            canExpand: true,
            displayData: localStorageData.items
        });
    }
    
    // Build simple, working entries display
    if (allEntries.length === 0) {
        entriesContainer.innerHTML = `
            <div style="text-align: center; padding: 20px; color: #ccc;">
                No cached data found. Upload portfolio files to see cached data.
            </div>
        `;
    } else {
        entriesContainer.innerHTML = allEntries.map(({key, source, entry, displayData}, index) => `
            <div style="background: #2a2a2a; border: 1px solid #444; border-radius: 8px; margin-bottom: 12px; overflow: hidden;">
                <div style="background: #333; padding: 12px; border-bottom: 1px solid #444; cursor: pointer;" onclick="toggleCacheData(${index})">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <div>
                            <span style="color: #fff; font-weight: bold; margin-right: 10px;">${key}</span>
                            <span style="background: #007acc; color: white; padding: 2px 8px; border-radius: 12px; font-size: 11px;">${source}</span>
                        </div>
                        <span id="toggle-icon-${index}" style="color: #fff; font-size: 16px;">▶</span>
                    </div>
                    <div style="color: #ccc; font-size: 12px; margin-top: 4px;">
                        ${entry.dataType} • ${entry.length || 0} items • ${entry.size || '< 1MB'}
                    </div>
                </div>
                <div id="cache-details-${index}" style="display: none; padding: 12px; background: #1a1a1a;">
                    ${entry.description ? `<p style="color: #aaa; font-style: italic; margin: 0 0 10px 0;">${entry.description}</p>` : ''}
                    <pre style="background: #000; color: #0f0; padding: 10px; border-radius: 4px; overflow-x: auto; font-size: 11px; margin: 0;">${JSON.stringify(displayData, null, 2)}</pre>
                </div>
            </div>
        `).join('');
    }
    
    // Show modal
    modal.style.display = 'flex';
    
    // Close modal when clicking outside
    modal.onclick = (e) => {
        if (e.target === modal) {
            closeCacheInspector();
        }
    };
}

function closeCacheInspector() {
    const modal = document.getElementById('cacheInspectorModal');
    modal.style.display = 'none';
}

// Helper functions for comprehensive cache inspection

async function getAllIndexedDBStoreCounts() {
    console.log('🔍 INSPECTOR DEBUG: getAllIndexedDBStoreCounts() called');
    const counts = {};
    const storeNames = [
        'portfolioData', 'historicalPrices', 'userPreferences', 'manualPrices',
        'historicalCurrencyData', 'currencyQualityScores'
    ];
    
    let totalItems = 0;
    
    for (const storeName of storeNames) {
        try {
            const transaction = equateDB.db.transaction([storeName], 'readonly');
            const store = transaction.objectStore(storeName);
            const count = await new Promise((resolve, reject) => {
                const request = store.count();
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => resolve(0);
            });
            counts[storeName] = count;
            totalItems += count;
            console.log(`🔍 INSPECTOR DEBUG: Store ${storeName} has ${count} items`);
        } catch (error) {
            counts[storeName] = 0;
            console.log(`🔍 INSPECTOR DEBUG: Error counting ${storeName}:`, error.message);
        }
    }
    
    console.log(`🔍 INSPECTOR DEBUG: Total items across all stores: ${totalItems}`);
    console.log(`🔍 INSPECTOR DEBUG: Store counts:`, counts);
    
    return counts;
}

function getLocalStorageData() {
    const data = {
        count: localStorage.length,
        items: {},
        hasExportPrefs: false
    };
    
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        const value = localStorage.getItem(key);
        data.items[key] = {
            size: new Blob([value]).size,
            preview: value.length > 100 ? value.substring(0, 100) + '...' : value
        };
        
        if (key.includes('export') || key.includes('section')) {
            data.hasExportPrefs = true;
        }
    }
    
    return data;
}

async function calculateCacheSizes() {
    const sizes = {
        sessionMB: 0,
        portfolioMB: 0,
        historicalMB: 0,
        currencyMB: 0,
        localStorageMB: 0,
        totalMB: 0
    };
    
    // Session cache size
    const sessionStats = window.CacheManager?.getStats();
    if (sessionStats) {
        const sessionEntries = window.CacheManager.getAllEntries();
        let sessionBytes = 0;
        Object.values(sessionEntries).forEach(entry => {
            try {
                sessionBytes += new Blob([JSON.stringify(entry.data)]).size;
            } catch (e) {
                sessionBytes += 1000; // Estimate for non-serializable data
            }
        });
        sizes.sessionMB = Math.round(sessionBytes / 1024 / 1024 * 100) / 100;
    }
    
    // IndexedDB sizes (estimates)
    try {
        const portfolioData = await equateDB.getPortfolioData();
        if (portfolioData) {
            sizes.portfolioMB = Math.round(new Blob([JSON.stringify(portfolioData)]).size / 1024 / 1024 * 100) / 100;
        }
        
        const historicalPrices = await equateDB.getHistoricalPrices();
        if (historicalPrices.length > 0) {
            sizes.historicalMB = Math.round(new Blob([JSON.stringify(historicalPrices)]).size / 1024 / 1024 * 100) / 100;
        }
        
        const currencyData = await equateDB.getHistoricalCurrencyData();
        if (currencyData.length > 0) {
            sizes.currencyMB = Math.round(new Blob([JSON.stringify(currencyData)]).size / 1024 / 1024 * 100) / 100;
        }
    } catch (error) {
        console.warn('Error calculating IndexedDB sizes:', error);
    }
    
    // localStorage size
    let localStorageBytes = 0;
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        const value = localStorage.getItem(key);
        localStorageBytes += new Blob([key + value]).size;
    }
    sizes.localStorageMB = Math.round(localStorageBytes / 1024 / 1024 * 100) / 100;
    
    // Total
    sizes.totalMB = Math.round((sizes.sessionMB + sizes.portfolioMB + sizes.historicalMB + sizes.currencyMB + sizes.localStorageMB) * 100) / 100;
    
    return sizes;
}

// Cache clearing functions

async function clearSessionCache() {
    if (!window.CacheManager) return;
    
    const stats = window.CacheManager.getStats();
    if (stats.totalEntries === 0) return;
    
    if (confirm(`Clear ${stats.totalEntries} session cache entries?`)) {
        window.CacheManager.clearAll();
        config.debug('🗑️ Session cache cleared');
        
        // Refresh the inspector
        setTimeout(() => openCacheInspector(), 100);
    }
}

async function clearIndexedDB() {
    const dbInfo = await equateDB.getDatabaseInfo();
    if (!dbInfo.hasPortfolioData) return;
    
    if (confirm('Clear all IndexedDB data (portfolio, historical data, preferences)?')) {
        try {
            await equateDB.clearAllData();
            config.debug('🗑️ IndexedDB cleared');
            
            // Update UI state
            const uploadStatusElement = document.getElementById('uploadStatus');
            if (uploadStatusElement) {
                uploadStatusElement.textContent = 'No cached data found';
            }
            
            // Refresh the inspector
            setTimeout(() => openCacheInspector(), 100);
        } catch (error) {
            alert('Error clearing IndexedDB: ' + error.message);
        }
    }
}

function clearLocalStorageCache() {
    const localData = getLocalStorageData();
    if (localData.count === 0) return;
    
    if (confirm(`Clear ${localData.count} localStorage items (export preferences, etc.)?`)) {
        localStorage.clear();
        config.debug('🗑️ localStorage cleared');
        
        // Refresh the inspector
        setTimeout(() => openCacheInspector(), 100);
    }
}

async function clearEverythingConfirm() {
    const stats = window.CacheManager?.getStats() || { totalEntries: 0 };
    const dbInfo = await equateDB.getDatabaseInfo();
    const localData = getLocalStorageData();
    const allStoreCounts = await getAllIndexedDBStoreCounts();
    
    const totalItems = stats.totalEntries + 
                      Object.values(allStoreCounts).reduce((sum, count) => sum + count, 0) + 
                      localData.count;
    
    if (totalItems === 0) {
        alert('No cached data to clear.');
        return;
    }
    
    const message = `⚠️ NUCLEAR OPTION ⚠️

This will permanently delete ALL cached data:

🧠 Session Cache: ${stats.totalEntries} items
🗃️ IndexedDB: ${Object.values(allStoreCounts).reduce((sum, count) => sum + count, 0)} items  
💾 localStorage: ${localData.count} items

Total: ${totalItems} items

This action CANNOT be undone. Continue?`;
    
    if (confirm(message)) {
        try {
            // Clear everything
            if (window.CacheManager) window.CacheManager.clearAll();
            await equateDB.clearAllData();
            
            // Clear the additional stores that clearAllData() misses
            const additionalStores = [
                'userPreferences', 'manualPrices', 'currencyQualityScores'
            ];
            
            console.log('🔍 NUCLEAR DEBUG: Clearing additional stores that clearAllData() missed');
            
            for (const storeName of additionalStores) {
                try {
                    console.log(`🔍 NUCLEAR DEBUG: Clearing additional store: ${storeName}`);
                    const transaction = equateDB.db.transaction([storeName], 'readwrite');
                    const store = transaction.objectStore(storeName);
                    
                    const countBefore = await new Promise((resolve) => {
                        const request = store.count();
                        request.onsuccess = () => resolve(request.result);
                        request.onerror = () => resolve(-1);
                    });
                    
                    await new Promise((resolve, reject) => {
                        const request = store.clear();
                        request.onsuccess = () => resolve();
                        request.onerror = () => reject(request.error);
                    });
                    
                    console.log(`🗑️ NUCLEAR DEBUG: Cleared ${storeName} (${countBefore} items)`);
                } catch (error) {
                    console.log(`⚠️ NUCLEAR DEBUG: Failed to clear ${storeName}:`, error.message);
                }
            }
            
            localStorage.clear();
            
            config.debug('☢️ ALL CACHE DATA CLEARED');
            alert('All cached data has been cleared successfully.');
            
            // Update UI state
            const uploadStatusElement = document.getElementById('uploadStatus');
            if (uploadStatusElement) {
                uploadStatusElement.textContent = 'No cached data found';
            }
            
            // Refresh the inspector to show cleared state
            setTimeout(() => openCacheInspector(), 100);
        } catch (error) {
            alert('Error during nuclear clear: ' + error.message);
        }
    }
}

function toggleCacheData(index) {
    const detailsElement = document.getElementById(`cache-details-${index}`);
    const iconElement = document.getElementById(`toggle-icon-${index}`);
    
    if (detailsElement && iconElement) {
        if (detailsElement.style.display === 'none') {
            detailsElement.style.display = 'block';
            iconElement.textContent = '▼';
        } else {
            detailsElement.style.display = 'none';
            iconElement.textContent = '▶';
        }
    }
}

function showFullCacheData(key) {
    if (!window.CacheManager || !window.CacheManager.isAvailable()) {
        return;
    }
    
    const fullData = window.CacheManager.getFullData(key);
    if (!fullData) {
        alert('Could not retrieve full data for: ' + key);
        return;
    }
    
    const dataElement = document.getElementById(`cache-data-${key}`);
    if (dataElement) {
        // Show confirmation for large datasets
        if (fullData.length > 1000) {
            if (!confirm(`This will display all ${fullData.length} entries. This may slow down the browser. Continue?`)) {
                return;
            }
        }
        
        dataElement.innerHTML = JSON.stringify(fullData, null, 2);
        dataElement.style.maxHeight = '600px';
        dataElement.style.fontSize = '10px';
        
        // Add a "Show Sample" button to revert
        const button = dataElement.parentElement.querySelector('.cache-show-full-btn');
        if (button) {
            button.textContent = 'Show Sample';
            button.onclick = () => showSampleCacheData(key);
        }
    }
}

function showSampleCacheData(key) {
    if (!window.CacheManager || !window.CacheManager.isAvailable()) {
        return;
    }
    
    const entry = window.CacheManager.getEntry(key);
    if (!entry) return;
    
    const dataElement = document.getElementById(`cache-data-${key}`);
    if (dataElement) {
        dataElement.innerHTML = JSON.stringify(entry.sampleData, null, 2);
        dataElement.style.maxHeight = '200px';
        dataElement.style.fontSize = '12px';
        
        // Revert button text
        const button = dataElement.parentElement.querySelector('.cache-show-full-btn');
        if (button) {
            button.textContent = `Show All ${entry.length}`;
            button.onclick = () => showFullCacheData(key);
        }
    }
}

// Initialize the application when DOM is loaded
const app = new EquateApp();
window.app = app;  // Make app globally accessible for CurrencyService

document.addEventListener('DOMContentLoaded', () => {
    app.init();
    
    // Show cache inspector on localhost only
    if (window.CacheManager && window.CacheManager.isAvailable()) {
        const cacheInspector = document.getElementById('cacheInspector');
        if (cacheInspector) {
            cacheInspector.style.display = 'block';
        }
    }
}); 