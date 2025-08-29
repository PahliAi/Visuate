/**
 * File Analyzer for Equate Portfolio Analysis
 * Handles efficient pre-parsing analysis of portfolio and transaction files as a unified set
 * Separates detection logic from parsing logic for clean architecture
 */

class FileAnalyzer {
    constructor() {
        // Initialize currency detector
        if (typeof CurrencyDetector !== 'undefined') {
            this.currencyDetector = new CurrencyDetector();
        } else {
            config.warn('⚠️ CurrencyDetector not available - using fallback method');
            this.currencyDetector = null;
        }
        
        // Initialize translation manager reference
        if (typeof translationManager !== 'undefined') {
            this.translationManager = translationManager;
        } else {
            config.warn('⚠️ TranslationManager not available');
            this.translationManager = null;
        }
    }

    /**
     * Analyze both files as a unified set with cache-aware compatibility validation
     * Handles all upload scenarios including sequential uploads with cached data
     * @param {File} portfolioFile - Required portfolio file (master for currency/language)
     * @param {File} transactionFile - Optional transaction file  
     * @param {Object} cachedData - Optional cached data from database
     * @returns {Object} Unified analysis result with compatibility validation and cache actions
     */
    async analyzeFileSet(portfolioFile, transactionFile = null, cachedData = null) {
        config.debug('🔍 Starting unified file set analysis...');
        
        const results = {
            portfolio: null,
            transaction: null,
            unified: null,
            warnings: [],
            errors: [],
            isCompatible: true,
            shouldClearCachedTransactions: false
        };

        try {
            // Step 1: Cache-aware validation for sequential uploads
            if (cachedData) {
                config.debug('🔍 Performing cache-aware validation...');
                const cacheValidation = await this.validateWithCache(portfolioFile, transactionFile, cachedData);
                
                // Merge cache validation results
                results.warnings.push(...cacheValidation.warnings);
                results.errors.push(...cacheValidation.errors);
                results.shouldClearCachedTransactions = cacheValidation.shouldClearCachedTransactions;
                
                // If transaction file was rejected due to currency mismatch, don't process it
                if (transactionFile && cacheValidation.errors.some(e => e.type === 'CURRENCY_MISMATCH')) {
                    config.debug('🚫 Transaction file rejected by cache validation - removing from processing');
                    transactionFile = null;
                }
            }

            // Step 2: Analyze portfolio file (if provided) or use cached data
            if (portfolioFile) {
                config.debug('📄 Analyzing portfolio file:', portfolioFile.name);
                results.portfolio = await this.analyzeFile(portfolioFile, 'portfolio');
            } else if (cachedData && cachedData.portfolioData) {
                // Use cached portfolio data as the master currency/language source
                config.debug('📄 Using cached portfolio data as master source');
                results.portfolio = {
                    fileName: 'cached-portfolio-data',
                    fileType: 'portfolio',
                    language: cachedData.portfolioData?.detectedInfo?.language || (cachedData.isEnglish ? 'english' : 'other'), // Use actual cached language if available
                    currency: cachedData.currency,
                    rawData: null, // Not needed for validation
                    worksheet: null // Not needed for validation
                };
            } else {
                throw new Error('No portfolio file provided and no cached portfolio data available');
            }
            
            // Step 3: Analyze transaction file if provided and not rejected
            if (transactionFile) {
                config.debug('📄 Analyzing transaction file:', transactionFile.name);
                results.transaction = await this.analyzeFile(transactionFile, 'transaction');
            }

            // Step 4: Validate compatibility and create unified result
            results.unified = this.validateAndUnify(results);
            
            config.debug('✅ File set analysis complete:', {
                portfolio: results.portfolio.fileName,
                transaction: results.transaction?.fileName || 'none',
                language: results.unified.language,
                currency: results.unified.currency,
                company: results.unified.company,
                warnings: results.warnings.length,
                errors: results.errors.length
            });

            return results;
            
        } catch (error) {
            config.error('❌ File set analysis failed:', error);
            results.errors.push({
                type: 'ANALYSIS_FAILED',
                message: `File analysis failed: ${error.message}`,
                resolution: 'Please check file format and try again.'
            });
            results.isCompatible = false;
            return results;
        }
    }

    /**
     * Analyze single file efficiently (language, currency, company indicators)
     * @param {File} file - Excel file to analyze
     * @param {String} type - 'portfolio' or 'transaction'
     * @returns {Object} Analysis result for single file
     */
    async analyzeFile(file, type) {
        const analysis = {
            fileName: file.name,
            fileType: type,
            language: null,
            currency: null,
            company: null, // Will be determined after parsing both files
            rawData: null,
            worksheet: null
        };

        try {
            // Single read of Excel file
            const data = await this.readExcelFile(file);
            const worksheet = data.Sheets[data.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                raw: true
            });

            analysis.rawData = rawData;
            analysis.worksheet = worksheet;

            // Detect language using translation manager
            if (this.translationManager) {
                analysis.language = await this.translationManager.detectLanguageFromFile(file);
                config.debug(`🌐 Language detected for ${file.name}:`, analysis.language);
            } else {
                analysis.language = 'english'; // Fallback
                config.warn('⚠️ No translation manager - defaulting to English');
            }

            // Detect currency from transaction price columns only (columns 5 & 6: costBasis, marketPrice)
            if (this.currencyDetector && type === 'portfolio') {
                // Use full data + worksheet to preserve Excel formatting that contains currency codes
                // Currency will still be detected from price columns (5 & 6) but via Excel formatting
                const currencyResult = await this.currencyDetector.detectCurrency(rawData, worksheet, {
                    interactive: false,
                    showDialogs: false
                });
                
                const detectedCurrency = currencyResult.currencyCode || currencyResult.primaryCurrency;
                
                if (!detectedCurrency) {
                    throw new Error(`❌ Currency detection failed for ${file.name}. No currency found in Excel file. Please ensure your Excel file contains currency codes or symbols in the price columns.`);
                }
                
                // Check for multiple currencies in single file (unhappy flow #2, #8, #9)
                if (currencyResult.isMixed || currencyResult.allCurrencies.size > 1) {
                    const currencies = Array.from(currencyResult.allCurrencies);
                    throw new Error(`❌ Multiple currencies detected in ${file.name}: ${currencies.join(', ')}. This usually indicates currency transitions or mixed-currency portfolios. Please ensure your portfolio file contains only one currency. If your country changed currencies, please use separate files for each currency period.`);
                }
                
                analysis.currency = detectedCurrency;
                analysis.currencyValidation = {
                    isMixed: currencyResult.isMixed,
                    allCurrencies: Array.from(currencyResult.allCurrencies),
                    confidence: currencyResult.confidence
                };
                config.debug(`💰 Currency detected from price columns for ${file.name}:`, analysis.currency);
            } else if (type === 'transaction') {
                // Transaction files must also detect currency for validation
                const currencyResult = await this.currencyDetector.detectCurrency(rawData, worksheet, {
                    interactive: false,
                    showDialogs: false
                });
                
                const detectedCurrency = currencyResult.currencyCode || currencyResult.primaryCurrency;
                
                if (!detectedCurrency) {
                    throw new Error(`❌ Currency detection failed for transaction file ${file.name}. No currency found in Excel file. Please ensure your transaction file contains currency codes or symbols.`);
                }
                
                // Check for multiple currencies in single transaction file
                if (currencyResult.isMixed || currencyResult.allCurrencies.size > 1) {
                    const currencies = Array.from(currencyResult.allCurrencies);
                    throw new Error(`❌ Multiple currencies detected in transaction file ${file.name}: ${currencies.join(', ')}. This usually indicates currency transitions or mixed-currency transactions. Please ensure your transaction file contains only one currency.`);
                }
                
                analysis.currency = detectedCurrency;
                analysis.currencyValidation = {
                    isMixed: currencyResult.isMixed,
                    allCurrencies: Array.from(currencyResult.allCurrencies),
                    confidence: currencyResult.confidence
                };
                config.debug(`💰 Currency detected from transaction file ${file.name}:`, analysis.currency);
            } else {
                throw new Error(`❌ No currency detector available for ${file.name}. Cannot analyze file without currency detection.`);
            }

            // Company detection moved to FileParser after parsing for better accuracy

            return analysis;

        } catch (error) {
            config.error(`❌ Failed to analyze ${file.name}:`, error);
            throw new Error(`File analysis failed for ${file.name}: ${error.message}`);
        }
    }

    /**
     * Validate compatibility between files and create unified result
     * Portfolio file is master for currency/language decisions
     * @param {Object} results - Results object with portfolio/transaction analysis
     * @returns {Object} Unified analysis result
     */
    validateAndUnify(results) {
        const { portfolio, transaction } = results;
        
        // Portfolio file is master for all decisions
        const unified = {
            language: portfolio.language,  // Auto-select portfolio language (user can override)
            currency: portfolio.currency,  // Portfolio currency is always master
            hasTransactions: !!transaction,
            portfolioFile: portfolio.fileName,
            transactionFile: transaction?.fileName || null
        };

        // Validate inter-file compatibility if transaction exists
        if (transaction) {
            this.validateCurrencyCompatibility(unified, portfolio, transaction, results);
            this.validateLanguageCompatibility(unified, portfolio, transaction, results);
        }

        // Company validation moved to FileParser after parsing

        // Update compatibility status
        results.isCompatible = results.errors.length === 0;

        config.debug('🔄 File compatibility validation:', {
            unified: unified,
            compatible: results.isCompatible,
            warnings: results.warnings.length,
            errors: results.errors.length
        });

        return unified;
    }

    /**
     * Validate currency compatibility between files
     */
    validateCurrencyCompatibility(unified, portfolioAnalysis, transactionAnalysis, results) {
        if (transactionAnalysis.currency !== unified.currency) {
            results.errors.push({
                type: 'CURRENCY_MISMATCH',
                message: `Currency mismatch: Portfolio (${unified.currency}) vs Transaction (${transactionAnalysis.currency})`,
                resolution: 'Transaction file will be ignored. Please upload a transaction file with matching currency.',
                details: {
                    portfolioCurrency: unified.currency,
                    transactionCurrency: transactionAnalysis.currency,
                    masterFile: 'portfolio'
                }
            });
            config.warn('⚠️ Currency mismatch detected - transaction file will be rejected');
        }
    }

    /**
     * Validate language compatibility between files (warning only)
     */
    validateLanguageCompatibility(unified, portfolioAnalysis, transactionAnalysis, results) {
        if (transactionAnalysis.language !== unified.language) {
            results.warnings.push({
                type: 'LANGUAGE_MISMATCH',
                message: `Language difference: Portfolio (${unified.language}) vs Transaction (${transactionAnalysis.language})`,
                resolution: 'Both files will be parsed consistently. UI language follows portfolio file.',
                details: {
                    portfolioLanguage: unified.language,
                    transactionLanguage: transactionAnalysis.language,
                    masterFile: 'portfolio'
                }
            });
            config.debug('ℹ️ Language difference detected - using portfolio language');
        }
    }



    /**
     * Validate files against cached data (handles sequential uploads)
     * @param {File|null} portfolioFile - New portfolio file (if any)
     * @param {File|null} transactionFile - New transaction file (if any) 
     * @param {Object|null} cachedData - Existing cached data from database
     * @returns {Object} Validation result with errors/warnings
     */
    async validateWithCache(portfolioFile, transactionFile, cachedData) {
        const results = {
            errors: [],
            warnings: [],
            shouldClearCachedTransactions: false
        };

        try {
            // If no cached data, no validation needed - regular flow
            if (!cachedData || (!cachedData.portfolioData && !cachedData.transactionData)) {
                config.debug('📄 No cached data - skipping cache validation');
                return results;
            }

            let masterCurrency = null;
            
            // Determine master currency (Portfolio always wins)
            if (portfolioFile) {
                // New portfolio file - analyze it to get currency
                config.debug('📄 New portfolio file uploaded - analyzing currency...');
                const portfolioAnalysis = await this.analyzeFile(portfolioFile, 'portfolio');
                masterCurrency = portfolioAnalysis.currency;
                config.debug(`💰 New portfolio currency: ${masterCurrency}`);
                
                // If we have cached transaction with different currency, clear it
                if (cachedData.transactionData && cachedData.transactionData.currency !== masterCurrency) {
                    results.shouldClearCachedTransactions = true;
                    results.warnings.push({
                        type: 'CACHE_CLEARED',
                        message: `Cached transaction data (${cachedData.transactionData.currency}) cleared due to new portfolio currency (${masterCurrency})`,
                        details: {
                            reason: 'Currency mismatch with new portfolio file'
                        }
                    });
                    config.debug('🗑️ Will clear cached transaction due to currency mismatch with new portfolio');
                }
            } else if (cachedData.portfolioData) {
                // Use cached portfolio as master
                masterCurrency = cachedData.currency;  // Currency is stored at top level of cached data
                config.debug(`💰 Using cached portfolio currency as master: ${masterCurrency}`);
            }

            // Validate new transaction file against master currency
            if (transactionFile && masterCurrency) {
                config.debug('📄 New transaction file uploaded - validating against master currency...');
                const transactionAnalysis = await this.analyzeFile(transactionFile, 'transaction');
                
                if (transactionAnalysis.currency !== masterCurrency) {
                    results.errors.push({
                        type: 'CURRENCY_MISMATCH',
                        message: `Currency mismatch: Portfolio (${masterCurrency}) vs Transaction file (${transactionAnalysis.currency})`,
                        resolution: 'Transaction file will be ignored. Please upload a transaction file with matching currency.',
                        details: {
                            portfolioCurrency: masterCurrency,
                            transactionCurrency: transactionAnalysis.currency,
                            masterFile: 'portfolio'
                        }
                    });
                    config.warn('⚠️ Transaction file currency mismatch - will be rejected');
                } else {
                    config.debug('✅ Transaction file currency matches master currency');
                }
            }

            return results;

        } catch (error) {
            config.error('❌ Error during cache validation:', error);
            results.errors.push({
                type: 'CACHE_VALIDATION_ERROR',
                message: `Cache validation failed: ${error.message}`,
                resolution: 'Please try clearing cached data and uploading files again.'
            });
            return results;
        }
    }

    /**
     * Read Excel file using SheetJS (reused from existing code)
     */
    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { 
                        type: 'array',
                        cellDates: true,
                        raw: false
                    });
                    resolve(workbook);
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = function() {
                reject(new Error('Failed to read file'));
            };
            
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Get analysis summary for debugging
     */
    getAnalysisSummary(analysisResult) {
        return {
            portfolioFile: analysisResult.portfolio?.fileName,
            transactionFile: analysisResult.transaction?.fileName || 'none',
            language: analysisResult.unified?.language,
            currency: analysisResult.unified?.currency,
            compatible: analysisResult.isCompatible,
            warnings: analysisResult.warnings.length,
            errors: analysisResult.errors.length
        };
    }


    /**
     * Find instrument column index based on language and file type
     * @param {Array} headerRow - First row with column headers
     * @param {String} language - Detected language
     * @param {String} fileType - 'portfolio' or 'transaction'  
     * @returns {Number} Column index or -1 if not found
     */
    findInstrumentColumnIndex(headerRow, language, fileType) {
        if (!headerRow) return -1;

        // Get translated instrument header for the detected language
        const instrumentHeader = this.getTranslatedHeader('instrument', language);
        config.debug(`🏢 Looking for instrument column with header: "${instrumentHeader}" in language: ${language}`);

        // Debug: Log all headers for inspection
        config.debug(`🔍 All column headers:`, headerRow.map((h, i) => `${i}: "${h}"`));
        
        // Search for instrument column (exact case insensitive match)
        for (let i = 0; i < headerRow.length; i++) {
            const header = headerRow[i]?.toString().trim() || '';
            
            config.debug(`🔍 Comparing header[${i}]: "${header}" vs expected: "${instrumentHeader}"`);
            
            // Use exact matching to avoid substring conflicts (e.g., "Instrument type" vs "Instrument")
            if (header.toLowerCase() === instrumentHeader.toLowerCase()) {
                config.debug(`🏢 Found instrument column at index ${i}: "${header}" (exact match)`);
                return i;
            }
        }

        // Fallback: Try manual search for any column containing "instrument" (case insensitive)
        config.debug(`🔍 Instrument column not found with exact match, trying fallback search...`);
        
        for (let i = 0; i < headerRow.length; i++) {
            const header = headerRow[i]?.toString().toLowerCase().trim() || '';
            
            // Look for any column that contains "instrument" but not "type"
            if (header.includes('instrument') && !header.includes('type')) {
                config.debug(`🏢 Found instrument column via fallback at index ${i}: "${headerRow[i]}"`);
                return i;
            }
        }
        
        // Last resort: Try common column positions
        // Portfolio files: instrument at column 3, Transaction files: instrument at column 6 
        const fallbackIndex = fileType === 'portfolio' ? 3 : 6;
        
        if (headerRow.length > fallbackIndex) {
            config.debug(`🏢 Using hardcoded fallback instrument column index ${fallbackIndex} for ${fileType}`);
            return fallbackIndex;
        }

        return -1;
    }

    /**
     * Get translated header for a given key and language
     * @param {String} key - Translation key (e.g., 'instrument')
     * @param {String} language - Language code (e.g., 'english', 'german')
     * @returns {String} Translated header text
     */
    getTranslatedHeader(key, language) {
        // Access global translation manager if available
        if (typeof translationManager !== 'undefined') {
            const originalLang = translationManager.currentLanguage;
            translationManager.currentLanguage = language;
            const header = translationManager.t(key);
            translationManager.currentLanguage = originalLang;
            return header;
        }
        
        // Fallback: Use TRANSLATION_DATA directly if translation manager not available
        if (typeof TRANSLATION_DATA !== 'undefined' && TRANSLATION_DATA[key] && TRANSLATION_DATA[key][language]) {
            return TRANSLATION_DATA[key][language];
        }
        
        // Last resort fallback
        const fallbacks = {
            instrument: 'Instrument',
            allocation_date: 'Allocation date',
            order_reference: 'Order reference'
        };
        
        return fallbacks[key] || key;
    }
}

// Global file analyzer instance
const fileAnalyzer = new FileAnalyzer();

// Node.js support
if (typeof module !== 'undefined' && module.exports) {
    module.exports = FileAnalyzer;
}