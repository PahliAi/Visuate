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
                    language: cachedData.isEnglish ? 'english' : 'other', // Simple fallback for cached data
                    currency: cachedData.currency,
                    company: cachedData.company || 'Other',
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
                
                analysis.currency = detectedCurrency;
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
                
                analysis.currency = detectedCurrency;
                config.debug(`💰 Currency detected from transaction file ${file.name}:`, analysis.currency);
            } else {
                throw new Error(`❌ No currency detector available for ${file.name}. Cannot analyze file without currency detection.`);
            }

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
            company: null, // Will be determined after parsing both files
            hasTransactions: !!transaction,
            portfolioFile: portfolio.fileName,
            transactionFile: transaction?.fileName || null
        };

        // Validate inter-file compatibility if transaction exists
        if (transaction) {
            this.validateCurrencyCompatibility(unified, portfolio, transaction, results);
            this.validateLanguageCompatibility(unified, portfolio, transaction, results);
        }

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
     * Determine company after both files are parsed (placeholder)
     * This will be called after actual parsing is complete
     * @param {Object} portfolioData - Parsed portfolio data
     * @param {Object} transactionData - Parsed transaction data (optional)
     * @returns {String} 'Allianz' or 'Other'
     */
    determineCompanyFromParsedData(portfolioData, transactionData = null) {
        // Check if ALL entries in BOTH files have "Allianz share" as instrument
        const portfolioAllAllianz = portfolioData.entries.every(entry => 
            entry.instrument && entry.instrument.toLowerCase().includes('allianz share')
        );

        if (!transactionData || !transactionData.entries || transactionData.entries.length === 0) {
            const company = portfolioAllAllianz ? 'Allianz' : 'Other';
            config.debug('🏢 Company determined (portfolio only):', company);
            return company;
        }

        // Check transaction file instruments
        const transactionInstruments = transactionData.entries
            .map(entry => entry.instrument)
            .filter(Boolean);
        
        const transactionAllAllianz = transactionInstruments.length > 0 && 
            transactionInstruments.every(instrument => 
                instrument.toLowerCase().includes('allianz share')
            );

        const company = (portfolioAllAllianz && transactionAllAllianz) ? 'Allianz' : 'Other';
        
        config.debug('🏢 Company determined (both files):', {
            company: company,
            portfolioAllAllianz: portfolioAllAllianz,
            transactionAllAllianz: transactionAllAllianz,
            transactionInstruments: transactionInstruments.length
        });

        return company;
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
            company: analysisResult.unified?.company || 'pending',
            compatible: analysisResult.isCompatible,
            warnings: analysisResult.warnings.length,
            errors: analysisResult.errors.length
        };
    }
}

// Global file analyzer instance
const fileAnalyzer = new FileAnalyzer();

// Node.js support
if (typeof module !== 'undefined' && module.exports) {
    module.exports = FileAnalyzer;
}