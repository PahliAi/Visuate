/**
 * Generic Currency Detection Module
 * Detects ANY currency from Excel files without hardcoding specific currencies
 * Used by both main application and unit tests
 */

class CurrencyDetector {
    constructor() {
        // Initialize currency mappings
        if (typeof CurrencyMappings !== 'undefined') {
            this.currencyMappings = new CurrencyMappings();
        } else {
            console.warn('⚠️ CurrencyMappings not available - using fallback');
            this.currencyMappings = null;
        }
        
        // Generic patterns to identify currency formatting
        this.currencyPatterns = {
            // Excel format patterns that indicate currency
            formatPatterns: [
                /["'][^"']*?[\$€£¥₹₩₽¢₡₪₦₨₫₵₲₴₸₺₼₾₿][^"']*?["']/,  // Quoted currency symbols
                /\[[^\]]*?[\$€£¥₹₩₽¢₡₪₦₨₫₵₲₴₸₺₼₾₿][^\]]*?\]/,      // Bracketed currency symbols
                /["'][A-Z]{3}["']/,                                             // Three-letter currency codes in quotes
                /\[[A-Z]{3}[^\]]*\]/,                                          // Three-letter currency codes in brackets
                /#.*?[\$€£¥₹₩₽¢₡₪₦₨₫₵₲₴₸₺₼₾₿]/,                     // Number format with currency symbol
                /[\$€£¥₹₩₽¢₡₪₦₨₫₵₲₴₸₺₼₾₿].*?#/                      // Currency symbol with number format
            ],
            
            // Display text patterns that indicate currency
            displayPatterns: [
                /[\$€£¥₹₩₽¢₡₪₦₨₫₵₲₴₸₺₼₾₿]/,                        // Any currency symbol
                /\b[A-Z]{3}\b/,                                               // Three-letter codes (USD, EUR, etc.)
                /\b\d+[.,]\d+\s*[A-Z]{3}\b/,                                 // Number with currency code
                /\b[A-Z]{3}\s*\d+[.,]\d+\b/                                  // Currency code with number
            ],
            
            // Common currency symbols for extraction
            currencySymbols: '\\$€£¥₹₩₽¢₡₪₦₨₫₵₲₴₸₺₼₾₿',
            
            // Three-letter currency code pattern
            currencyCodePattern: /\b[A-Z]{3}\b/g
        };
    }

    /**
     * Detect currency from Excel file data with enhanced user interaction
     * @param {Array} rawData - Raw Excel data from XLSX.utils.sheet_to_json
     * @param {Object} worksheet - Excel worksheet object
     * @param {Object} options - Detection options (interactive, showDialogs, etc.)
     * @returns {Object} Enhanced detection result with symbol, code, and user choices
     */
    async detectCurrency(rawData, worksheet = null, options = {}) {
        const { interactive = false, showDialogs = true } = options;
        
        const result = {
            // Core detection results
            primaryCurrency: null,
            currencySymbol: null,
            currencyCode: null,
            allCurrencies: new Set(),
            confidence: 0,
            isMixed: false,
            detectionMethod: null,
            
            // Enhanced results
            userChoice: null,
            requiresUserInput: false,
            scenario: null, // 'none_found', 'multiple_in_file', 'multiple_between_files', 'success'
            
            // Technical details
            details: {
                formatFindings: [],
                displayFindings: [],
                extractedCodes: new Set(),
                extractedSymbols: new Set()
            }
        };

        // Step 1: Check Excel cell formatting (most reliable)
        if (worksheet) {
            this._scanCellFormats(worksheet, rawData, result);
        }

        // Step 2: Check display text content (including formatted display text)
        this._scanDisplayContent(rawData, result, worksheet);

        // Step 3: Analyze results and determine primary currency
        this._analyzeCurrencyFindings(result);
        
        // Step 4: Enhance results with symbol/code mapping
        await this._enhanceWithCurrencyMapping(result);
        
        // Step 5: Handle different scenarios with user interaction
        if (interactive && showDialogs) {
            await this._handleUserInteraction(result);
        }
        
        return result;
    }
    
    /**
     * Enhance results with currency symbol/code mapping
     */
    async _enhanceWithCurrencyMapping(result) {
        // Determine scenario for user interaction first
        const currencyCount = result.allCurrencies.size;
        
        if (currencyCount === 0) {
            result.scenario = 'none_found';
            result.requiresUserInput = true;
        } else if (currencyCount === 1) {
            result.scenario = 'success';
            result.requiresUserInput = false;
        } else if (currencyCount > 1) {
            result.scenario = 'multiple_in_file';
            result.requiresUserInput = true;
            result.isMixed = true;
        }
        
        // Handle currency symbol/code mapping
        if (!this.currencyMappings) {
            // Fallback without mappings - use currency code as symbol
            result.currencyCode = result.primaryCurrency;
            result.currencySymbol = result.primaryCurrency;
            return;
        }
        
        if (result.primaryCurrency) {
            const currencyInfo = this.currencyMappings.getByCode(result.primaryCurrency);
            if (currencyInfo) {
                result.currencyCode = result.primaryCurrency;
                result.currencySymbol = currencyInfo.symbol;
            } else {
                // Currency not found in mappings - use code as fallback
                result.currencyCode = result.primaryCurrency;
                result.currencySymbol = result.primaryCurrency;
            }
        }
    }
    
    /**
     * Handle user interaction for different scenarios
     */
    async _handleUserInteraction(result) {
        switch (result.scenario) {
            case 'none_found':
                result.userChoice = await this._showCurrencySelectionDialog('all');
                break;
                
            case 'multiple_in_file':
                const foundCurrencies = Array.from(result.allCurrencies);
                result.userChoice = await this._showCurrencySelectionDialog('found', foundCurrencies);
                break;
                
            case 'success':
                // No user interaction needed
                break;
        }
        
        // Apply user choice if made
        if (result.userChoice) {
            result.primaryCurrency = result.userChoice.code;
            result.currencyCode = result.userChoice.code;
            result.currencySymbol = result.userChoice.symbol;
            result.detectionMethod = 'user_selected';
            result.confidence = 100; // User selection is 100% confident
        }
    }
    
    /**
     * Show currency selection dialog to user
     */
    async _showCurrencySelectionDialog(type, foundCurrencies = []) {
        if (!this.currencyMappings) return null;
        
        return new Promise((resolve) => {
            let isResolved = false;
            
            const overlay = document.createElement('div');
            overlay.className = 'currency-selection-overlay';
            overlay.style.cssText = `
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(0, 0, 0, 0.8);
                display: flex;
                justify-content: center;
                align-items: center;
                z-index: 10000;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            `;
            
            const modal = document.createElement('div');
            modal.style.cssText = `
                background: white;
                padding: 30px;
                border-radius: 12px;
                max-width: 500px;
                max-height: 80vh;
                overflow-y: auto;
                text-align: center;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
            `;
            
            const title = type === 'all' ? 
                'Select Currency' : 
                'Multiple Currencies Found';
                
            const message = type === 'all' ?
                'No currency could be detected from your Excel files. Please select your currency:' :
                `Multiple currencies were found in your file: ${foundCurrencies.join(', ')}. Please choose which currency to use:`;
            
            const currencyOptions = type === 'all' ? 
                this.currencyMappings.createCurrencyOptions() :
                this.currencyMappings.createCurrencyOptions(foundCurrencies);
            
            modal.innerHTML = `
                <h3 style="margin: 0 0 15px 0; color: #333;">${title}</h3>
                <p style="margin: 0 0 25px 0; color: #666; line-height: 1.5;">${message}</p>
                <div style="max-height: 300px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px;">
                    ${currencyOptions.map(currency => `
                        <div class="currency-option" data-code="${currency.code}" style="
                            padding: 12px 15px;
                            border-bottom: 1px solid #eee;
                            cursor: pointer;
                            display: flex;
                            justify-content: space-between;
                            align-items: center;
                            transition: background-color 0.2s;
                        " onmouseover="this.style.backgroundColor='#f5f5f5'" 
                           onmouseout="this.style.backgroundColor='white'">
                            <span style="font-weight: 500; font-size: 18px;">${currency.symbol}</span>
                            <span style="color: #666;">${currency.code} - ${currency.name}</span>
                        </div>
                    `).join('')}
                </div>
                <button onclick="this.closest('.currency-selection-overlay').remove()" style="
                    margin-top: 20px;
                    padding: 10px 20px;
                    background: #6c757d;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    cursor: pointer;
                ">Cancel</button>
            `;
            
            const handleClick = (event) => {
                if (isResolved) return;
                
                const option = event.target.closest('.currency-option');
                if (!option) return;
                
                const code = option.dataset.code;
                const selectedCurrency = currencyOptions.find(c => c.code === code);
                
                if (selectedCurrency) {
                    isResolved = true;
                    
                    // Remove all currency selection overlays
                    document.querySelectorAll('.currency-selection-overlay').forEach(el => el.remove());
                    
                    resolve(selectedCurrency);
                }
            };
            
            modal.addEventListener('click', handleClick);
            overlay.appendChild(modal);
            document.body.appendChild(overlay);
        });
    }

    /**
     * Scan Excel cell formats for currency information
     */
    _scanCellFormats(worksheet, rawData, result) {
        // Check if XLSX is available (for cell reference encoding)
        const XLSX = this._getXLSX();
        if (!XLSX) {
            if (typeof config !== 'undefined' && config.warn) {
                config.warn('⚠️ XLSX not available for cell format scanning');
            }
            return;
        }
        
        const scanRows = Math.min(100, rawData.length);
        const scanCols = 20;

        for (let r = 0; r < scanRows; r++) {
            for (let c = 0; c < scanCols; c++) {
                const cellRef = XLSX.utils.encode_cell({r, c});
                const cell = worksheet[cellRef];

                if (cell && cell.z) {
                    const format = cell.z.toString();
                    const currencies = this._extractCurrenciesFromFormat(format);
                    
                    if (currencies.length > 0) {
                        result.details.formatFindings.push({
                            cell: cellRef,
                            format: format,
                            currencies: currencies
                        });
                        
                        currencies.forEach(currency => {
                            result.allCurrencies.add(currency);
                            if (currency.length === 3) {
                                result.details.extractedCodes.add(currency);
                            } else {
                                result.details.extractedSymbols.add(currency);
                            }
                        });
                    }
                }
            }
        }

        if (result.details.formatFindings.length > 0) {
            result.confidence += 100; // High confidence from format detection
            result.detectionMethod = 'format';
        }
    }

    /**
     * Scan display content for currency information
     * This includes both raw cell values AND formatted display text (cell.w)
     */
    _scanDisplayContent(rawData, result, worksheet = null) {
        const XLSX = this._getXLSX();
        const scanRows = Math.min(100, rawData.length);

        for (let r = 0; r < scanRows; r++) {
            const row = rawData[r];
            if (!row) continue;

            for (let c = 0; c < Math.min(20, row.length); c++) {
                const cellValue = row[c];

                // Check raw cell value for currency text
                if (typeof cellValue === 'string' && cellValue.trim()) {
                    const currencies = this._extractCurrenciesFromText(cellValue);
                    
                    if (currencies.length > 0) {
                        result.details.displayFindings.push({
                            cell: `R${r}C${c}`,
                            text: cellValue,
                            currencies: currencies,
                            source: 'raw'
                        });

                        currencies.forEach(currency => {
                            result.allCurrencies.add(currency);
                            if (currency.length === 3) {
                                result.details.extractedCodes.add(currency);
                            } else {
                                result.details.extractedSymbols.add(currency);
                            }
                        });
                    }
                }

                // Also check formatted display text (cell.w) if available
                if (worksheet && XLSX) {
                    const cellRef = XLSX.utils.encode_cell({r, c});
                    const cell = worksheet[cellRef];
                    
                    if (cell && cell.w && typeof cell.w === 'string') {
                        const formattedText = cell.w.toString().trim();
                        const currencies = this._extractCurrenciesFromText(formattedText);
                        
                        if (currencies.length > 0) {
                            result.details.displayFindings.push({
                                cell: cellRef,
                                text: formattedText,
                                currencies: currencies,
                                source: 'formatted'
                            });

                            currencies.forEach(currency => {
                                result.allCurrencies.add(currency);
                                if (currency.length === 3) {
                                    result.details.extractedCodes.add(currency);
                                } else {
                                    result.details.extractedSymbols.add(currency);
                                }
                            });
                        }
                    }
                }
            }
        }

        if (result.details.displayFindings.length > 0) {
            // Higher confidence for formatted text findings
            const formattedFindings = result.details.displayFindings.filter(f => f.source === 'formatted');
            const rawFindings = result.details.displayFindings.filter(f => f.source === 'raw');
            
            result.confidence += (formattedFindings.length * 50) + (rawFindings.length * 10);
            
            if (!result.detectionMethod) {
                result.detectionMethod = formattedFindings.length > 0 ? 'formatted_display' : 'raw_display';
            }
        }
    }

    /**
     * Extract currency codes/symbols from Excel format string
     */
    _extractCurrenciesFromFormat(format) {
        const currencies = [];
        
        // Extract currency symbols
        const symbolMatches = format.match(new RegExp(`[${this.currencyPatterns.currencySymbols}]`, 'g'));
        if (symbolMatches) {
            currencies.push(...symbolMatches);
        }

        // Extract currency codes (like USD, EUR, CNY)
        const codeMatches = format.match(this.currencyPatterns.currencyCodePattern);
        if (codeMatches) {
            currencies.push(...codeMatches.filter(code => this._isLikelyCurrencyCode(code)));
        }

        // Special patterns (US$, CN¥, etc.)
        const specialPatterns = [
            /US\$/g, /CN¥/g, /JP¥/g, /HK\$/g, /NZ\$/g, /A\$/g, /C\$/g, /S\$/g
        ];
        
        specialPatterns.forEach(pattern => {
            const matches = format.match(pattern);
            if (matches) {
                currencies.push(...matches);
            }
        });

        return [...new Set(currencies)]; // Remove duplicates
    }

    /**
     * Extract currency codes/symbols from display text
     */
    _extractCurrenciesFromText(text) {
        const currencies = [];
        
        // Special patterns for compound currency codes (like "US$", "CN¥") - MOST SPECIFIC FIRST
        const specialPatterns = [
            { pattern: /US\$/g, currency: 'USD' },
            { pattern: /CN¥/g, currency: 'CNY' },
            { pattern: /JP¥/g, currency: 'JPY' },
            { pattern: /HK\$/g, currency: 'HKD' },
            { pattern: /NZ\$/g, currency: 'NZD' },
            { pattern: /A\$/g, currency: 'AUD' },
            { pattern: /C\$/g, currency: 'CAD' },
            { pattern: /S\$/g, currency: 'SGD' },
            { pattern: /NT\$/g, currency: 'TWD' },
            { pattern: /R\$/g, currency: 'BRL' }
        ];
        
        // Check special patterns first (these are definitive)
        // Track which parts of the text have been matched to avoid overlap
        let foundSpecialPattern = false;
        let textCopy = text;
        
        specialPatterns.forEach(({ pattern, currency }) => {
            const matches = textCopy.match(pattern);
            if (matches) {
                currencies.push(currency);
                foundSpecialPattern = true;
                
                // Remove matched patterns from text to prevent overlap
                // For example, remove "US$" so "S$" won't also match
                textCopy = textCopy.replace(pattern, '');
            }
        });
        
        // If we found a specific pattern like "US$", don't also add generic "$" mapping
        if (!foundSpecialPattern) {
            // Extract basic currency symbols and map to codes (only if no specific pattern found)
            const symbolMappings = {
                '$': 'USD',   // Default $ to USD only if no US$, HK$, etc. found
                '€': 'EUR',
                '£': 'GBP',
                '¥': 'JPY',   // Default ¥ to JPY only if no CN¥, JP¥ found
                '₹': 'INR',
                '₩': 'KRW',
                '₽': 'RUB',
                '¢': 'USD',   // Cents usually USD
                '₡': 'CRC',
                '₪': 'ILS',
                '₦': 'NGN',
                '₨': 'PKR',
                '₫': 'VND',
                '₵': 'GHS',
                '₲': 'PYG',
                '₴': 'UAH',
                '₸': 'KZT',
                '₺': 'TRY',
                '₼': 'AZN',
                '₾': 'GEL',
                '₿': 'BTC'
            };
            
            Object.keys(symbolMappings).forEach(symbol => {
                if (text.includes(symbol)) {
                    currencies.push(symbolMappings[symbol]);
                }
            });
        }

        // Extract 3-letter currency codes
        const codeMatches = text.match(this.currencyPatterns.currencyCodePattern);
        if (codeMatches) {
            currencies.push(...codeMatches.filter(code => this._isLikelyCurrencyCode(code)));
        }

        return [...new Set(currencies)]; // Remove duplicates
    }

    /**
     * Check if a 3-letter code is a valid ISO 4217 currency code
     */
    _isLikelyCurrencyCode(code) {
        // Valid ISO 4217 currency codes - only accept these to avoid false positives like "IBM"
        const validIsoCodes = [
            'AED', 'AFN', 'ALL', 'AMD', 'ANG', 'AOA', 'ARS', 'AUD', 'AWG', 'AZN',
            'BAM', 'BBD', 'BDT', 'BGN', 'BHD', 'BIF', 'BMD', 'BND', 'BOB', 'BRL', 'BSD', 'BTN', 'BWP', 'BYN', 'BZD',
            'CAD', 'CDF', 'CHF', 'CLP', 'CNY', 'COP', 'CRC', 'CUP', 'CVE', 'CZK',
            'DJF', 'DKK', 'DOP', 'DZD',
            'EGP', 'ERN', 'ETB', 'EUR',
            'FJD', 'FKP',
            'GBP', 'GEL', 'GGP', 'GHS', 'GIP', 'GMD', 'GNF', 'GTQ', 'GYD',
            'HKD', 'HNL', 'HRK', 'HTG', 'HUF',
            'IDR', 'ILS', 'IMP', 'INR', 'IQD', 'IRR', 'ISK',
            'JEP', 'JMD', 'JOD', 'JPY',
            'KES', 'KGS', 'KHR', 'KMF', 'KPW', 'KRW', 'KWD', 'KYD', 'KZT',
            'LAK', 'LBP', 'LKR', 'LRD', 'LSL', 'LYD',
            'MAD', 'MDL', 'MGA', 'MKD', 'MMK', 'MNT', 'MOP', 'MRU', 'MUR', 'MVR', 'MWK', 'MXN', 'MYR', 'MZN',
            'NAD', 'NGN', 'NIO', 'NOK', 'NPR', 'NZD',
            'OMR',
            'PAB', 'PEN', 'PGK', 'PHP', 'PKR', 'PLN', 'PYG',
            'QAR',
            'RON', 'RSD', 'RUB', 'RWF',
            'SAR', 'SBD', 'SCR', 'SDG', 'SEK', 'SGD', 'SHP', 'SLE', 'SLL', 'SOS', 'SRD', 'SSP', 'STN', 'SYP', 'SZL',
            'THB', 'TJS', 'TMT', 'TND', 'TOP', 'TRY', 'TTD', 'TWD', 'TZS',
            'UAH', 'UGX', 'USD', 'UYU', 'UYW', 'UZS',
            'VED', 'VES', 'VND', 'VUV',
            'WST',
            'XAF', 'XCD', 'XDR', 'XOF', 'XPF',
            'YER',
            'ZAR', 'ZMW', 'ZWL'
        ];
        
        return validIsoCodes.includes(code.toUpperCase());
    }

    /**
     * Analyze findings and determine primary currency
     */
    _analyzeCurrencyFindings(result) {
        const allCurrencies = Array.from(result.allCurrencies);
        
        if (allCurrencies.length === 0) {
            result.primaryCurrency = null;
            result.confidence = 0;
            return;
        }

        if (allCurrencies.length === 1) {
            result.primaryCurrency = allCurrencies[0];
            result.isMixed = false;
            return;
        }

        // Multiple currencies detected
        result.isMixed = true;
        
        // Count occurrences to determine primary
        const currencyCount = {};
        
        [...result.details.formatFindings, ...result.details.displayFindings].forEach(finding => {
            finding.currencies.forEach(currency => {
                currencyCount[currency] = (currencyCount[currency] || 0) + 1;
            });
        });

        // Sort by frequency
        const sortedCurrencies = Object.keys(currencyCount)
            .sort((a, b) => currencyCount[b] - currencyCount[a]);

        result.primaryCurrency = sortedCurrencies[0];
        
        // Log mixed currency warning
        if (typeof config !== 'undefined' && config.debug) {
            config.debug('💰 Mixed currencies detected:', currencyCount);
        }
    }

    /**
     * Create user-friendly error message for mixed currencies
     */
    createMixedCurrencyError(result) {
        if (!result.isMixed) return null;

        const currencies = Array.from(result.allCurrencies).join(', ');
        return {
            title: 'Mixed Currencies Detected',
            message: `Your Excel file contains multiple currencies: ${currencies}. 
                     Equate can only process files with a single currency. 
                     Please ensure all values in your portfolio and transaction files use the same currency.`,
            details: {
                foundCurrencies: Array.from(result.allCurrencies),
                suggestedAction: 'Check your EquatePlus export settings or contact your administrator for files in a single currency.'
            }
        };
    }

    /**
     * Validate currency consistency between portfolio and transaction files with enhanced UX
     */
    async validateCurrencyConsistency(portfolioResult, transactionResult, options = {}) {
        const { interactive = false, showDialogs = true } = options;
        
        // Handle cases where currency couldn't be determined
        if (!portfolioResult.primaryCurrency && !transactionResult.primaryCurrency) {
            if (interactive && showDialogs) {
                const selectedCurrency = await this._showCurrencySelectionDialog('all');
                if (selectedCurrency) {
                    return {
                        isValid: true,
                        scenario: 'user_selected_both',
                        currency: selectedCurrency.code,
                        currencySymbol: selectedCurrency.symbol,
                        message: `Selected currency: ${selectedCurrency.symbol} (${selectedCurrency.code})`
                    };
                }
            }
            return {
                isValid: false,
                scenario: 'both_undetermined',
                error: 'Currency could not be determined from either file',
                requiresUserInput: true
            };
        }

        // Handle mixed currencies in individual files
        if (portfolioResult.isMixed) {
            if (interactive && showDialogs) {
                const foundCurrencies = Array.from(portfolioResult.allCurrencies);
                const selectedCurrency = await this._showCurrencySelectionDialog('found', foundCurrencies);
                if (selectedCurrency) {
                    portfolioResult.primaryCurrency = selectedCurrency.code;
                    portfolioResult.currencySymbol = selectedCurrency.symbol;
                    portfolioResult.isMixed = false;
                }
            } else {
                return {
                    isValid: false,
                    scenario: 'portfolio_mixed',
                    error: 'Portfolio file contains multiple currencies',
                    mixedCurrencyError: this.createEnhancedMixedCurrencyError(portfolioResult, 'portfolio'),
                    requiresUserInput: true
                };
            }
        }

        if (transactionResult.isMixed) {
            if (interactive && showDialogs) {
                const foundCurrencies = Array.from(transactionResult.allCurrencies);
                const selectedCurrency = await this._showCurrencySelectionDialog('found', foundCurrencies);
                if (selectedCurrency) {
                    transactionResult.primaryCurrency = selectedCurrency.code;
                    transactionResult.currencySymbol = selectedCurrency.symbol;
                    transactionResult.isMixed = false;
                }
            } else {
                return {
                    isValid: false,
                    scenario: 'transaction_mixed',
                    error: 'Transaction file contains multiple currencies',
                    mixedCurrencyError: this.createEnhancedMixedCurrencyError(transactionResult, 'transaction'),
                    requiresUserInput: true
                };
            }
        }

        // Handle currency mismatch between files
        if (portfolioResult.primaryCurrency && transactionResult.primaryCurrency &&
            portfolioResult.primaryCurrency !== transactionResult.primaryCurrency) {
            
            const portfolioSymbol = this.currencyMappings ? 
                this.currencyMappings.getSymbol(portfolioResult.primaryCurrency) : 
                portfolioResult.primaryCurrency;
            const transactionSymbol = this.currencyMappings ? 
                this.currencyMappings.getSymbol(transactionResult.primaryCurrency) : 
                transactionResult.primaryCurrency;

            return {
                isValid: false,
                scenario: 'currency_mismatch',
                error: `Currency mismatch between files`,
                portfolioCurrency: {
                    code: portfolioResult.primaryCurrency,
                    symbol: portfolioSymbol
                },
                transactionCurrency: {
                    code: transactionResult.primaryCurrency,
                    symbol: transactionSymbol
                },
                message: `Portfolio file uses ${portfolioSymbol} (${portfolioResult.primaryCurrency}) but transaction file uses ${transactionSymbol} (${transactionResult.primaryCurrency}). Please upload a transaction file that matches your portfolio currency.`,
                recommendation: `Upload a transaction file with ${portfolioSymbol} (${portfolioResult.primaryCurrency}) currency to match your portfolio.`
            };
        }

        // Success case
        const finalCurrency = portfolioResult.primaryCurrency || transactionResult.primaryCurrency;
        const finalSymbol = this.currencyMappings ? 
            this.currencyMappings.getSymbol(finalCurrency) : 
            finalCurrency;

        return {
            isValid: true,
            scenario: 'success',
            currency: finalCurrency,
            currencySymbol: finalSymbol,
            confidence: Math.min(portfolioResult.confidence || 0, transactionResult.confidence || 0),
            message: `Files are consistent with ${finalSymbol} (${finalCurrency}) currency`
        };
    }
    
    /**
     * Create enhanced mixed currency error with symbols
     */
    createEnhancedMixedCurrencyError(result, fileType) {
        if (!result.isMixed) return null;

        const currencies = Array.from(result.allCurrencies);
        const currencyDisplays = this.currencyMappings ? 
            currencies.map(code => {
                const info = this.currencyMappings.getByCode(code);
                return info ? `${info.symbol} (${code})` : code;
            }) : currencies;

        return {
            title: 'Multiple Currencies Detected',
            fileType: fileType,
            currencies: currencies,
            currencyDisplays: currencyDisplays,
            message: `Your ${fileType} file contains multiple currencies: ${currencyDisplays.join(', ')}. 
                     Equate can only process files with a single currency. 
                     Please ensure all values in your ${fileType} file use the same currency.`,
            details: {
                foundCurrencies: currencies,
                foundSymbols: currencyDisplays,
                suggestedAction: `Check your EquatePlus export settings to ensure a single currency is used, or choose which currency to use for calculations.`
            }
        };
    }

    /**
     * Get XLSX library reference - works in both Node.js and browser
     */
    _getXLSX() {
        // Try global XLSX first (browser)
        if (typeof XLSX !== 'undefined') {
            return XLSX;
        }
        
        // Try require (Node.js)
        if (typeof require !== 'undefined') {
            try {
                return require('xlsx');
            } catch (error) {
                // XLSX not available
                return null;
            }
        }
        
        return null;
    }
}

// Export for both browser and Node.js
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CurrencyDetector;
} else if (typeof window !== 'undefined') {
    window.CurrencyDetector = CurrencyDetector;
}