/**
 * Excel File Parser for Equate Portfolio Analysis
 * Handles parsing of PortfolioDetails, CompletedTransactions, and historical price data
 */

class FileParser {
    /**
     * Create a new FileParser instance
     * @constructor
     */
    constructor() {
        this.portfolioData = null;
        this.transactionData = null;
        this.historicalPrices = null;
        this.detectedDateFormat = null; // 'DD-MM-YYYY' or 'MM-DD-YYYY'
    }

    /**
     * Detect date format from Excel files before parsing data
     * Returns 'DD-MM-YYYY', 'MM-DD-YYYY', or null if ambiguous
     */
    async detectDateFormat(portfolioFile, transactionFile = null) {
        const dateFormats = {
            ddMm: 0,  // Evidence for DD-MM format
            mmDd: 0,  // Evidence for MM-DD format
            totalStrings: 0  // Track if we found any string dates at all
        };
        
        try {
            // Analyze portfolio file dates
            if (portfolioFile) {
                await this.analyzeDateFormatsInFile(portfolioFile, dateFormats);
            }
            
            // Analyze transaction file dates if provided
            if (transactionFile) {
                await this.analyzeDateFormatsInFile(transactionFile, dateFormats);
            }
            
            // Determine format based on evidence from both files combined (ALL date types)
            config.debug(`📅 Date format evidence: DD-MM=${dateFormats.ddMm}, MM-DD=${dateFormats.mmDd}, total strings=${dateFormats.totalStrings}`);
            
            if (dateFormats.ddMm > 0 || dateFormats.mmDd > 0) {
                // Found at least one unambiguous date (any type: string, Date object, or serial number)
                if (dateFormats.ddMm >= dateFormats.mmDd) {
                    this.detectedDateFormat = 'DD-MM-YYYY';
                    config.debug('📅 Detected date format: DD-MM-YYYY (European) - found unambiguous evidence');
                    return 'DD-MM-YYYY';
                } else {
                    this.detectedDateFormat = 'MM-DD-YYYY';
                    config.debug('📅 Detected date format: MM-DD-YYYY (US) - found unambiguous evidence');
                    return 'MM-DD-YYYY';
                }
            } else if (dateFormats.totalStrings === 0) {
                // No string dates found and no unambiguous evidence from other types
                config.debug('📅 No string dates and no unambiguous evidence - defaulting to DD-MM');
                this.detectedDateFormat = 'DD-MM-YYYY';
                return 'DD-MM-YYYY';
            } else {
                config.debug('📅 ALL dates across ALL files are ambiguous - will need user input');
                return null; // Truly ALL dates ambiguous - show popup
            }
            
        } catch (error) {
            config.warn('⚠️ Error during date format detection:', error.message);
            return null;
        }
    }
    
    /**
     * Analyze date formats in a single Excel file
     */
    async analyzeDateFormatsInFile(file, dateFormats) {
        const data = await this.readExcelFile(file);
        const worksheet = data.Sheets[data.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            raw: true
        });
        
        // Look for date patterns in all cells - analyze ALL date types
        for (let rowIndex = 0; rowIndex < rawData.length; rowIndex++) {
            const row = rawData[rowIndex];
            for (let colIndex = 0; colIndex < row.length; colIndex++) {
                const cellValue = row[colIndex];
                
                if (typeof cellValue === 'string') {
                    // Analyze string dates
                    dateFormats.totalStrings++;
                    this.analyzeStringDateFormat(cellValue, dateFormats);
                } else if (cellValue instanceof Date) {
                    // Date objects from Excel: check if day > 12 to determine format
                    const day = cellValue.getDate();
                    const month = cellValue.getMonth() + 1; // getMonth() is 0-based
                    
                    if (day > 12) {
                        dateFormats.ddMm++; // Day > 12 means original was DD-MM
                        config.debug(`📅 Found DD-MM evidence from Date object: day=${day} > 12`);
                    } else if (month > 12) {
                        dateFormats.mmDd++; // Month > 12 means original was MM-DD (shouldn't happen but check anyway)
                        config.debug(`📅 Found MM-DD evidence from Date object: month=${month} > 12`);
                    }
                }
                
                // Early exit if we found unambiguous evidence
                if (dateFormats.ddMm > 0 || dateFormats.mmDd > 0) {
                    config.debug('📅 Found unambiguous evidence, stopping scan early');
                    return;
                }
            }
        }
    }
    
    /**
     * Analyze a single string date for format clues
     */
    analyzeStringDateFormat(dateStr, dateFormats) {
        // Look for patterns like DD-MM-YYYY or MM-DD-YYYY with unambiguous day/month
        const datePattern = /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/;
        const match = dateStr.match(datePattern);
        
        if (match) {
            const first = parseInt(match[1]);
            const second = parseInt(match[2]);
            
            // If first number > 12, it must be day (DD-MM format)
            if (first > 12 && second <= 12) {
                dateFormats.ddMm++;
                config.debug(`📅 Found DD-MM evidence: ${dateStr} (day=${first} > 12)`);
            }
            // If second number > 12, it must be day (MM-DD format)
            else if (second > 12 && first <= 12) {
                dateFormats.mmDd++;
                config.debug(`📅 Found MM-DD evidence: ${dateStr} (day=${second} > 12)`);
            }
            // If both <= 12, it's ambiguous - no evidence either way
        }
    }
    
    /**
     * Show user dialog to choose date format when ambiguous
     */
    async promptUserForDateFormat() {
        return new Promise((resolve) => {
            let isResolved = false;
            
            const overlay = document.createElement('div');
            overlay.className = 'date-format-overlay';
            overlay.style.cssText = `
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(0, 0, 0, 0.7);
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
                max-width: 400px;
                text-align: center;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.3);
            `;
            
            modal.innerHTML = `
                <h3 style="margin: 0 0 20px 0; color: #333;">Choose Date Format</h3>
                <p style="margin: 0 0 25px 0; color: #666; line-height: 1.5;">
                    Your Excel files contain ambiguous dates. Please choose your date format:
                </p>
                <div style="display: flex; gap: 15px; justify-content: center;">
                    <button data-format="DD-MM-YYYY" style="
                        padding: 12px 20px;
                        background: #007bff;
                        color: white;
                        border: none;
                        border-radius: 6px;
                        cursor: pointer;
                        font-size: 14px;
                    ">DD-MM-YYYY<br><small style="opacity: 0.8;">(European)</small></button>
                    <button data-format="MM-DD-YYYY" style="
                        padding: 12px 20px;
                        background: #007bff;
                        color: white;
                        border: none;
                        border-radius: 6px;
                        cursor: pointer;
                        font-size: 14px;
                    ">MM-DD-YYYY<br><small style="opacity: 0.8;">(US)</small></button>
                </div>
            `;
            
            const handleClick = (event) => {
                if (isResolved) return;
                
                const format = event.target.closest('button')?.dataset.format;
                if (!format) return;
                
                isResolved = true;
                this.detectedDateFormat = format;
                
                // Remove all date format overlays
                document.querySelectorAll('.date-format-overlay').forEach(el => el.remove());
                
                resolve(format);
            };
            
            modal.addEventListener('click', handleClick);
            overlay.appendChild(modal);
            document.body.appendChild(overlay);
        });
    }

    /**
     * Parse Portfolio Details Excel file from EquatePlus
     * @param {File} file - Excel file object containing portfolio data
     * @returns {Promise<Object>} Parsed portfolio data with validated entries
     * @throws {Error} If file format is invalid or data is malformed
     */
    async parsePortfolioFile(file) {
        try {
            const data = await this.readExcelFile(file);
            const worksheet = data.Sheets[data.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                raw: true         // Get raw values - numbers as numbers, dates as dates
            });

            // Find the data start row (after headers)
            let dataStartRow = -1;
            for (let i = 0; i < rawData.length; i++) {
                if (rawData[i][0] === 'Allocation date') {
                    dataStartRow = i + 1;
                    break;
                }
            }

            if (dataStartRow === -1) {
                throw new Error('Could not find portfolio data headers');
            }

            // Extract user ID
            const userId = this.extractUserId(rawData);
            if (!userId) {
                throw new Error('Could not find User ID in portfolio file');
            }

            // Extract "As of date" from cell B3
            const asOfDate = this.extractAsOfDate(rawData);

            // Parse portfolio entries
            const portfolioEntries = [];
            for (let i = dataStartRow; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row[0] || row[0] === '') break; // Stop at empty row

                // Skip if this is a calculation row (contains formulas or totals)
                if (typeof row[0] === 'string' && row[0].includes('Total')) {
                    break;
                }

                const entry = this.parsePortfolioRow(row);
                if (entry) {
                    portfolioEntries.push(entry);
                }
            }

            // Final validation - ensure we have valid data (include sold shares with outstanding = 0)
            const validEntries = portfolioEntries.filter(entry => {
                return entry.allocationDate && 
                       entry.allocatedQuantity > 0 && // Must have been allocated shares originally
                       entry.marketPrice > 0;
            });

            if (validEntries.length === 0) {
                throw new Error('No valid portfolio entries found after validation');
            }

            if (validEntries.length !== portfolioEntries.length) {
                config.warn(`⚠️ Filtered out ${portfolioEntries.length - validEntries.length} invalid entries`);
            }

            // Return validated data - GUARANTEED CLEAN
            this.portfolioData = {
                userId: userId,
                asOfDate: asOfDate,
                entries: validEntries,
                uploadDate: new Date().toISOString(),
                fileName: file.name
            };

            config.debug('✅ Portfolio data validated and parsed:', this.portfolioData.entries.length, 'clean entries');
            return this.portfolioData;

        } catch (error) {
            config.error('Error parsing portfolio file:', error);
            throw new Error(`Failed to parse portfolio file: ${error.message}`);
        }
    }

    /**
     * Parse Completed Transactions Excel file from EquatePlus
     * @param {File} file - Excel file object containing transaction history
     * @returns {Promise<Object>} Parsed transaction data with validated entries
     * @throws {Error} If file format is invalid or user ID mismatch
     */
    async parseTransactionFile(file) {
        try {
            const data = await this.readExcelFile(file);
            const worksheet = data.Sheets[data.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                raw: true         // Get raw values - numbers as numbers, dates as dates
            });

            // Find the data start row
            let dataStartRow = -1;
            for (let i = 0; i < rawData.length; i++) {
                if (rawData[i][0] === 'Order reference') {
                    dataStartRow = i + 1;
                    break;
                }
            }

            if (dataStartRow === -1) {
                throw new Error('Could not find transaction data headers');
            }

            // Extract user ID and validate
            const userId = this.extractUserId(rawData);
            if (!userId) {
                throw new Error('Could not find User ID in transaction file');
            }

            // Parse transaction entries
            const transactionEntries = [];
            for (let i = dataStartRow; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row[0] || row[0] === '') break; // Stop at empty row

                const entry = this.parseTransactionRow(row);
                if (entry) {
                    transactionEntries.push(entry);
                }
            }

            // Final validation - ensure we have valid transaction data
            const validTransactions = transactionEntries.filter(transaction => {
                return transaction.transactionDate && 
                       transaction.quantity !== 0 && 
                       transaction.status === 'Executed' && // Only process executed transactions
                       (transaction.executionPrice > 0 || transaction.netProceeds > 0); // Allow dividends with no execution price
            });

            if (validTransactions.length === 0) {
                config.warn('⚠️ No valid transactions found after validation');
            }

            if (validTransactions.length !== transactionEntries.length) {
                config.warn(`⚠️ Filtered out ${transactionEntries.length - validTransactions.length} invalid transactions`);
            }

            // Return validated data - GUARANTEED CLEAN
            this.transactionData = {
                userId: userId,
                entries: validTransactions,
                uploadDate: new Date().toISOString(),
                fileName: file.name
            };

            config.debug('✅ Transaction data validated and parsed:', this.transactionData.entries.length, 'clean transactions');
            
            // Debug: Show March 2025 transactions if any
            const march2025Transactions = validTransactions.filter(t => t.transactionDate && t.transactionDate.includes('2025-03'));
            if (march2025Transactions.length > 0) {
                config.debug('🔍 March 2025 transactions found:', march2025Transactions);
            }
            
            return this.transactionData;

        } catch (error) {
            config.error('Error parsing transaction file:', error);
            throw new Error(`Failed to parse transaction file: ${error.message}`);
        }
    }

    /**
     * Parse historical prices Excel file
     */
    async parseHistoricalPricesFile(file) {
        try {
            const data = await this.readExcelFile(file);
            const worksheet = data.Sheets[data.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                raw: true         // Get raw values - numbers as numbers, dates as dates
            });

            // Skip header row, start from row 1
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
                        date: jsDate.toISOString().split('T')[0], // YYYY-MM-DD format
                        price: price
                    });
                }
            }

            // Sort by date
            priceEntries.sort((a, b) => new Date(a.date) - new Date(b.date));

            this.historicalPrices = priceEntries;
            return this.historicalPrices;

        } catch (error) {
            config.error('Error parsing historical prices file:', error);
            throw new Error(`Failed to parse historical prices file: ${error.message}`);
        }
    }

    /**
     * Read Excel file using SheetJS
     */
    readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { 
                        type: 'array',
                        cellDates: true,  // Parse dates as Date objects instead of strings
                        raw: false        // Don't use raw values for text, but dates will still be Date objects
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
     * Extract User ID from raw Excel data
     */
    extractUserId(rawData) {
        for (let i = 0; i < Math.min(10, rawData.length); i++) {
            const row = rawData[i];
            if (row[0] === 'User ID' && row[1]) {
                return row[1].toString();
            }
        }
        return null;
    }

    /**
     * Extract "As of date" from raw Excel data (typically cell B3)
     */
    extractAsOfDate(rawData) {
        for (let i = 0; i < Math.min(10, rawData.length); i++) {
            const row = rawData[i];
            if (row[0] === 'As of date' && row[1]) {
                try {
                    // Parse the date value (could be Date object or string)
                    const dateValue = row[1];
                    if (dateValue instanceof Date) {
                        return dateValue.toISOString().split('T')[0];
                    } else if (typeof dateValue === 'string') {
                        // Handle timezone strings in various formats - timezone-safe
                        let cleanDateString = dateValue
                            // Remove timezone abbreviations from end: "Aug 1, 2025, 8:40:50 PM CEST"
                            .replace(/\s+(CEST|CET|EST|PST|GMT|UTC)$/, '')
                            // Remove timezone abbreviations from middle: "2025年8月2日 CEST 下午9:50:06"
                            .replace(/\s+(CEST|CET|EST|PST|GMT|UTC)\s+/, ' ');
                        
                        const parsed = new Date(cleanDateString);
                        if (!isNaN(parsed.getTime())) {
                            // Convert to UTC date to ensure consistency across timezones
                            const utcDateString = parsed.toISOString().split('T')[0];
                            const utcDate = new Date(utcDateString + 'T00:00:00.000Z');
                            return utcDate.toISOString().split('T')[0];
                        }
                    } else if (typeof dateValue === 'number') {
                        // Excel serial date number
                        const jsDate = this.excelDateToJSDate(dateValue);
                        return jsDate.toISOString().split('T')[0];
                    }
                } catch (error) {
                    config.warn('⚠️ Could not parse "As of date":', error.message);
                }
            }
        }
        return null;
    }

    /**
     * Parse a single portfolio row
     */
    parsePortfolioRow(row) {
        try {
            return {
                allocationDate: this.parseExcelDate(row[0], 'allocationDate'),
                plan: this.validateString(row[1], 'plan'),
                instrumentType: this.validateString(row[2], 'instrumentType'),
                instrument: this.validateString(row[3], 'instrument'),
                contributionType: this.validateString(row[4], 'contributionType'),
                costBasis: this.validateNumber(row[5], 'costBasis'),
                marketPrice: this.validateNumber(row[6], 'marketPrice'),
                availableFrom: this.parseExcelDate(row[7], 'availableFrom'),
                expiryDate: this.parseExcelDate(row[8], 'expiryDate') || null,
                allocatedQuantity: this.validateNumber(row[9], 'allocatedQuantity'),
                outstandingQuantity: this.validateNumber(row[10], 'outstandingQuantity'),
                availableQuantity: this.validateNumber(row[11], 'availableQuantity'),
                currentOutstandingValue: this.validateNumber(row[12], 'currentOutstandingValue'),
                currentAvailableValue: this.validateNumber(row[13], 'currentAvailableValue'),
                purchaseAmount: this.validateNumber(row[14], 'purchaseAmount')
            };
        } catch (error) {
            config.error('❌ Error parsing portfolio row:', error.message);
            throw error; // Fail fast on bad data
        }
    }

    /**
     * Parse a single transaction row
     */
    parseTransactionRow(row) {
        try {
            return {
                orderReference: this.validateString(row[0], 'orderReference'),
                transactionDate: this.parseExcelDate(row[1], 'transactionDate'),
                orderType: this.validateString(row[2], 'orderType'),
                quantity: this.validateNumber(row[3], 'quantity'),
                status: this.validateString(row[4], 'status'),
                executionPrice: this.validateNumber(row[5], 'executionPrice'),
                instrument: this.validateString(row[6], 'instrument'),
                productType: this.validateString(row[7], 'productType'),
                costBasis: this.validateString(row[8], 'costBasis'),
                taxesWithheld: this.validateNumber(row[9], 'taxesWithheld'),
                fees: this.validateNumber(row[10], 'fees'),
                netProceeds: this.validateNumber(row[11], 'netProceeds')
            };
        } catch (error) {
            config.error('❌ Error parsing transaction row:', error.message);
            throw error; // Fail fast on bad data
        }
    }

    /**
     * Universal Excel date parser - works with any locale since SheetJS gives us Date objects
     * With cellDates: true, we get the actual internal date value (3-3-2025 16:09:25) 
     * instead of the localized display text ("3 mrt 2025")
     */
    parseExcelDateValueRobust(cellValue) {
        if (cellValue === null || cellValue === undefined) {
            return null;
        }
        
        // Handle Excel error values
        if (typeof cellValue === 'object' && cellValue.error !== undefined) {
            config.warn('📊 Excel error in date cell:', cellValue.error);
            return null;
        }
        
        // Handle Excel formula objects
        if (typeof cellValue === 'object' && cellValue.result !== undefined) {
            // Check if formula result is an error
            if (typeof cellValue.result === 'object' && cellValue.result.error !== undefined) {
                config.warn('📊 Excel formula error in date:', cellValue.result.error);
                return null;
            }
            // Recursively parse the formula result
            return this.parseExcelDateValueRobust(cellValue.result);
        }
        
        // Handle Date objects directly (this is what we expect with cellDates: true)
        if (cellValue instanceof Date) {
            if (isNaN(cellValue.getTime())) {
                return null;
            }
            // Convert to UTC to avoid timezone issues
            const utcDateString = cellValue.toISOString().split('T')[0];
            return new Date(utcDateString + 'T00:00:00.000Z');
        }
        
        // Handle Excel serial date numbers (fallback)
        if (typeof cellValue === 'number') {
            if (cellValue < 1 || cellValue > 2958465) { // Valid Excel date range
                return null;
            }
            // Convert Excel serial number to JavaScript date
            const jsDate = this.excelDateToJSDate(cellValue);
            if (isNaN(jsDate.getTime())) {
                return null;
            }
            // Convert to UTC to avoid timezone issues
            const utcDateString = jsDate.toISOString().split('T')[0];
            return new Date(utcDateString + 'T00:00:00.000Z');
        }
        
        // Handle string date values (fallback - shouldn't be needed with cellDates: true)
        if (typeof cellValue === 'string') {
            const trimmed = cellValue.trim();
            if (!trimmed || trimmed.startsWith('#')) {
                return null;
            }
            
            config.warn('📊 Got string date value when expecting Date object:', trimmed);
            // Try universal date parsing as fallback
            let jsDate = this.parseUniversalDateString(trimmed);
            if (!jsDate || isNaN(jsDate.getTime())) {
                return null;
            }
            
            // Convert to UTC to avoid timezone issues
            const utcDateString = jsDate.toISOString().split('T')[0];
            return new Date(utcDateString + 'T00:00:00.000Z');
        }
        
        return null;
    }

    /**
     * Universal string date parser - uses detected date format
     */
    parseUniversalDateString(dateStr) {
        // Strategy 1: Try ISO formats first (most reliable)
        if (dateStr.match(/^\d{4}-\d{2}-\d{2}/)) {
            return new Date(dateStr);
        }
        
        // Strategy 2: Use detected date format for ambiguous dates
        const dateMatch = dateStr.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
        if (dateMatch) {
            const [, first, second, year] = dateMatch;
            
            // Use detected format or fall back to DD-MM if not detected yet
            if (this.detectedDateFormat === 'MM-DD-YYYY') {
                // MM-DD-YYYY format (US)
                const month = parseInt(first);
                const day = parseInt(second);
                return new Date(parseInt(year), month - 1, day);
            } else {
                // DD-MM-YYYY format (European) - default
                const day = parseInt(first);
                const month = parseInt(second);
                return new Date(parseInt(year), month - 1, day);
            }
        }
        
        // Strategy 3: Fall back to native Date parsing
        return new Date(dateStr);
    }

    /**
     * Parse Excel date (can be date object or Excel serial number)
     */
    parseExcelDate(value, fieldName = 'date') {
        if (!value) {
            throw new Error(`Missing ${fieldName} - cannot be empty`);
        }
        
        let parsedDate;
        
        try {
            // Use robust Excel date parsing inspired by I2E common utilities
            parsedDate = this.parseExcelDateValueRobust(value);
            
            if (!parsedDate) {
                throw new Error(`Cannot parse "${value}" as date for ${fieldName}`);
            }
            
            // Final validation - must be reasonable date (except for expiry dates which can be far future)
            const year = parsedDate.getFullYear();
            if (fieldName === 'expiryDate') {
                // Expiry dates can be far in the future, but not before 1900
                if (year < 1900 || year > 9999) {
                    config.warn(`${fieldName} year ${year} is extreme, treating as null`);
                    return null; // Return null for unreasonable expiry dates
                }
            } else {
                // Regular date fields should be within reasonable range
                if (year < 1900 || year > 2100) {
                    throw new Error(`${fieldName} year ${year} is outside reasonable range (1900-2100)`);
                }
            }
            
            // Return standardized YYYY-MM-DD format
            const isoString = parsedDate.toISOString().split('T')[0];
            return isoString;
            
        } catch (error) {
            config.error(`❌ Date parsing failed for ${fieldName}:`, error.message);
            throw new Error(`Invalid ${fieldName}: ${error.message}`);
        }
    }

    /**
     * Validate and sanitize string fields
     */
    validateString(value, fieldName) {
        if (value === null || value === undefined) {
            return ''; // Empty string for missing values
        }
        
        const stringValue = String(value).trim();
        return stringValue;
    }

    /**
     * Validate and sanitize numeric fields
     */
    validateNumber(value, fieldName) {
        if (value === null || value === undefined || value === '') {
            return 0;
        }
        
        const numValue = parseFloat(value);
        if (isNaN(numValue)) {
            throw new Error(`Invalid number for ${fieldName}: "${value}"`);
        }
        
        // Reasonable bounds check for financial data
        if (Math.abs(numValue) > 1000000000) { // 1 billion limit
            throw new Error(`${fieldName} value ${numValue} exceeds reasonable bounds`);
        }
        
        return numValue;
    }

    /**
     * Convert Dutch date format to English
     */
    convertDutchDate(dutchDateStr) {
        const dutchMonths = {
            'januari': 'January', 'jan': 'Jan',
            'februari': 'February', 'feb': 'Feb', 
            'maart': 'March', 'mrt': 'Mar',
            'april': 'April', 'apr': 'Apr',
            'mei': 'May',
            'juni': 'June', 'jun': 'Jun',
            'juli': 'July', 'jul': 'Jul',
            'augustus': 'August', 'aug': 'Aug',
            'september': 'September', 'sep': 'Sep',
            'oktober': 'October', 'okt': 'Oct',
            'november': 'November', 'nov': 'Nov',
            'december': 'December', 'dec': 'Dec'
        };

        let englishDate = dutchDateStr.toLowerCase().trim();
        
        // Replace Dutch months with English equivalents
        Object.keys(dutchMonths).forEach(dutchMonth => {
            const regex = new RegExp(`\\b${dutchMonth}\\b`, 'gi');
            englishDate = englishDate.replace(regex, dutchMonths[dutchMonth]);
        });

        return englishDate;
    }

    /**
     * Convert Excel date serial number to JavaScript Date
     */
    excelDateToJSDate(excelDate) {
        // Excel's epoch is 1900-01-01, but it incorrectly treats 1900 as a leap year
        // So we need to account for this
        const excelEpoch = new Date(1899, 11, 30); // December 30, 1899
        const jsDate = new Date(excelEpoch.getTime() + (excelDate * 24 * 60 * 60 * 1000));
        return jsDate;
    }

    /**
     * Validate that portfolio and transaction files are from the same user
     */
    validateUserIds() {
        if (!this.portfolioData || !this.transactionData) {
            return true; // No validation needed if one is missing
        }

        if (this.portfolioData.userId !== this.transactionData.userId) {
            throw new Error(`User ID mismatch: Portfolio (${this.portfolioData.userId}) vs Transactions (${this.transactionData.userId})`);
        }

        return true;
    }

    /**
     * Get parsed data summary
     */
    getDataSummary() {
        return {
            portfolio: this.portfolioData ? {
                userId: this.portfolioData.userId,
                entriesCount: this.portfolioData.entries.length,
                uploadDate: this.portfolioData.uploadDate,
                fileName: this.portfolioData.fileName
            } : null,
            
            transactions: this.transactionData ? {
                userId: this.transactionData.userId,
                entriesCount: this.transactionData.entries.length,
                uploadDate: this.transactionData.uploadDate,
                fileName: this.transactionData.fileName
            } : null,
            
            historicalPrices: this.historicalPrices ? {
                entriesCount: this.historicalPrices.length,
                dateRange: this.historicalPrices.length > 0 ? {
                    from: this.historicalPrices[0].date,
                    to: this.historicalPrices[this.historicalPrices.length - 1].date
                } : null
            } : null
        };
    }

    /**
     * Detect if the portfolio file is in English by checking raw headers
     */
    async detectLanguage(file) {
        try {
            const data = await this.readExcelFile(file);
            const worksheet = data.Sheets[data.SheetNames[0]];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1,
                raw: true
            });

            config.debug('🌐 Language Detection - checking raw headers...');
            
            // Check for required English headers
            const requiredHeaders = [
                'User ID',
                'As of date', 
                'Allocation date'
            ];
            
            let foundHeaders = 0;
            let headerLocations = {};
            
            // Check first 15 rows for headers
            for (let i = 0; i < Math.min(15, rawData.length); i++) {
                const row = rawData[i];
                if (!row) continue;
                
                for (let j = 0; j < row.length; j++) {
                    const cellValue = row[j];
                    if (typeof cellValue === 'string') {
                        for (const header of requiredHeaders) {
                            if (cellValue === header) {
                                foundHeaders++;
                                headerLocations[header] = `Row ${i+1}, Col ${j+1}`;
                                config.debug(`✅ Found English header: "${header}" at ${headerLocations[header]}`);
                                break;
                            }
                        }
                    }
                }
            }
            
            const isEnglish = foundHeaders >= requiredHeaders.length;
            
            config.debug('🌐 Language Detection Result:', {
                isEnglish: isEnglish,
                foundHeaders: foundHeaders,
                requiredHeaders: requiredHeaders.length,
                headerLocations: headerLocations
            });
            
            if (!isEnglish) {
                config.debug('❌ Missing English headers. Found headers in first 15 rows:');
                for (let i = 0; i < Math.min(15, rawData.length); i++) {
                    const row = rawData[i];
                    if (row && row.length > 0) {
                        config.debug(`  Row ${i+1}:`, row.filter(cell => typeof cell === 'string' && cell.trim() !== ''));
                    }
                }
            }

            return isEnglish;
        } catch (error) {
            config.error('❌ Error during language detection:', error);
            return false;
        }
    }

    /**
     * Detect the company based on instrument field
     */
    detectCompany(portfolioData) {
        if (!portfolioData || !portfolioData.entries || portfolioData.entries.length === 0) {
            config.warn('⚠️ No portfolio data available for company detection');
            return 'Other';
        }

        try {
            // Check if ALL portfolio entries have "Allianz share" as instrument
            config.debug('🔍 Checking instruments for company detection:');
            portfolioData.entries.forEach((entry, index) => {
                const instrument = entry.instrument || '';
                const isAllianz = instrument.toLowerCase().includes('allianz share');
                config.debug(`  Entry ${index}: "${instrument}" -> isAllianz: ${isAllianz}`);
            });
            
            const allHaveAllianzShare = portfolioData.entries.every(entry => 
                entry.instrument && 
                entry.instrument.toLowerCase().includes('allianz share')
            );

            // Default to 'Other', only set to 'Allianz' if ALL entries are Allianz shares
            const company = allHaveAllianzShare ? 'Allianz' : 'Other';
            
            config.debug('🏢 Company Detection Result:', {
                company: company,
                allHaveAllianzShare: allHaveAllianzShare,
                totalEntries: portfolioData.entries.length,
                instruments: portfolioData.entries.map(entry => entry.instrument).filter(Boolean)
            });

            return company;
        } catch (error) {
            config.error('❌ Error during company detection:', error);
            return 'Other';
        }
    }

    /**
     * Clear all parsed data
     */
    clearData() {
        this.portfolioData = null;
        this.transactionData = null;
        this.historicalPrices = null;
    }
}

// Global file parser instance
const fileParser = new FileParser();