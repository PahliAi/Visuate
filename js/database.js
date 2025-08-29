/**
 * IndexedDB Database Manager for Equate Portfolio Analysis
 * Handles local storage of portfolio data, historical prices, and user preferences
 */

class EquateDB {
    /**
     * Create a new EquateDB instance
     * @constructor
     */
    constructor() {
        this.dbName = 'EquatePortfolio';
        this.version = 5; // Increased for multi-currency price storage support
        this.db = null;
    }

    /**
     * Initialize the IndexedDB database with required object stores
     * @returns {Promise<EquateDB>} Promise that resolves to the database instance
     * @throws {Error} If IndexedDB is not supported or database fails to open
     */
    async init() {
        return new Promise((resolve, reject) => {
            config.debug(`🗄️ Initializing IndexedDB: ${this.dbName} v${this.version}`);
            
            if (!window.indexedDB) {
                const error = new Error('IndexedDB not supported in this browser');
                config.error('❌', error.message);
                reject(error);
                return;
            }

            const request = indexedDB.open(this.dbName, this.version);

            request.onerror = () => {
                config.error('❌ Database failed to open:', request.error);
                reject(request.error);
            };

            request.onsuccess = () => {
                this.db = request.result;
                config.debug('✅ Database opened successfully');
                
                // Add error handler for the database connection
                this.db.onerror = (event) => {
                    config.error('❌ Database error:', event.target.error);
                };
                
                resolve(this.db);
            };

            request.onupgradeneeded = (e) => {
                this.db = e.target.result;

                // Portfolio data store
                if (!this.db.objectStoreNames.contains('portfolioData')) {
                    const portfolioStore = this.db.createObjectStore('portfolioData', { keyPath: 'id' });
                    portfolioStore.createIndex('userId', 'userId', { unique: false });
                    portfolioStore.createIndex('uploadDate', 'uploadDate', { unique: false });
                }

                // Historical prices store
                if (!this.db.objectStoreNames.contains('historicalPrices')) {
                    const pricesStore = this.db.createObjectStore('historicalPrices', { keyPath: 'date' });
                }

                // User preferences store
                if (!this.db.objectStoreNames.contains('userPreferences')) {
                    this.db.createObjectStore('userPreferences', { keyPath: 'key' });
                }

                // Manual price entries store
                if (!this.db.objectStoreNames.contains('manualPrices')) {
                    const manualPricesStore = this.db.createObjectStore('manualPrices', { keyPath: 'timestamp' });
                }

                // Historical currency data store
                if (!this.db.objectStoreNames.contains('historicalCurrencyData')) {
                    const currencyStore = this.db.createObjectStore('historicalCurrencyData', { keyPath: 'date' });
                    config.debug('Created historicalCurrencyData object store');
                }

                // Currency quality scores store
                if (!this.db.objectStoreNames.contains('currencyQualityScores')) {
                    const qualityStore = this.db.createObjectStore('currencyQualityScores', { keyPath: 'currency' });
                    config.debug('Created currencyQualityScores object store');
                }

                // Multi-currency prices store (Sprint 3: Performance optimization)
                if (!this.db.objectStoreNames.contains('multiCurrencyPrices')) {
                    const multiPricesStore = this.db.createObjectStore('multiCurrencyPrices', { keyPath: 'date' });
                    config.debug('Created multiCurrencyPrices object store for instant currency switching');
                }

                config.debug('Database setup complete');
            };
        });
    }

    /**
     * Save portfolio data to IndexedDB
     */
    async savePortfolioData(portfolioData, transactionData = null, userId, isEnglish = true, company = null, currency = null, transactionCompany = null) {
        const transaction = this.db.transaction(['portfolioData'], 'readwrite');
        const store = transaction.objectStore('portfolioData');

        // Add detectedInfo and stats to portfolioData if not already present
        if (portfolioData && !portfolioData.detectedInfo) {
            portfolioData.detectedInfo = {
                originalFilename: portfolioData.fileName || 'Portfolio File',
                language: (typeof translationManager !== 'undefined') ? translationManager.getCurrentLanguage() : 'english',
                currency: currency,
                company: company,
                dateFormat: (typeof fileParser !== 'undefined') ? fileParser.detectedDateFormat || 'DD-MM-YYYY' : 'DD-MM-YYYY'
            };
            portfolioData.stats = {
                entryCount: portfolioData.entries?.length || 0
            };
        }

        // Add detectedInfo and stats to transactionData if present and not already there
        if (transactionData && !transactionData.detectedInfo) {
            // Use passed transaction company (pre-detected by FileAnalyzer) or default to Other
            let detectedTransactionCompany = transactionCompany || 'Other';

            transactionData.detectedInfo = {
                originalFilename: transactionData.fileName || 'Transaction File', 
                language: (typeof translationManager !== 'undefined') ? translationManager.getCurrentLanguage() : 'english',
                currency: transactionData.currency || currency,
                dateFormat: (typeof fileParser !== 'undefined') ? fileParser.detectedDateFormat || 'DD-MM-YYYY' : 'DD-MM-YYYY',
                company: detectedTransactionCompany,
                asOfDate: transactionData.asOfDate || null
            };
            transactionData.stats = {
                entryCount: transactionData.entries?.length || 0
            };
        }

        const data = {
            id: 'current', // Only store one current portfolio
            userId: userId,
            portfolioData: portfolioData,
            transactionData: transactionData,
            uploadDate: new Date().toISOString(),
            hasTransactions: transactionData !== null,
            isEnglish: isEnglish,
            company: company,
            currency: currency // Store the detected currency
        };

        return store.put(data);
    }

    /**
     * Get portfolio data from IndexedDB
     */
    async getPortfolioData() {
        const transaction = this.db.transaction(['portfolioData'], 'readonly');
        const store = transaction.objectStore('portfolioData');
        
        return new Promise((resolve, reject) => {
            const request = store.get('current');
            
            request.onsuccess = () => {
                resolve(request.result);
            };
            
            request.onerror = () => {
                reject(request.error);
            };
        });
    }

    /**
     * Save historical prices to IndexedDB
     */
    async saveHistoricalPrices(pricesArray) {
        return new Promise((resolve, reject) => {
            const transaction = this.db.transaction(['historicalPrices'], 'readwrite');
            const store = transaction.objectStore('historicalPrices');

            // Clear existing prices first (in same transaction)
            const clearRequest = store.clear();
            
            clearRequest.onsuccess = () => {
                // Add all price entries in optimized batches to avoid overwhelming the transaction
                const batchSize = 100; // Process 100 entries at a time for optimal performance
                let batchIndex = 0;
                let completed = 0;
                const total = pricesArray.length;
                
                if (total === 0) {
                    resolve();
                    return;
                }
                
                const processBatch = () => {
                    const batchStart = batchIndex * batchSize;
                    const batchEnd = Math.min(batchStart + batchSize, total);
                    const batch = pricesArray.slice(batchStart, batchEnd);
                    
                    if (batch.length === 0) {
                        // All batches completed successfully
                        resolve();
                        return;
                    }
                    
                    let batchCompleted = 0;
                    batch.forEach(priceEntry => {
                        const addRequest = store.add({
                            date: priceEntry.date,
                            price: priceEntry.price
                        });
                        
                        addRequest.onsuccess = () => {
                            batchCompleted++;
                            completed++;
                            
                            if (batchCompleted === batch.length) {
                                // Current batch complete, process next batch immediately
                                batchIndex++;
                                processBatch();
                            }
                        };
                        
                        addRequest.onerror = () => {
                            config.error('Error adding price entry:', addRequest.error);
                            reject(addRequest.error);
                        };
                    });
                };
                
                processBatch();
            };
            
            clearRequest.onerror = () => {
                config.error('Error clearing historical prices:', clearRequest.error);
                reject(clearRequest.error);
            };

            transaction.onerror = () => {
                config.error('Historical prices transaction failed:', transaction.error);
                reject(transaction.error);
            };
        });
    }

    /**
     * Append single historical price entry (e.g., Portfolio AsOfDate)
     * Only adds if date doesn't already exist
     */
    async appendHistoricalPrice(date, price) {
        return new Promise((resolve, reject) => {
            const transaction = this.db.transaction(['historicalPrices'], 'readwrite');
            const store = transaction.objectStore('historicalPrices');

            // Check if date already exists
            const getRequest = store.get(date);
            
            getRequest.onsuccess = () => {
                if (getRequest.result) {
                    // Date already exists, update if price is different
                    if (getRequest.result.price !== price) {
                        const updateRequest = store.put({
                            date: date,
                            price: price
                        });
                        
                        updateRequest.onsuccess = () => {
                            config.debug(`📈 Updated historical price for ${date}: €${price}`);
                            resolve();
                        };
                        
                        updateRequest.onerror = () => {
                            config.error('Error updating historical price:', updateRequest.error);
                            reject(updateRequest.error);
                        };
                    } else {
                        // Same price, no need to update
                        resolve();
                    }
                } else {
                    // Date doesn't exist, add new entry
                    const addRequest = store.add({
                        date: date,
                        price: price
                    });
                    
                    addRequest.onsuccess = () => {
                        config.debug(`📈 Added new historical price for ${date}: €${price}`);
                        resolve();
                    };
                    
                    addRequest.onerror = () => {
                        config.error('Error adding historical price:', addRequest.error);
                        reject(addRequest.error);
                    };
                }
            };
            
            getRequest.onerror = () => {
                config.error('Error checking existing historical price:', getRequest.error);
                reject(getRequest.error);
            };

            transaction.onerror = () => {
                config.error('Append historical price transaction failed:', transaction.error);
                reject(transaction.error);
            };
        });
    }

    /**
     * Save multi-currency prices to IndexedDB for instant currency switching
     * Each entry contains all currency prices for a specific date
     * @param {Array} multiCurrencyArray - Array of objects like {date, EUR: 284.50, USD: 295.88, GBP: 236.14, ...}
     * @returns {Promise} Promise that resolves when all prices are saved
     */
    async saveMultiCurrencyPrices(multiCurrencyArray) {
        return new Promise((resolve, reject) => {
            const transaction = this.db.transaction(['multiCurrencyPrices'], 'readwrite');
            const store = transaction.objectStore('multiCurrencyPrices');

            // Clear existing prices first (in same transaction)
            const clearRequest = store.clear();
            
            clearRequest.onsuccess = () => {
                // Add all multi-currency entries in optimized batches
                const batchSize = 100;
                let batchIndex = 0;
                let completed = 0;
                const total = multiCurrencyArray.length;
                
                if (total === 0) {
                    resolve();
                    return;
                }

                const processBatch = () => {
                    const startIndex = batchIndex * batchSize;
                    const endIndex = Math.min(startIndex + batchSize, total);
                    const batch = multiCurrencyArray.slice(startIndex, endIndex);
                    
                    if (batch.length === 0) {
                        resolve();
                        return;
                    }
                    
                    let batchCompleted = 0;
                    batch.forEach(multiCurrencyEntry => {
                        const addRequest = store.add(multiCurrencyEntry);
                        
                        addRequest.onsuccess = () => {
                            batchCompleted++;
                            completed++;
                            
                            if (batchCompleted === batch.length) {
                                if (completed === total) {
                                    config.debug(`✅ Multi-currency prices saved: ${total} entries with ${Object.keys(multiCurrencyArray[0]).length - 1} currencies each`);
                                    resolve();
                                } else {
                                    batchIndex++;
                                    processBatch();
                                }
                            }
                        };
                        
                        addRequest.onerror = () => {
                            config.error('Error adding multi-currency price:', addRequest.error);
                            reject(addRequest.error);
                        };
                    });
                };
                
                processBatch();
            };

            clearRequest.onerror = () => {
                config.error('Error clearing multi-currency prices:', clearRequest.error);
                reject(clearRequest.error);
            };

            transaction.onerror = () => {
                config.error('Multi-currency price transaction failed:', transaction.error);
                reject(transaction.error);
            };
        });
    }

    /**
     * Get multi-currency prices from IndexedDB
     * Returns array of objects with all currency prices for each date
     * @returns {Promise<Array>} Array of {date, EUR, USD, GBP, ARS, ...} objects
     */
    async getMultiCurrencyPrices() {
        const transaction = this.db.transaction(['multiCurrencyPrices'], 'readonly');
        const store = transaction.objectStore('multiCurrencyPrices');
        
        return new Promise((resolve, reject) => {
            const getAllRequest = store.getAll();
            
            getAllRequest.onsuccess = () => {
                const prices = getAllRequest.result || [];
                // Sort by date for consistent timeline order
                prices.sort((a, b) => new Date(a.date) - new Date(b.date));
                
                config.debug(`📊 Retrieved ${prices.length} multi-currency price entries from database`);
                if (prices.length > 0) {
                    const availableCurrencies = Object.keys(prices[0]).filter(key => key !== 'date');
                    config.debug(`💰 Available currencies: ${availableCurrencies.join(', ')}`);
                }
                resolve(prices);
            };
            
            getAllRequest.onerror = () => {
                config.error('Error retrieving multi-currency prices:', getAllRequest.error);
                reject(getAllRequest.error);
            };
        });
    }

    /**
     * Clear all multi-currency prices from IndexedDB
     * @returns {Promise} Promise that resolves when prices are cleared
     */
    async clearMultiCurrencyPrices() {
        return new Promise((resolve, reject) => {
            const transaction = this.db.transaction(['multiCurrencyPrices'], 'readwrite');
            const store = transaction.objectStore('multiCurrencyPrices');
            
            const clearRequest = store.clear();
            
            clearRequest.onsuccess = () => {
                config.debug('🗑️ Multi-currency prices cleared');
                resolve();
            };
            
            clearRequest.onerror = () => {
                config.error('Error clearing multi-currency prices:', clearRequest.error);
                reject(clearRequest.error);
            };
        });
    }

    /**
     * Get historical prices from IndexedDB
     * For currency-specific requests, uses multiCurrencyPrices for efficiency
     * For backward compatibility, falls back to historicalPrices when no currency specified
     */
    async getHistoricalPrices(currency = null) {
        // If currency specified, use multi-currency data for efficiency
        if (currency) {
            try {
                const multiCurrencyData = await this.getMultiCurrencyPrices();
                
                if (multiCurrencyData && multiCurrencyData.length > 0) {
                    // Extract prices for the requested currency, preserving source information
                    const currencyPrices = multiCurrencyData
                        .filter(entry => entry[currency] !== undefined && entry[currency] !== null)
                        .map(entry => ({
                            date: entry.date,
                            price: entry[currency],
                            // Preserve source property if it exists (for AsOfDate green diamonds)
                            ...(entry.source && { source: entry.source })
                        }))
                        .sort((a, b) => new Date(a.date) - new Date(b.date));
                    
                    config.debug(`📈 Retrieved ${currencyPrices.length} ${currency} prices from multi-currency data`);
                    return currencyPrices;
                }
                
                // Fallback to old method if no multi-currency data
                config.warn('No multi-currency data found, falling back to historicalPrices store');
            } catch (error) {
                config.error('Error accessing multi-currency prices, falling back to historicalPrices:', error);
            }
        }

        // Backward compatibility: use historicalPrices store (no currency filtering since data has no currency property)
        const transaction = this.db.transaction(['historicalPrices'], 'readonly');
        const store = transaction.objectStore('historicalPrices');
        
        return new Promise((resolve, reject) => {
            const request = store.getAll();
            
            request.onsuccess = () => {
                const allPrices = request.result;
                
                // Sort by date for consistency
                allPrices.sort((a, b) => new Date(a.date) - new Date(b.date));
                
                if (currency) {
                    config.warn(`⚠️ Currency filtering requested (${currency}) but historicalPrices has no currency property. Returning all prices.`);
                }
                
                resolve(allPrices);
            };
            
            request.onerror = () => {
                reject(request.error);
            };
        });
    }

    /**
     * Get latest historical price
     */
    async getLatestPrice() {
        const transaction = this.db.transaction(['historicalPrices'], 'readonly');
        const store = transaction.objectStore('historicalPrices');
        
        return new Promise((resolve, reject) => {
            // Get all prices and find the latest one
            const request = store.getAll();
            
            request.onsuccess = () => {
                const prices = request.result;
                if (prices.length === 0) {
                    resolve(null);
                    return;
                }
                
                // Sort by date and get the latest
                prices.sort((a, b) => new Date(b.date) - new Date(a.date));
                resolve(prices[0]);
            };
            
            request.onerror = () => {
                reject(request.error);
            };
        });
    }

    /**
     * Save user preference
     */
    async savePreference(key, value) {
        const transaction = this.db.transaction(['userPreferences'], 'readwrite');
        const store = transaction.objectStore('userPreferences');

        return store.put({ key: key, value: value });
    }

    /**
     * Get user preference
     */
    async getPreference(key) {
        const transaction = this.db.transaction(['userPreferences'], 'readonly');
        const store = transaction.objectStore('userPreferences');
        
        return new Promise((resolve, reject) => {
            const request = store.get(key);
            
            request.onsuccess = () => {
                resolve(request.result ? request.result.value : null);
            };
            
            request.onerror = () => {
                reject(request.error);
            };
        });
    }

    /**
     * Save manual price entry
     */
    async saveManualPrice(price) {
        const transaction = this.db.transaction(['manualPrices'], 'readwrite');
        const store = transaction.objectStore('manualPrices');

        const data = {
            timestamp: Date.now(),
            price: price,
            date: new Date().toISOString()
        };

        return store.put(data);
    }

    /**
     * Get latest manual price
     */
    async getLatestManualPrice() {
        const transaction = this.db.transaction(['manualPrices'], 'readonly');
        const store = transaction.objectStore('manualPrices');
        
        return new Promise((resolve, reject) => {
            const request = store.getAll();
            
            request.onsuccess = () => {
                const prices = request.result;
                if (prices.length === 0) {
                    resolve(null);
                    return;
                }
                
                // Sort by timestamp and get the latest
                prices.sort((a, b) => b.timestamp - a.timestamp);
                resolve(prices[0]);
            };
            
            request.onerror = () => {
                reject(request.error);
            };
        });
    }

    /**
     * Clear all manual price entries
     */
    async clearManualPrices() {
        const transaction = this.db.transaction(['manualPrices'], 'readwrite');
        const store = transaction.objectStore('manualPrices');
        return store.clear();
    }

    /**
     * Clear currency quality scores
     */
    async clearCurrencyQualityScores() {
        if (!this.db.objectStoreNames.contains('currencyQualityScores')) {
            return; // Store doesn't exist, nothing to clear
        }
        const transaction = this.db.transaction(['currencyQualityScores'], 'readwrite');
        const store = transaction.objectStore('currencyQualityScores');
        return store.clear();
    }

    /**
     * Clear all portfolio data
     */
    async clearPortfolioData() {
        const transaction = this.db.transaction(['portfolioData'], 'readwrite');
        const store = transaction.objectStore('portfolioData');
        return store.clear();
    }

    /**
     * Clear historical prices
     */
    async clearHistoricalPrices() {
        const transaction = this.db.transaction(['historicalPrices'], 'readwrite');
        const store = transaction.objectStore('historicalPrices');
        return store.clear();
    }

    /**
     * Clear only transaction data from cached portfolio data
     */
    async clearTransactionData() {
        const cachedData = await this.getPortfolioData();
        if (cachedData && cachedData.portfolioData) {
            // Keep portfolio data, clear transaction data
            await this.savePortfolioData(
                cachedData.portfolioData, 
                null, // Clear transaction data
                cachedData.userId,
                cachedData.isEnglish,
                cachedData.company,
                cachedData.currency  // Preserve the currency
            );
            config.debug('✅ Cleared transaction data from cache, kept portfolio data');
        }
    }

    /**
     * Clear all data
     */
    async clearAllData() {
        const promises = [
            this.clearPortfolioData(),
            this.clearHistoricalPrices(),
            this.clearMultiCurrencyPrices(),
            this.clearHistoricalCurrencyData()
        ];
        
        return Promise.all(promises);
    }

    /**
     * Get database info/statistics
     */
    async getDatabaseInfo() {
        try {
            const portfolioData = await this.getPortfolioData();
            const pricesCount = await this.getHistoricalPricesCount();
            
            return {
                hasPortfolioData: !!portfolioData,
                lastUpload: portfolioData ? portfolioData.uploadDate : null,
                userId: portfolioData ? portfolioData.userId : null,
                historicalPricesCount: pricesCount,
                hasTransactions: portfolioData ? !!(portfolioData.hasTransactions && portfolioData.transactionData) : false
            };
        } catch (error) {
            config.error('Error getting database info:', error);
            return null;
        }
    }

    /**
     * Get transaction data from cache
     */
    async getTransactionData() {
        try {
            const portfolioData = await this.getPortfolioData();
            return portfolioData ? portfolioData.transactionData : null;
        } catch (error) {
            config.error('Error getting transaction data:', error);
            return null;
        }
    }

    /**
     * Get count of historical prices
     */
    async getHistoricalPricesCount() {
        const transaction = this.db.transaction(['historicalPrices'], 'readonly');
        const store = transaction.objectStore('historicalPrices');
        
        return new Promise((resolve, reject) => {
            const request = store.count();
            
            request.onsuccess = () => {
                resolve(request.result);
            };
            
            request.onerror = () => {
                reject(request.error);
            };
        });
    }

    /**
     * Save historical currency data and quality scores to the database
     * @param {Array} currencyDataArray - Array of currency data objects with date and currency rates
     * @param {Object} qualityScores - Currency quality scores object (optional)
     * @returns {Promise<void>}
     */
    async saveHistoricalCurrencyData(currencyDataArray, qualityScores = null) {
        return new Promise((resolve, reject) => {
            // Check if currencyQualityScores store exists before including it in transaction
            const hasQualityStore = this.db.objectStoreNames.contains('currencyQualityScores');
            const storeNames = (qualityScores && hasQualityStore) ? ['historicalCurrencyData', 'currencyQualityScores'] : ['historicalCurrencyData'];
            const transaction = this.db.transaction(storeNames, 'readwrite');
            const store = transaction.objectStore('historicalCurrencyData');

            // Clear existing currency data first (in same transaction)
            const clearRequest = store.clear();
            
            clearRequest.onsuccess = () => {
                // Add all currency entries first
                let currencyCompleted = 0;
                const currencyTotal = currencyDataArray.length;
                
                if (currencyTotal === 0) {
                    resolve();
                    return;
                }
                
                // Save currency data in optimized batches (reduces IndexedDB overhead)
                const batchSize = 100; // Process 100 entries at a time for optimal performance
                let batchIndex = 0;
                
                const processBatch = () => {
                    const batchStart = batchIndex * batchSize;
                    const batchEnd = Math.min(batchStart + batchSize, currencyTotal);
                    const batch = currencyDataArray.slice(batchStart, batchEnd);
                    
                    if (batch.length === 0) {
                        // All batches completed successfully
                        if (qualityScores && hasQualityStore) {
                            this.saveCurrencyQualityScores(qualityScores, transaction, resolve, reject);
                        } else {
                            if (qualityScores && !hasQualityStore) {
                                config.warn('Cannot save currency quality scores: currencyQualityScores store does not exist (database needs upgrade)');
                            }
                            resolve();
                        }
                        return;
                    }
                    
                    let batchCompleted = 0;
                    batch.forEach(currencyEntry => {
                        const addRequest = store.add(currencyEntry);
                        
                        addRequest.onsuccess = () => {
                            batchCompleted++;
                            currencyCompleted++;
                            
                            if (batchCompleted === batch.length) {
                                // Current batch complete, process next batch immediately
                                batchIndex++;
                                processBatch();
                            }
                        };
                        
                        addRequest.onerror = () => {
                            config.error('Error adding currency entry:', addRequest.error);
                            reject(addRequest.error);
                        };
                    });
                };
                
                processBatch();
            };
            
            clearRequest.onerror = () => {
                config.error('Error clearing historical currency data:', clearRequest.error);
                reject(clearRequest.error);
            };

            transaction.onerror = () => {
                config.error('Historical currency data transaction failed:', transaction.error);
                reject(transaction.error);
            };
        });
    }

    /**
     * Save currency quality scores (helper method for saveHistoricalCurrencyData)
     * @param {Object} qualityScores - Currency quality scores object
     * @param {IDBTransaction} transaction - Existing transaction
     * @param {Function} resolve - Promise resolve function
     * @param {Function} reject - Promise reject function
     */
    saveCurrencyQualityScores(qualityScores, transaction, resolve, reject) {
        const qualityStore = transaction.objectStore('currencyQualityScores');
        
        // Clear existing quality scores first
        const clearQualityRequest = qualityStore.clear();
        
        clearQualityRequest.onsuccess = () => {
            const currencyEntries = Object.entries(qualityScores);
            let qualityCompleted = 0;
            const qualityTotal = currencyEntries.length;
            
            if (qualityTotal === 0) {
                resolve();
                return;
            }
            
            // Save each currency quality score
            currencyEntries.forEach(([currency, scoreData]) => {
                const qualityEntry = {
                    currency: currency,
                    ...scoreData
                };
                
                const addQualityRequest = qualityStore.add(qualityEntry);
                
                addQualityRequest.onsuccess = () => {
                    qualityCompleted++;
                    if (qualityCompleted === qualityTotal) {
                        config.debug(`📊 Saved ${qualityTotal} currency quality scores to database`);
                        resolve();
                    }
                };
                
                addQualityRequest.onerror = () => {
                    config.error('Error adding currency quality score:', addQualityRequest.error);
                    reject(addQualityRequest.error);
                };
            });
        };
        
        clearQualityRequest.onerror = () => {
            config.error('Error clearing currency quality scores:', clearQualityRequest.error);
            reject(clearQualityRequest.error);
        };
    }

    /**
     * Get all historical currency data from the database
     * @returns {Promise<Array>} Array of currency data objects
     */
    async getHistoricalCurrencyData() {
        const transaction = this.db.transaction(['historicalCurrencyData'], 'readonly');
        const store = transaction.objectStore('historicalCurrencyData');
        
        return new Promise((resolve, reject) => {
            const request = store.getAll();
            
            request.onsuccess = () => {
                resolve(request.result || []);
            };
            
            request.onerror = () => {
                config.error('Error getting historical currency data:', request.error);
                reject(request.error);
            };
        });
    }

    /**
     * Get count of historical currency data entries
     */
    async getHistoricalCurrencyDataCount() {
        const transaction = this.db.transaction(['historicalCurrencyData'], 'readonly');
        const store = transaction.objectStore('historicalCurrencyData');
        
        return new Promise((resolve, reject) => {
            const request = store.count();
            
            request.onsuccess = () => {
                resolve(request.result);
            };
            
            request.onerror = () => {
                reject(request.error);
            };
        });
    }

    /**
     * Clear historical currency data and quality scores
     * @returns {Promise<void>}
     */
    async clearHistoricalCurrencyData() {
        return new Promise((resolve, reject) => {
            const storeNames = ['historicalCurrencyData'];
            
            // Include currency quality scores store if it exists
            if (this.db.objectStoreNames.contains('currencyQualityScores')) {
                storeNames.push('currencyQualityScores');
            }
            
            const transaction = this.db.transaction(storeNames, 'readwrite');
            
            let completed = 0;
            const total = storeNames.length;
            
            storeNames.forEach(storeName => {
                const store = transaction.objectStore(storeName);
                const clearRequest = store.clear();
                
                clearRequest.onsuccess = () => {
                    completed++;
                    config.debug(`✅ Cleared ${storeName} store`);
                    if (completed === total) {
                        resolve();
                    }
                };
                
                clearRequest.onerror = () => {
                    config.error(`Error clearing ${storeName}:`, clearRequest.error);
                    reject(clearRequest.error);
                };
            });
            
            transaction.onerror = () => {
                config.error('Currency data clearing transaction failed:', transaction.error);
                reject(transaction.error);
            };
        });
    }

    /**
     * Get all currency quality scores from the database
     * @returns {Promise<Object>} Currency quality scores object (currency -> scoreData)
     */
    async getCurrencyQualityScores() {
        try {
            // Check if the object store exists (database may be old version)
            if (!this.db.objectStoreNames.contains('currencyQualityScores')) {
                config.warn('currencyQualityScores object store does not exist - database needs upgrade');
                return {};
            }
            
            const transaction = this.db.transaction(['currencyQualityScores'], 'readonly');
            const store = transaction.objectStore('currencyQualityScores');
            
            return new Promise((resolve, reject) => {
                const request = store.getAll();
                
                request.onsuccess = () => {
                    const results = request.result || [];
                    
                    // Convert array back to object format
                    const qualityScores = {};
                    results.forEach(entry => {
                        const { currency, ...scoreData } = entry;
                        qualityScores[currency] = scoreData;
                    });
                    
                    resolve(qualityScores);
                };
                
                request.onerror = () => {
                    config.error('Error getting currency quality scores:', request.error);
                    reject(request.error);
                };
            });
        } catch (error) {
            config.error('Error accessing currency quality scores:', error);
            return {};
        }
    }
}

// Global database instance
const equateDB = new EquateDB();