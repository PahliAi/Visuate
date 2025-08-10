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
        this.version = 1;
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

                config.debug('Database setup complete');
            };
        });
    }

    /**
     * Save portfolio data to IndexedDB
     */
    async savePortfolioData(portfolioData, transactionData = null, userId, isEnglish = true, company = 'Allianz', currency = null) {
        const transaction = this.db.transaction(['portfolioData'], 'readwrite');
        const store = transaction.objectStore('portfolioData');

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
                // Add all price entries in batches to avoid overwhelming the transaction
                let completed = 0;
                const total = pricesArray.length;
                
                if (total === 0) {
                    resolve();
                    return;
                }
                
                pricesArray.forEach(priceEntry => {
                    const addRequest = store.add({
                        date: priceEntry.date,
                        price: priceEntry.price
                    });
                    
                    addRequest.onsuccess = () => {
                        completed++;
                        if (completed === total) {
                            resolve();
                        }
                    };
                    
                    addRequest.onerror = () => {
                        config.error('Error adding price entry:', addRequest.error);
                        reject(addRequest.error);
                    };
                });
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
     * Get historical prices from IndexedDB
     */
    async getHistoricalPrices() {
        const transaction = this.db.transaction(['historicalPrices'], 'readonly');
        const store = transaction.objectStore('historicalPrices');
        
        return new Promise((resolve, reject) => {
            const request = store.getAll();
            
            request.onsuccess = () => {
                resolve(request.result);
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
            this.clearHistoricalPrices()
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
                hasTransactions: portfolioData ? portfolioData.hasTransactions : false
            };
        } catch (error) {
            config.error('Error getting database info:', error);
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
}

// Global database instance
const equateDB = new EquateDB();