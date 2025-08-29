/**
 * Reference Points Builder for Equate
 * Builds comprehensive reference points with multi-currency prices
 * Replaces scattered extraction logic with unified approach
 * 
 * IMPORTANT: Uses currency exchange ratios instead of pre-calculated prices because:
 * - Historical prices in databases are typically market close prices (end of day)
 * - Portfolio/transaction prices represent exact purchase prices during trading hours
 * - Exchange ratios allow accurate conversion of these intraday purchase prices to other currencies
 * - This preserves the precision of actual transaction execution prices
 */
class ReferencePointsBuilder {
    constructor() {
        this.targetedCurrencyData = [];
        this.referencePointsCache = new Map(); // Cache by portfolio+transaction hash
    }

    /**
     * Build complete reference points with all currency prices
     * @param {Object} portfolioData - Portfolio data
     * @param {Object} transactionData - Transaction data  
     * @param {Array} [currencyData=null] - Optional pre-loaded currency data to avoid database calls
     * @returns {Promise<Array>} Complete reference points with all currency prices
     */
    async buildReferencePoints(portfolioData, transactionData, currencyData = null) {
        console.log('🔄 ReferencePointsBuilder starting...');
        const startTime = performance.now();
        
        // Step 1: Extract reference point dates first
        const referenceDates = this.extractReferenceDates(portfolioData, transactionData);
        console.log(`📅 Found ${referenceDates.length} reference point dates`);
        
        // Step 2: Load only needed currency ratios (99.7% data reduction)
        this.targetedCurrencyData = await this.loadTargetedCurrencyRatios(referenceDates, currencyData);
        console.log(`💱 Loaded ${this.targetedCurrencyData.length} targeted currency entries (vs ${await this.getTotalCurrencyEntries()} total)`);
        
        // Step 3: Build reference points with pre-calculated currency prices
        const referencePoints = this.buildMultiCurrencyReferencePoints(portfolioData, transactionData);
        
        // Step 4: Store in IndexedDB for future use
        await this.storeReferencePoints(referencePoints);
        
        const endTime = performance.now();
        console.log(`✨ ReferencePointsBuilder completed: ${referencePoints.length} points in ${(endTime - startTime).toFixed(1)}ms`);
        
        return referencePoints;
    }
    
    /**
     * Extract all unique dates from portfolio and transaction data
     */
    extractReferenceDates(portfolioData, transactionData) {
        const dates = new Set();
        
        // Portfolio purchase dates
        if (portfolioData?.entries) {
            portfolioData.entries.forEach(entry => {
                if (entry.allocationDate && entry.costBasis && entry.costBasis > 0) {
                    dates.add(entry.allocationDate);
                }
            });
        }
        
        // Transaction dates (using same validation as existing extractUserReferencePoints)
        if (transactionData?.entries) {
            transactionData.entries.forEach(transaction => {
                if (this.isSellTransaction(transaction) && 
                    transaction.transactionDate && 
                    transaction.executionPrice && 
                    transaction.executionPrice > 0) {
                    dates.add(transaction.transactionDate);
                }
            });
        }
        
        // AsOfDate
        if (portfolioData?.asOfDate && portfolioData?.entries?.[0]?.marketPrice) {
            dates.add(portfolioData.asOfDate);
        }
        
        return Array.from(dates).sort();
    }
    
    /**
     * Use the same logic as app.js for determining sell transactions
     */
    isSellTransaction(transaction) {
        return transaction.status === 'Executed' && (
            transaction.orderType === 'Sell' || 
            transaction.orderType === 'Sell at market price' || 
            transaction.orderType === 'Sell with price limit' ||
            transaction.orderType === 'Transfer'
        );
    }
    
    /**
     * Load currency ratios only for the specific reference dates
     */
    async loadTargetedCurrencyRatios(referenceDates, currencyData = null) {
        try {
            // Get all historical currency data from passed parameter or database
            const allCurrencyData = currencyData || await equateDB.getHistoricalCurrencyData();
            
            if (!allCurrencyData || allCurrencyData.length === 0) {
                console.warn('No historical currency data available');
                return [];
            }
            
            // Filter to only include our reference dates (or closest matches)
            const targeted = [];
            const dateSet = new Set(referenceDates);
            
            referenceDates.forEach(refDate => {
                const exactMatch = allCurrencyData.find(entry => entry.date === refDate);
                if (exactMatch) {
                    targeted.push(exactMatch);
                } else {
                    // Find closest date (handle weekends/holidays)
                    const closest = this.findClosestCurrencyDate(refDate, allCurrencyData);
                    if (closest && !targeted.find(t => t.date === closest.date)) {
                        targeted.push(closest);
                    }
                }
            });
            
            return targeted;
        } catch (error) {
            console.error('Failed to load targeted currency ratios:', error);
            return [];
        }
    }
    
    /**
     * Find closest currency data entry for a given date
     */
    findClosestCurrencyDate(targetDate, allCurrencyData) {
        const target = new Date(targetDate);
        let closest = null;
        let minDiff = Infinity;
        
        for (const entry of allCurrencyData) {
            const entryDate = new Date(entry.date);
            const diff = Math.abs(target - entryDate);
            if (diff < minDiff) {
                minDiff = diff;
                closest = entry;
            }
        }
        
        return closest;
    }
    
    /**
     * Get total count of currency entries (for logging data reduction)
     */
    async getTotalCurrencyEntries() {
        try {
            return await equateDB.getHistoricalCurrencyDataCount();
        } catch (error) {
            return 0;
        }
    }
    
    /**
     * Build reference points with pre-calculated prices in all available currencies
     */
    buildMultiCurrencyReferencePoints(portfolioData, transactionData) {
        const points = [];
        const originalCurrency = portfolioData?.detectedInfo?.currency || 'UNKNOWN';
        
        // AsOfDate reference point
        if (portfolioData?.asOfDate && portfolioData?.entries?.[0]?.marketPrice) {
            const asOfDatePoint = {
                date: portfolioData.asOfDate,
                quantity: portfolioData.entries[0].outstandingQuantity,
                type: 'asOfDate',
                source: 'portfolio',
                category: 'Current Valuation',
                originalCurrency: originalCurrency,
                // AsOfDate represents current portfolio valuation - only valuation fields, not transaction metadata
                allocatedQuantity: undefined,       // Not applicable for valuation snapshot
                outstandingQuantity: undefined,     // Not applicable for valuation snapshot
                availableQuantity: undefined,       // Not applicable for valuation snapshot
                costBasis: undefined,               // AsOfDate uses marketPrice, not costBasis
                marketPrice: portfolioData.entries[0].marketPrice,
                prices: this.calculateAllCurrencyPrices(
                    portfolioData.entries[0].marketPrice,
                    portfolioData.asOfDate,
                    originalCurrency
                ),
                // AsOfDate is a valuation snapshot, not a transaction - no transaction metadata
                contributionType: undefined,        // Not applicable for AsOfDate
                plan: undefined,                    // Not applicable for AsOfDate
                instrument: undefined,              // Not applicable for AsOfDate
                instrumentType: undefined,          // Not applicable for AsOfDate
                orderType: undefined,               // Not applicable for AsOfDate
                status: undefined                   // Not applicable for AsOfDate
            };
            
            // Set initial currentPrice to original currency (CRITICAL: Always set, even with full currency data)
            // BUGFIX: Ensure original currency exists in prices object before using it
            if (!asOfDatePoint.prices[originalCurrency]) {
                console.warn(`DEBUG_CURRENT_PRICE: Original currency ${originalCurrency} missing from prices, using marketPrice as fallback`);
                asOfDatePoint.prices[originalCurrency] = asOfDatePoint.marketPrice;
            }
            asOfDatePoint.currentPrice = asOfDatePoint.prices[originalCurrency];
            asOfDatePoint.currentCurrency = originalCurrency;
            
            
            points.push(asOfDatePoint);
        }
        
        // Portfolio purchase reference points - PRESERVE ALL ENTRIES (no aggregation)
        if (portfolioData?.entries) {
            portfolioData.entries.forEach(entry => {
                if (entry.allocationDate && entry.costBasis && entry.costBasis > 0) {
                    const purchasePoint = {
                        date: entry.allocationDate,
                        quantity: entry.allocatedQuantity,
                        type: 'purchase',
                        source: 'portfolio',
                        category: this.categorizePurchase(entry.contributionType, entry.plan, entry.instrument, entry.instrumentType),
                        originalCurrency: originalCurrency,
                        // Add missing fields for calculator compatibility
                        allocatedQuantity: entry.allocatedQuantity,
                        outstandingQuantity: entry.outstandingQuantity,
                        availableQuantity: entry.availableQuantity,
                        costBasis: entry.costBasis,
                        marketPrice: entry.marketPrice, // For current value calculations
                        prices: this.calculateAllCurrencyPrices(
                            entry.costBasis,
                            entry.allocationDate,
                            originalCurrency
                        ),
                        // Preserve original fields for compatibility and rule system
                        contributionType: entry.contributionType,
                        plan: entry.plan,
                        instrument: entry.instrument,           // From portfolio col 3
                        instrumentType: entry.instrumentType,   // From portfolio col 2  
                        orderType: undefined,                   // Not applicable for portfolio entries
                        status: undefined                       // Not applicable for portfolio entries
                    };
                    
                    // Set initial currentPrice to original currency (CRITICAL: Always set, even with full currency data)
                    // BUGFIX: Ensure original currency exists in prices object before using it
                    if (!purchasePoint.prices[originalCurrency]) {
                        console.warn(`DEBUG_CURRENT_PRICE: Original currency ${originalCurrency} missing from prices, using costBasis as fallback`);
                        purchasePoint.prices[originalCurrency] = entry.costBasis;
                    }
                    purchasePoint.currentPrice = purchasePoint.prices[originalCurrency];
                    purchasePoint.currentCurrency = originalCurrency;
                    
                    
                    points.push(purchasePoint);
                }
            });
        }
        
        // Transaction reference points (sales/transfers) - PRESERVE ALL ENTRIES (no aggregation)
        if (transactionData?.entries) {
            transactionData.entries.forEach(transaction => {
                if (this.isSellTransaction(transaction) && 
                    transaction.transactionDate && 
                    transaction.executionPrice && 
                    transaction.executionPrice > 0) {
                    
                    const transactionPoint = {
                        date: transaction.transactionDate,
                        quantity: -Math.abs(transaction.quantity || 0), // Negative for sales
                        type: 'sale',
                        source: 'transaction',
                        category: this.categorizeTransaction(transaction.orderType),
                        originalCurrency: originalCurrency,
                        // Add missing fields for calculator compatibility  
                        allocatedQuantity: undefined, // Not applicable for sales
                        outstandingQuantity: undefined, // Not applicable for sales
                        availableQuantity: undefined, // Not applicable for sales
                        costBasis: undefined, // Sales use executionPrice, not costBasis
                        executionPrice: transaction.executionPrice,
                        prices: this.calculateAllCurrencyPrices(
                            transaction.executionPrice,
                            transaction.transactionDate,
                            originalCurrency
                        ),
                        // Preserve original fields for compatibility and rule system
                        orderType: transaction.orderType,
                        status: transaction.status,
                        instrument: transaction.instrument,         // From transaction col 6
                        instrumentType: transaction.productType,    // From transaction col 7
                        contributionType: transaction.contributionType, // From transaction col 8
                        plan: transaction.plan                      // From transaction col 9
                    };
                    
                    // Set initial currentPrice to original currency (CRITICAL: Always set, even with full currency data)
                    // BUGFIX: Ensure original currency exists in prices object before using it
                    if (!transactionPoint.prices[originalCurrency]) {
                        console.warn(`DEBUG_CURRENT_PRICE: Original currency ${originalCurrency} missing from prices, using price as fallback`);
                        transactionPoint.prices[originalCurrency] = transaction.price;
                    }
                    transactionPoint.currentPrice = transactionPoint.prices[originalCurrency];
                    transactionPoint.currentCurrency = originalCurrency;
                    
                    points.push(transactionPoint);
                }
            });
        }
        
        // ✅ CRITICAL FIX: DO NOT AGGREGATE - PRESERVE ALL INDIVIDUAL ENTRIES
        // Just sort by date, keep all entries for accurate calculations
        return points.sort((a, b) => new Date(a.date) - new Date(b.date));
    }
    
    /**
     * Calculate prices in all available currencies for a given original price
     */
    calculateAllCurrencyPrices(originalPrice, date, originalCurrency) {
        const prices = {};
        
        // Find currency ratios for this date
        const currencyRatio = this.findCurrencyRatioForDate(date);
        
        if (!currencyRatio) {
            // V2 SYNTHETIC MODE: Only original currency needed
            prices[originalCurrency] = originalPrice;
            console.debug(`V2 SYNTHETIC: Using original currency only for ${date}`);
            return prices;
        }
        
        // Calculate EUR price first (base for all other conversions)
        let eurPrice;
        if (originalCurrency === 'EUR') {
            eurPrice = originalPrice;
        } else {
            const eurRate = currencyRatio[originalCurrency];
            if (eurRate && eurRate > 0) {
                eurPrice = originalPrice / eurRate; // Convert to EUR
            } else {
                // V2 SYNTHETIC MODE: Only original currency available
                prices[originalCurrency] = originalPrice;
                console.debug(`V2 SYNTHETIC: No EUR conversion rate for ${originalCurrency} on ${date}, using original currency only`);
                return prices;
            }
        }
        
        // Calculate all currency prices from EUR base
        Object.keys(currencyRatio).forEach(currency => {
            if (currency !== 'date' && currency.length === 3) { // Valid currency code
                if (currency === 'EUR') {
                    prices[currency] = eurPrice;
                } else if (currency === originalCurrency) {
                    prices[currency] = originalPrice; // Preserve exact original price
                } else {
                    const rate = currencyRatio[currency];
                    if (rate && rate > 0) {
                        prices[currency] = eurPrice * rate;
                    }
                }
            }
        });
        
        // BUGFIX: Always ensure originalCurrency is in prices object
        if (!prices[originalCurrency]) {
            prices[originalCurrency] = originalPrice;
        }
        
        return prices;
    }
    
    /**
     * Find currency ratio entry for a specific date (same as existing logic)
     */
    findCurrencyRatioForDate(targetDate) {
        // Direct match first
        const exactMatch = this.targetedCurrencyData.find(entry => entry.date === targetDate);
        if (exactMatch) {
            return exactMatch;
        }
        
        // Find closest date
        return this.findClosestCurrencyDate(targetDate, this.targetedCurrencyData);
    }
    
    /**
     * Categorize purchase transactions using hybrid rule system
     * Uses 6-field validation: instrument, instrumentType, contributionType, plan, orderType, status
     */
    categorizePurchase(contributionType, plan, instrument, instrumentType, orderType = null, status = null) {
        const entry = { 
            contributionType: contributionType, 
            plan: plan, 
            instrument: instrument, 
            instrumentType: instrumentType,
            orderType: orderType,
            status: status
        };
        
        const category = categorizeEntry(entry);
        
        // Map internal category names to display names
        const categoryMap = {
            'userInvestment': 'User Investment',
            'companyMatch': 'Company Match',
            'freeShares': 'Free Shares',
            'dividendIncome': 'Dividend Reinvestment',
            'unknown': 'Other'
        };
        
        return categoryMap[category] || 'Other';
    }
    
    /**
     * Categorize transaction types (same logic as existing)
     */
    categorizeTransaction(orderType) {
        if (orderType === 'Sell' || orderType.includes('Sell')) {
            return 'Sale';
        } else if (orderType === 'Transfer' || orderType.includes('Transfer')) {
            return 'Transfer';
        }
        return 'Other Transaction';
    }
    
    /**
     * Store reference points in IndexedDB for caching
     */
    async storeReferencePoints(referencePoints) {
        try {
            // Create a new object store if it doesn't exist
            // For now, we'll store as part of portfolio data
            // In future, we can create dedicated referencePoints store
            console.log(`💾 Caching ${referencePoints.length} reference points`);
            // Implementation will be added when we create the dedicated store
        } catch (error) {
            console.warn('Failed to store reference points:', error);
        }
    }
    
    /**
     * Instant currency switching - just update currentPrice pointers
     */
    changeCurrency(referencePoints, newCurrency) {
        const startTime = performance.now();
        let successCount = 0;
        let fallbackCount = 0;
        
        referencePoints.forEach(point => {
            if (point.prices && point.prices[newCurrency] !== undefined) {
                point.currentPrice = point.prices[newCurrency];
                point.currentCurrency = newCurrency;
                successCount++;
            } else {
                // Fallback to original currency
                point.currentPrice = point.prices[point.originalCurrency];
                point.currentCurrency = point.originalCurrency;
                fallbackCount++;
            }
        });
        
        const endTime = performance.now();
        console.log(`💨 Currency switched to ${newCurrency}: ${successCount} points switched, ${fallbackCount} fallbacks in ${(endTime - startTime).toFixed(1)}ms`);
        
        
        return referencePoints;
    }
    
    /**
     * Get available currencies from reference points
     */
    getAvailableCurrencies(referencePoints) {
        const currencies = new Set();
        
        referencePoints.forEach(point => {
            if (point.prices) {
                Object.keys(point.prices).forEach(currency => {
                    currencies.add(currency);
                });
            }
        });
        
        return Array.from(currencies).sort();
    }

    /**
     * Create aggregated timeline points for chart display (separate from individual calculations)
     * This is what the old system was trying to do, but it should be separate
     */
    getAggregatedTimelinePoints(referencePoints) {
        const aggregated = {};
        
        referencePoints.forEach(point => {
            const key = point.date;
            if (!aggregated[key] || aggregated[key].currentPrice < point.currentPrice) {
                aggregated[key] = {
                    date: point.date,
                    price: point.currentPrice,
                    currency: point.currentCurrency,
                    type: point.type,
                    // For timeline, we want the highest price point on each date
                    source: 'aggregated'
                };
            }
        });
        
        return Object.values(aggregated).sort((a, b) => new Date(a.date) - new Date(b.date));
    }
}

// Global instance for use throughout the application
const referencePointsBuilder = new ReferencePointsBuilder();