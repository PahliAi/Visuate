/**
 * Portfolio Calculation Engine for Equate
 * Implements all portfolio metrics calculations with manual price override capability
 */

class PortfolioCalculator {
    /**
     * Create a new PortfolioCalculator instance
     * @constructor
     */
    constructor() {
        this.portfolioData = null;
        this.transactionData = null;
        this.historicalPrices = null;
        this.currentPrice = null;
        this.priceSource = 'historical'; // 'historical' or 'manual'
        this.calculations = null;
        this.currency = null; // Set externally by FileAnalyzer
    }

    /**
     * Parse date string as UTC to avoid timezone issues
     * This ensures consistent behavior across all timezones globally
     * @param {string|Date|null} dateString - Date string or Date object to parse
     * @returns {Date|null} UTC Date object or null if invalid
     */
    parseUTCDate(dateString) {
        if (!dateString) return null;
        
        // If it's already a Date object, create UTC version
        if (dateString instanceof Date) {
            const isoString = dateString.toISOString().split('T')[0];
            return new Date(isoString + 'T00:00:00.000Z');
        }
        
        // If it's a string in YYYY-MM-DD format, parse as UTC
        if (typeof dateString === 'string') {
            // Extract just the date part if it contains time
            const datePart = dateString.split('T')[0];
            return new Date(datePart + 'T00:00:00.000Z');
        }
        
        return null;
    }

    /**
     * Set portfolio and transaction data for calculations
     * @param {Object} portfolioData - Portfolio data object containing entries array
     * @param {Object|null} [transactionData=null] - Optional transaction data for realized gains
     * @returns {void}
     */
    setPortfolioData(portfolioData, transactionData = null, currency = null) {
        this.portfolioData = portfolioData;
        this.transactionData = transactionData;
        this.currency = currency;
        
        config.debug('🔧 Calculator received data:', {
            portfolioEntries: portfolioData?.entries?.length || 0,
            transactionEntries: transactionData?.entries?.length || 0,
            hasTransactionData: !!transactionData
        });
        
        if (transactionData && transactionData.entries) {
            config.debug('📋 Sample transaction data:', transactionData.entries.slice(0, 2));
        }
    }

    /**
     * Set historical prices and enhance with current price data
     */
    async setHistoricalPrices(historicalPrices) {
        this.historicalPrices = [...historicalPrices]; // Copy array
        await this.enhanceHistoricalPrices();
        this.updateCurrentPrice();
    }

    /**
     * Enhance historical prices with current portfolio data and manual prices
     * Now async to save AsOfDate to IndexedDB
     */
    async enhanceHistoricalPrices() {
        
        if (!this.portfolioData) {
            return;
        }

        // Add portfolio "As of date" entry if it's more recent than latest historical price
        const asOfDate = this.extractAsOfDate();
        const marketPrice = this.extractMarketPrice();
        
        if (asOfDate && marketPrice) {
            // AsOfDate should only be ADDED to historical data, never used to LIMIT it
            // Always use safe add/update logic that preserves existing historical data
            const existingEntry = this.historicalPrices.find(p => p.date === asOfDate);
            
            if (!existingEntry) {
                config.debug(`📈 Portfolio AsOfDate ${asOfDate} (${marketPrice}) not found in historical prices, adding...`);
                
                // Add to in-memory array
                this.historicalPrices.push({
                    date: asOfDate,
                    price: marketPrice
                });
                
                // Save to IndexedDB for persistence across sessions
                try {
                    await equateDB.appendHistoricalPrice(asOfDate, marketPrice);
                    config.debug(`✅ AsOfDate price saved to IndexedDB permanently`);
                } catch (error) {
                    config.error('❌ Failed to save AsOfDate to IndexedDB:', error);
                }
            } else {
                // Historical price already exists for AsOfDate - keep the historical price
                // Historical prices (from hist.xlsx) are authoritative closing prices and should take precedence
                // over intraday portfolio market prices
                config.debug(`📈 Portfolio AsOfDate ${asOfDate} already has historical price €${existingEntry.price}, keeping historical price (portfolio has €${marketPrice})`);
                config.debug(`🔍 DEBUG: Preserving historical closing price over portfolio intraday price`);
            }
        }

        // Add manual price for scenario testing (today or tomorrow)
        if (this.priceSource === 'manual' && this.currentPrice) {
            const todayStr = new Date().toISOString().split('T')[0];
            const tomorrowStr = new Date(Date.now() + 24*60*60*1000).toISOString().split('T')[0];
            
            // Check if today already has a price point
            const todayPrice = this.historicalPrices.find(p => p.date === todayStr);
            
            if (!todayPrice) {
                // No price for today, add manual price for today
                this.historicalPrices.push({
                    date: todayStr,
                    price: this.currentPrice
                });
                config.debug(`📊 Manual price point added for TODAY ${todayStr} with price €${this.currentPrice}`);
            } else {
                // Today has a price, check if it's the same
                if (Math.abs(todayPrice.price - this.currentPrice) < 0.01) {
                    // Same price as today - no point in scenario testing the same value
                    config.debug(`📊 Manual price €${this.currentPrice} matches today's price - no scenario needed`);
                    // Don't add any manual price point since it's identical to reality
                } else {
                    // Different price - add scenario for tomorrow
                    this.historicalPrices.push({
                        date: tomorrowStr,
                        price: this.currentPrice
                    });
                    config.debug(`📊 Manual price point added for TOMORROW ${tomorrowStr} with price €${this.currentPrice} (today has €${todayPrice.price})`);
                }
            }
        }

        // Sort by date to maintain chronological order
        this.historicalPrices.sort((a, b) => new Date(a.date) - new Date(b.date));
        
        // DEBUG: Final state after enhanceHistoricalPrices
        config.debug(`🔍 DEBUG enhanceHistoricalPrices FINAL state:`, {
            count: this.historicalPrices.length,
            lastFew: this.historicalPrices.slice(-3),
            hasAsOfDate: this.historicalPrices.some(p => p.date === '2025-08-01'),
            priceSource: this.priceSource
        });
        
    }

    /**
     * Set manual price override for current valuation
     * @param {number|string} price - Manual share price to use for calculations
     * @returns {Promise<void>} Promise that resolves when price is set and historical data enhanced
     */
    async setManualPrice(price) {
        config.debug('🎯 Calculator.setManualPrice called with:', price);
        this.currentPrice = parseFloat(price);
        this.priceSource = 'manual';
        
        // Initialize empty historical prices array if null (for non-Allianz companies)
        if (!this.historicalPrices) {
            this.historicalPrices = [];
            config.debug('🎯 Initialized empty historical prices array for manual price mode');
        }
        
        config.debug('🎯 Price source set to manual, current price:', this.currentPrice);
        await this.enhanceHistoricalPrices(); // Re-enhance with new manual price
        config.debug('🎯 Historical prices enhanced, returning to caller for recalculation...');
        // Don't call recalculate here - let the caller handle it to avoid double calculation
    }

    /**
     * Extract "As of date" from portfolio file (cell B3)
     */
    extractAsOfDate() {
        if (!this.portfolioData || !this.portfolioData.asOfDate) {
            return null;
        }
        return this.portfolioData.asOfDate;
    }

    /**
     * Extract market price from portfolio data
     */
    extractMarketPrice() {
        if (!this.portfolioData || !this.portfolioData.entries || this.portfolioData.entries.length === 0) {
            return null;
        }
        return this.portfolioData.entries[0].marketPrice;
    }

    /**
     * Get latest historical price date
     */
    getLatestHistoricalDate() {
        if (!this.historicalPrices || this.historicalPrices.length === 0) {
            return null;
        }
        const sorted = [...this.historicalPrices].sort((a, b) => new Date(b.date) - new Date(a.date));
        return sorted[0].date;
    }

    /**
     * Clear manual price and use historical
     */
    clearManualPrice() {
        this.priceSource = 'historical';
        this.updateCurrentPrice();
        this.recalculate();
    }

    /**
     * Update current price from historical data
     */
    updateCurrentPrice() {
        if (this.priceSource === 'manual') return;

        if (this.historicalPrices && this.historicalPrices.length > 0) {
            // Get the latest price from historical data
            const sortedPrices = [...this.historicalPrices].sort((a, b) => new Date(b.date) - new Date(a.date));
            this.currentPrice = sortedPrices[0].price;
        } else {
            // Fallback to market price from portfolio data
            if (this.portfolioData && this.portfolioData.entries.length > 0) {
                this.currentPrice = this.portfolioData.entries[0].marketPrice;
            }
        }
    }

    /**
     * Calculate all portfolio metrics including returns, CAGR, and current values
     * @returns {Object} Complete portfolio calculations object with all metrics
     * @throws {Error} If portfolio data or current price is missing
     */
    async calculate() {
        if (!this.portfolioData || !this.currentPrice) {
            throw new Error('Missing portfolio data or current price');
        }

        const entries = this.portfolioData.entries;
        
        // Use externally provided currency (from FileAnalyzer)
        const portfolioCurrency = this.currency || 'UNKNOWN';
        
        // Currency validation is now handled by FileAnalyzer during pre-parsing
        let currencyWarning = null;
        
        // Calculate core metrics
        const userInvestment = this.calculateUserInvestment(entries);
        const companyMatch = this.calculateCompanyMatch(entries);
        const freeShares = this.calculateFreeShares(entries);
        const companyInvestment = companyMatch + freeShares; // Legacy total
        const dividendIncome = this.calculateDividendIncome(entries);
        const totalSold = this.calculateTotalSold();
        const currentValue = this.calculateCurrentValue(entries);
        const totalCurrentPortfolio = currentValue; // Current value of remaining shares
        const totalInvestment = userInvestment + companyInvestment + dividendIncome; // Sum of all investments
        const totalValue = currentValue + totalSold; // Current + Sold
        const totalReturn = totalValue - userInvestment; // Total Value - Your Investment (RETURN ON YOUR INVESTMENT)
        const returnPercentage = userInvestment > 0 ? ((totalReturn / userInvestment) * 100) : 0; // RETURN % ON YOUR INVESTMENT
        const returnOnTotalInvestment = this.calculateReturnOnTotalInvestment(totalValue, totalInvestment);
        const returnPercentageOnTotalInvestment = this.calculateReturnPercentageOnTotalInvestment(totalValue, totalInvestment);
        const availableShares = this.calculateAvailableShares(entries);
        const blockedShares = this.calculateBlockedShares(entries);
        const xirrUserInvestment = await this.calculateXIRRUserInvestment(entries, totalValue);
        const xirrTotalInvestment = await this.calculateXIRRTotalInvestment(entries, totalValue);

        this.calculations = {
            userInvestment,
            companyMatch,
            freeShares,
            companyInvestment, // Legacy total (companyMatch + freeShares)
            dividendIncome,
            totalInvestment,
            totalSold,
            currentValue,
            totalCurrentPortfolio,
            totalValue,
            totalReturn, // RETURN ON YOUR INVESTMENT
            returnPercentage, // RETURN % ON YOUR INVESTMENT
            returnOnTotalInvestment, // RETURN ON TOTAL INVESTMENT
            returnPercentageOnTotalInvestment, // RETURN % ON TOTAL INVESTMENT
            xirrUserInvestment, // XIRR ON YOUR INVESTMENT
            xirrTotalInvestment, // XIRR ON TOTAL INVESTMENT
            annualGrowth: xirrUserInvestment, // Legacy compatibility
            monthlyReturn: xirrUserInvestment / 12,
            availableShares,
            blockedShares,
            totalShares: availableShares + blockedShares.total,
            currentPrice: this.currentPrice,
            priceSource: this.priceSource,
            priceDate: this.getPriceDate(),
            currency: portfolioCurrency,
            currencyWarning: currencyWarning,
            lastCalculated: new Date().toISOString()
        };

        return this.calculations;
    }

    /**
     * Recalculate with current settings
     */
    async recalculate() {
        config.debug('🔄 Recalculate called, has portfolio data:', !!this.portfolioData);
        if (this.portfolioData) {
            config.debug('🔄 Calling calculate() with current price:', this.currentPrice, 'source:', this.priceSource);
            return await this.calculate();
        }
        return null;
    }

    /**
     * Calculate total user investment
     * SUMIFS(Purchase amounts, "Employee Share Purchase Plan", Outstanding > 0)
     */
    calculateUserInvestment(entries) {
        // All user investments ever made (including those already sold)
        return entries
            .filter(entry => 
                entry.contributionType === 'Purchase' && 
                entry.plan === 'Employee Share Purchase Plan'
            )
            .reduce((sum, entry) => sum + (entry.costBasis * entry.allocatedQuantity), 0);
    }

    /**
     * Calculate real company match investment
     * SUMIFS(Company match, Employee Share Purchase Plan, Outstanding > 0)
     */
    calculateCompanyMatch(entries) {
        // Company match contributions for Employee Share Purchase Plan
        return entries
            .filter(entry => 
                entry.contributionType === 'Company match' && 
                entry.plan === 'Employee Share Purchase Plan'
            )
            .reduce((sum, entry) => sum + (entry.costBasis * entry.allocatedQuantity), 0);
    }

    /**
     * Calculate free shares investment
     * SUMIFS(Award, Free Share, Outstanding > 0)
     */
    calculateFreeShares(entries) {
        // Award contributions for Free Share plan
        return entries
            .filter(entry => 
                entry.contributionType === 'Award' && 
                entry.plan === 'Free Share'
            )
            .reduce((sum, entry) => sum + (entry.costBasis * entry.allocatedQuantity), 0);
    }

    /**
     * Calculate total company investment (Company match + Free shares)
     * Legacy method for backward compatibility
     */
    calculateCompanyInvestment(entries) {
        return this.calculateCompanyMatch(entries) + this.calculateFreeShares(entries);
    }

    /**
     * Calculate total dividend income
     * SUMIFS(Dividend reinvestments, Outstanding > 0)
     */
    calculateDividendIncome(entries) {
        // All dividend reinvestments ever made (including those already sold)
        return entries
            .filter(entry => 
                entry.plan === 'Allianz Dividend Reinvestment'
            )
            .reduce((sum, entry) => sum + (entry.costBasis * entry.allocatedQuantity), 0);
    }

    /**
     * Check if an order type represents a sale or transfer (shares leaving the account)
     */
    isSellOrTransferOrderType(orderType) {
        return orderType === 'Sell' || 
               orderType === 'Sell with price limit' ||
               orderType === 'Sell at market price' ||
               orderType === 'Transfer';
    }

    /**
     * Calculate total sold amount from completed transactions
     */
    calculateTotalSold() {
        if (!this.transactionData || !this.transactionData.entries) {
            return 0;
        }

        return this.transactionData.entries
            .filter(entry => this.isSellOrTransferOrderType(entry.orderType))
            .filter(entry => entry.status === 'Executed')
            .reduce((sum, entry) => sum + (Math.abs(entry.quantity) * (entry.executionPrice || 0)), 0);
    }

    /**
     * Calculate current portfolio value
     * Use manual price when available, otherwise use Excel values
     */
    calculateCurrentValue(entries) {
        if (this.priceSource === 'manual' && this.currentPrice) {
            // Calculate using manual price: outstanding shares * manual price
            const totalValue = entries
                .reduce((sum, entry) => sum + (entry.outstandingQuantity * this.currentPrice), 0);
            config.debug('💰 Manual price calculation result:', totalValue);
            return totalValue;
        } else if (this.currentPrice) {
            // Calculate using most recent price (from historical data or portfolio): outstanding shares * current price
            // This ensures cards and timeline chart use the same current price
            const totalValue = entries
                .reduce((sum, entry) => sum + (entry.outstandingQuantity * this.currentPrice), 0);
            config.debug(`💰 Current price calculation result: ${totalValue} (using price ${this.currentPrice})`);
            return totalValue;
        } else {
            // Fallback to market price from Excel: outstanding shares * market price
            // This ensures we always use transaction currency, not display currency
            const totalValue = entries
                .reduce((sum, entry) => sum + (entry.outstandingQuantity * entry.marketPrice), 0);
            config.debug('💰 Fallback to market price from portfolio file');
            return totalValue;
        }
    }

    /**
     * Calculate Compound Annual Growth Rate (CAGR)
     */
    calculateCAGR(entries, currentValue) {
        // Find the earliest investment date (including sold shares)
        const dates = entries
            .filter(entry => entry.allocationDate)
            .map(entry => this.parseUTCDate(entry.allocationDate))
            .filter(date => date !== null)
            .sort((a, b) => a - b);

        if (dates.length === 0) return 0;

        const startDate = dates[0];
        const endDate = new Date();
        const years = (endDate - startDate) / (365.25 * 24 * 60 * 60 * 1000);

        if (years <= 0) return 0;

        // Calculate user investment only (company match and dividends are returns, not investments)
        const userInvestment = this.calculateUserInvestment(entries);

        if (userInvestment <= 0) return 0;

        // CAGR = (Ending Value / Beginning Value)^(1/years) - 1
        const cagr = (Math.pow(currentValue / userInvestment, 1 / years) - 1) * 100;
        
        return isNaN(cagr) ? 0 : cagr;
    }

    /**
     * Calculate return on total investment
     * Total Value - Total Investment
     */
    calculateReturnOnTotalInvestment(totalValue, totalInvestment) {
        return totalValue - totalInvestment;
    }

    /**
     * Calculate return percentage on total investment  
     * (Total Value - Total Investment) / Total Investment * 100
     */
    calculateReturnPercentageOnTotalInvestment(totalValue, totalInvestment) {
        return totalInvestment > 0 ? ((totalValue - totalInvestment) / totalInvestment * 100) : 0;
    }

    /**
     * Calculate available shares (can be sold now)
     */
    calculateAvailableShares(entries) {
        // Total shares owned = sum of Outstanding Quantity column
        const total = entries
            .reduce((sum, entry) => sum + (entry.outstandingQuantity || 0), 0);
        return Math.round(total * 100) / 100; // Round to 2 decimals
    }

    /**
     * Calculate blocked shares by unlock date
     */
    calculateBlockedShares(entries) {
        const blocked = {
            total: 0,
            byYear: {}
        };

        entries
            .filter(entry => entry.outstandingQuantity > entry.availableQuantity)
            .forEach(entry => {
                const blockedQty = entry.outstandingQuantity - entry.availableQuantity;
                
                if (entry.availableFrom) {
                    const availableFromDate = this.parseUTCDate(entry.availableFrom);
                    if (availableFromDate) {
                        const unlockYear = availableFromDate.getUTCFullYear();
                        blocked.byYear[unlockYear] = (blocked.byYear[unlockYear] || 0) + blockedQty;
                    }
                }
                
                blocked.total += blockedQty;
            });

        return blocked;
    }

    /**
     * Get price date based on source
     */
    getPriceDate() {
        if (this.priceSource === 'manual') {
            return new Date().toISOString().split('T')[0];
        }

        if (this.historicalPrices && this.historicalPrices.length > 0) {
            const sortedPrices = [...this.historicalPrices].sort((a, b) => new Date(b.date) - new Date(a.date));
            return sortedPrices[0].date;
        }

        return new Date().toISOString().split('T')[0];
    }

    /**
     * Create complete daily timeline from first purchase to today for charting
     * Each day gets: shares_prev, shares_current, portfolio_value, reason, profit_loss
     * @returns {Promise<Array>} Array of timeline objects with daily portfolio progression
     */
    async getPortfolioTimeline() {
        

        if (!this.portfolioData || !this.historicalPrices || this.historicalPrices.length === 0) {
            return [];
        }

        // DEBUG: Check what historical prices we actually have when timeline is generated
        config.debug('🔍 DEBUG getPortfolioTimeline - historicalPrices:', {
            count: this.historicalPrices.length,
            lastFew: this.historicalPrices.slice(-3),
            hasAsOfDate: this.historicalPrices.some(p => p.date === '2025-08-01'),
            asOfDateEntry: this.historicalPrices.find(p => p.date === '2025-08-01')
        });

        // Step 1: Find the earliest purchase date (timezone-safe)
        const entries = this.portfolioData.entries;
        const earliestPurchaseDate = entries
            .map(entry => this.parseUTCDate(entry.allocationDate))
            .filter(date => date !== null)
            .reduce((earliest, current) => current < earliest ? current : earliest);


        // Step 2: Create transaction lookup by date
        const transactionsByDate = this.buildTransactionLookup();
        const purchasesByDate = this.buildPurchaseLookup();

        // Step 3: Create price lookup by date
        const pricesByDate = {};
        this.historicalPrices.forEach(price => {
            pricesByDate[price.date] = price.price;
        });


        // Step 4: Generate daily timeline from first purchase to latest enhanced price date
        const timeline = [];
        
        // Find the latest price date (could be today or tomorrow if manual price scenario)
        let latestPriceDate = new Date();
        
        if (this.historicalPrices.length > 0) {
            // Sort prices and get the latest date
            const sortedPrices = [...this.historicalPrices].sort((a, b) => new Date(b.date) - new Date(a.date));
            // Parse as UTC date to avoid timezone issues with AsOfDate
            latestPriceDate = new Date(sortedPrices[0].date + 'T00:00:00.000Z');
            
            // DEBUG: Timeline boundary calculation
            config.debug('🔍 DEBUG timeline boundaries:', {
                latestHistoricalDate: sortedPrices[0].date,
                latestPriceDate: latestPriceDate.toISOString().split('T')[0],
                isAsOfDate: sortedPrices[0].date === '2025-08-01'
            });
        }
        
        
        latestPriceDate.setUTCHours(23, 59, 59, 999); // Ensure we include all of the latest date in UTC
        let currentDate = new Date(earliestPurchaseDate);
        currentDate.setUTCHours(0, 0, 0, 0); // Start at beginning of day in UTC
        
        let outstandingShares = 0; // Track actual shares currently owned (not cumulative)
        let cumulativeUserInvestment = 0; // Track total user money invested (never decreases)
        let cumulativeSaleProceeds = 0; // Track total money received from sales
        
        while (currentDate <= latestPriceDate) {
            const dateStr = currentDate.toISOString().split('T')[0];
            
            // DEBUG: Loop progress tracking
            // if (dateStr >= '2025-07-29') {
            //     config.debug('🔍 DEBUG loop near end:', {
            //         currentDateStr: dateStr,
            //         currentDateTime: currentDate.toISOString(),
            //         latestPriceDateTime: latestPriceDate.toISOString(),
            //         comparison: currentDate <= latestPriceDate
            //     });
            // }
            
            const prevOutstandingShares = outstandingShares;
            
            // Check for transactions on this date
            let reason = '';
            let shareChange = 0;
            let userInvestmentChange = 0; // User money invested this day
            let saleProceedsChange = 0; // Money received from sales this day

            // Process purchases
            if (purchasesByDate[dateStr]) {
                purchasesByDate[dateStr].forEach(purchase => {
                    // Add shares purchased on this date (use historical allocated quantity)
                    shareChange += purchase.allocatedQuantity;
                    
                    // Only count user purchases as investment (company match and dividends are profits, not investments)
                    if (purchase.contributionType === 'Purchase' && purchase.plan === 'Employee Share Purchase Plan') {
                        userInvestmentChange += purchase.costBasis * purchase.allocatedQuantity;
                    }
                    
                    reason += `Bought ${purchase.allocatedQuantity} shares (${purchase.contributionType}); `;
                });
            }

            // Process sales only (dividends are already in portfolio data as reinvestments)
            if (transactionsByDate[dateStr]) {
                transactionsByDate[dateStr].forEach(transaction => {
                    if (this.isSellOrTransferOrderType(transaction.orderType)) {
                        // Sale removes shares from outstanding count
                        const quantitySold = Math.abs(transaction.quantity);
                        shareChange -= quantitySold;
                        
                        // Add sale proceeds to cumulative total
                        const saleAmount = quantitySold * (transaction.executionPrice || 0);
                        saleProceedsChange += saleAmount;
                        
                        const currencySymbol = this.getCurrencySymbol(this.currency);
                        reason += `Sold ${quantitySold} shares for ${currencySymbol}${saleAmount.toFixed(2)}; `;
                    }
                    // Note: Dividends are ignored here as they're already included in portfolio data as reinvestments
                });
            }

            // Update running totals
            outstandingShares += shareChange;
            cumulativeUserInvestment += userInvestmentChange;
            cumulativeSaleProceeds += saleProceedsChange;

            // Get current price (use closest available price)
            const currentPrice = this.findClosestPrice(dateStr, pricesByDate);
            
            // Calculate portfolio value using outstanding shares
            const portfolioValue = outstandingShares * currentPrice;
            
            // Calculate profit/loss: Portfolio Value + Sale Proceeds - User Investment
            const profitLoss = portfolioValue + cumulativeSaleProceeds - cumulativeUserInvestment;

            // Create daily record
            timeline.push({
                date: dateStr,
                sharesPrevDay: prevOutstandingShares,
                sharesCurrentDay: outstandingShares,
                portfolioValue: portfolioValue,
                reason: reason.trim() || null,
                profitLoss: profitLoss,
                cumulativeUserInvestment: cumulativeUserInvestment,
                cumulativeSaleProceeds: cumulativeSaleProceeds,
                currentPrice: currentPrice,
                hasTransaction: !!(shareChange !== 0)
            });

            // DEBUG: Check if we're processing the asOfDate
            if (dateStr === '2025-08-01') {
                config.debug('🎯 DEBUG: Processing asOfDate 2025-08-01 in timeline loop:', {
                    portfolioValue: portfolioValue,
                    currentPrice: currentPrice,
                    outstandingShares: outstandingShares,
                    timelineLength: timeline.length
                });
            }

            // Move to next day (using UTC to avoid timezone issues)
            currentDate.setUTCDate(currentDate.getUTCDate() + 1);
            
        }

        // DEBUG: Final timeline state
        config.debug('🔍 DEBUG final timeline:', {
            totalDays: timeline.length,
            lastFewDays: timeline.slice(-3).map(d => ({date: d.date, portfolioValue: d.portfolioValue})),
            hasAsOfDate: timeline.some(d => d.date === '2025-08-01'),
            asOfDateEntry: timeline.find(d => d.date === '2025-08-01')
        });
        
        return timeline;
    }

    /**
     * Build transaction lookup by date
     */
    buildTransactionLookup() {
        const transactionsByDate = {};
        if (this.transactionData && this.transactionData.entries) {
            this.transactionData.entries.forEach(transaction => {
                const date = transaction.transactionDate;
                if (!transactionsByDate[date]) {
                    transactionsByDate[date] = [];
                }
                transactionsByDate[date].push(transaction);
            });
        }
        return transactionsByDate;
    }

    /**
     * Build purchase lookup by date
     */
    buildPurchaseLookup() {
        const purchasesByDate = {};
        this.portfolioData.entries.forEach(entry => {
            const date = entry.allocationDate;
            if (!purchasesByDate[date]) {
                purchasesByDate[date] = [];
            }
            purchasesByDate[date].push(entry);
        });
        return purchasesByDate;
    }

    /**
     * Find closest available price for a given date
     */
    findClosestPrice(targetDate, pricesByDate) {
        // Try exact match first
        if (pricesByDate[targetDate]) {
            return pricesByDate[targetDate];
        }

        // Find closest prior date with price data
        const targetTime = new Date(targetDate).getTime();
        let closestPrice = null;
        let closestTimeDiff = Infinity;

        Object.keys(pricesByDate).forEach(dateStr => {
            const priceTime = new Date(dateStr).getTime();
            const timeDiff = Math.abs(targetTime - priceTime);
            
            if (timeDiff < closestTimeDiff) {
                closestTimeDiff = timeDiff;
                closestPrice = pricesByDate[dateStr];
            }
        });

        return closestPrice || 0;
    }

    /**
     * Get current calculations
     */
    getCalculations() {
        return this.calculations;
    }

    /**
     * Get currency symbol for a given currency code
     */
    getCurrencySymbol(currency) {
        const symbols = {
            'EUR': '€',
            'USD': '$',
            'GBP': '£',
            'CNY': '¥',
            'JPY': '¥',
            'CAD': 'C$',
            'AUD': 'A$',
            'CHF': 'CHF',
            'SEK': 'kr',
            'NOK': 'kr',
            'DKK': 'kr',
            'PLN': 'zł',
            'CZK': 'Kč',
            'HUF': 'Ft',
            'RON': 'lei',
            'BGN': 'лв',
            'HRK': 'kn',
            'RUB': '₽',
            'TRY': '₺',
            'INR': '₹',
            'KRW': '₩',
            'SGD': 'S$',
            'HKD': 'HK$',
            'NZD': 'NZ$',
            'ZAR': 'R',
            'BRL': 'R$',
            'MXN': '$'
        };
        return symbols[currency] || currency || '';
    }

    /**
     * Format currency value
     */
    formatCurrency(value, currency = null, decimals = 2) {
        const symbol = this.getCurrencySymbol(currency);
        return `${symbol} ${value.toLocaleString('en-US', { 
            minimumFractionDigits: decimals, 
            maximumFractionDigits: decimals 
        })}`;
    }

    /**
     * Format percentage value
     */
    formatPercentage(value, decimals = 2) {
        return `${value.toFixed(decimals)}%`;
    }

    /**
     * Clear all data
     */
    clearData() {
        this.portfolioData = null;
        this.transactionData = null;
        this.historicalPrices = null;
        this.currentPrice = null;
        this.priceSource = 'historical';
        this.calculations = null;
    }


    /**
     * Calculate XIRR based only on user investment (excluding company contributions)
     * Uses direct Excel data for consistency
     */
    async calculateXIRRUserInvestment(entries, totalValue) {
        try {
            
            const cashFlows = [];
            
            // Add all user purchases as negative cash flows (from portfolio Excel)
            entries.forEach(entry => {
                if (entry.allocationDate && 
                    entry.contributionType === 'Purchase' && 
                    entry.plan === 'Employee Share Purchase Plan') {
                    
                    const date = this.parseUTCDate(entry.allocationDate);
                    if (date) {
                        const amount = -(entry.costBasis * entry.allocatedQuantity);
                        cashFlows.push({
                            date: date,
                            amount: amount,
                            type: 'user_investment',
                            description: `User purchase: ${entry.allocatedQuantity} shares`
                        });
                    }
                }
            });
            
            // Add user sales from transaction file (if available)
            if (this.transactionData && this.transactionData.entries) {
                this.transactionData.entries.forEach(transaction => {
                    if (transaction.transactionDate && 
                        this.isSellOrTransferOrderType(transaction.orderType) &&
                        transaction.status === 'Executed') {
                        
                        const date = this.parseUTCDate(transaction.transactionDate);
                        if (date) {
                            const amount = Math.abs(transaction.quantity) * (transaction.executionPrice || 0); // Positive for sales
                            cashFlows.push({
                                date: date,
                                amount: amount,
                                type: 'user_sale',
                                description: `User sale: ${Math.abs(transaction.quantity)} shares`
                            });
                            config.debug(`💹 User sale: ${transaction.transactionDate} = +${amount.toFixed(2)}`);
                        }
                    }
                });
            }
            
            // Add current portfolio value as final cash flow
            const currentPortfolioValue = this.calculateCurrentValue(entries);
            if (currentPortfolioValue > 0) {
                cashFlows.push({
                    date: new Date(),
                    amount: currentPortfolioValue,
                    type: 'current_value',
                    description: 'Current portfolio value'
                });
            }
            
            // Sort by date
            cashFlows.sort((a, b) => a.date - b.date);
            
            if (cashFlows.length < 2) {
                return 0;
            }
            
            
            return this.calculateXIRRFromCashFlows(cashFlows);
        } catch (error) {
            config.warn('XIRR User Investment calculation error:', error);
            return 0;
        }
    }

    /**
     * Calculate XIRR based on total investment (including all contributions)
     * Uses direct Excel data for all contribution types
     */
    async calculateXIRRTotalInvestment(entries, totalValue) {
        try {
            
            const cashFlows = [];
            
            // Add ALL contributions as negative cash flows (from portfolio Excel)
            entries.forEach(entry => {
                if (entry.allocationDate && entry.costBasis && entry.allocatedQuantity) {
                    const date = this.parseUTCDate(entry.allocationDate);
                    if (date) {
                        const amount = -(entry.costBasis * entry.allocatedQuantity);
                        let contributionDescription = '';
                        
                        // Categorize contribution type
                        if (entry.contributionType === 'Purchase' && entry.plan === 'Employee Share Purchase Plan') {
                            contributionDescription = `User purchase: ${entry.allocatedQuantity} shares`;
                        } else if (entry.contributionType === 'Company match') {
                            contributionDescription = `Company match: ${entry.allocatedQuantity} shares`;
                        } else if (entry.contributionType === 'Award') {
                            contributionDescription = `Free shares award: ${entry.allocatedQuantity} shares`;
                        } else if (entry.plan === 'Allianz Dividend Reinvestment') {
                            contributionDescription = `Dividend reinvestment: ${entry.allocatedQuantity} shares`;
                        } else {
                            contributionDescription = `${entry.contributionType} (${entry.plan}): ${entry.allocatedQuantity} shares`;
                        }
                        
                        cashFlows.push({
                            date: date,
                            amount: amount,
                            type: 'total_investment',
                            contributionType: entry.contributionType,
                            plan: entry.plan,
                            description: contributionDescription
                        });
                        
                    }
                }
            });
            
            // Add ALL sales from transaction file (if available)
            if (this.transactionData && this.transactionData.entries) {
                this.transactionData.entries.forEach(transaction => {
                    if (transaction.transactionDate && 
                        this.isSellOrTransferOrderType(transaction.orderType) &&
                        transaction.status === 'Executed') {
                        
                        const date = this.parseUTCDate(transaction.transactionDate);
                        if (date) {
                            const amount = Math.abs(transaction.quantity) * (transaction.executionPrice || 0); // Positive for sales
                            cashFlows.push({
                                date: date,
                                amount: amount,
                                type: 'sale',
                                description: `Sale: ${Math.abs(transaction.quantity)} shares from ${transaction.plan}`
                            });
                        }
                    }
                });
            }
            
            // Add current portfolio value as final cash flow
            const currentPortfolioValue = this.calculateCurrentValue(entries);
            if (currentPortfolioValue > 0) {
                cashFlows.push({
                    date: new Date(),
                    amount: currentPortfolioValue,
                    type: 'current_value',
                    description: 'Current portfolio value'
                });
            }
            
            // Sort by date
            cashFlows.sort((a, b) => a.date - b.date);
            
            if (cashFlows.length < 2) {
                return 0;
            }
            
            
            return this.calculateXIRRFromCashFlows(cashFlows);
        } catch (error) {
            config.warn('XIRR Total Investment calculation error:', error);
            return 0;
        }
    }


    /**
     * Calculate XIRR using Newton-Raphson method
     */
    calculateXIRRFromCashFlows(cashFlows) {
        let rate = 0.1; // Initial guess: 10%
        const tolerance = 1e-6;
        const maxIterations = 100;
        
        
        for (let i = 0; i < maxIterations; i++) {
            const { npv, derivative } = this.calculateNPVAndDerivative(cashFlows, rate);
            
            if (Math.abs(npv) < tolerance) {
                return rate * 100; // Convert to percentage
            }
            
            if (Math.abs(derivative) < tolerance) {
                break; // Avoid division by zero
            }
            
            const newRate = rate - (npv / derivative);
            
            // Prevent extreme values
            if (newRate < -0.99) rate = -0.99;
            else if (newRate > 10) rate = 10;
            else rate = newRate;
        }
        
        // If Newton-Raphson didn't converge, try a fallback calculation
        const fallbackRate = this.calculateFallbackXIRR(cashFlows);
        
        return isNaN(fallbackRate) ? 0 : fallbackRate;
    }

    /**
     * Calculate NPV and its derivative for Newton-Raphson method
     */
    calculateNPVAndDerivative(cashFlows, rate) {
        const baseDate = cashFlows[0].date;
        let npv = 0;
        let derivative = 0;
        
        for (const cf of cashFlows) {
            const years = (cf.date - baseDate) / (365.25 * 24 * 60 * 60 * 1000);
            const factor = Math.pow(1 + rate, years);
            
            npv += cf.amount / factor;
            derivative -= (cf.amount * years) / (factor * (1 + rate));
        }
        
        return { npv, derivative };
    }

    /**
     * Fallback XIRR calculation using simple approximation
     */
    calculateFallbackXIRR(cashFlows) {
        if (cashFlows.length < 2) return 0;
        
        const totalInvested = Math.abs(cashFlows.filter(cf => cf.amount < 0).reduce((sum, cf) => sum + cf.amount, 0));
        const totalReturned = cashFlows.filter(cf => cf.amount > 0).reduce((sum, cf) => sum + cf.amount, 0);
        
        if (totalInvested <= 0 || totalReturned <= 0) return 0;
        
        const firstDate = cashFlows[0].date;
        const lastDate = cashFlows[cashFlows.length - 1].date;
        const years = (lastDate - firstDate) / (365.25 * 24 * 60 * 60 * 1000);
        
        if (years <= 0) return 0;
        
        // Simple approximation: (Total Return / Total Investment)^(1/years) - 1
        const approximateXIRR = (Math.pow(totalReturned / totalInvested, 1 / years) - 1) * 100;
        
        return approximateXIRR;
    }

    /**
     * Export calculations as object for storage/export
     */
    exportCalculations() {
        return {
            ...this.calculations,
            portfolioSummary: {
                userId: this.portfolioData?.userId,
                totalEntries: this.portfolioData?.entries.length || 0,
                hasTransactions: !!this.transactionData
            },
            settings: {
                currency: this.currency || null,  // Use detected currency, never default to EUR
                decimals: 2
            }
        };
    }
}

// Global calculator instance
const portfolioCalculator = new PortfolioCalculator();