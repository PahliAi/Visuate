/**
 * Portfolio Calculation Engine for Equate
 * Implements all portfolio metrics calculations with manual price override capability
 * UNIFIED: Now uses currencyConverter.js for all timeline currency conversion
 */

// Import currencyConverter for unified timeline conversion
// This replaces the basic calculator conversion logic
// Note: currencyConverter instance is globally available from currencyConverter.js

class PortfolioCalculator {
    /**
     * Create a new PortfolioCalculator instance
     * @constructor
     */
    constructor() {
        this.portfolioData = null;
        this.transactionData = null;
        this.historicalPrices = null; // Timeline data from multiCurrencyPrices (V2 Dynamic Currency System)
        
        // V2 Dynamic Currency System - No more data copies needed!
        
        this.currentPrice = null;
        this.priceSource = 'historical'; // 'historical' or 'manual'
        this.calculations = null;
        this.currency = null; // Set externally by FileAnalyzer
        this.calculationBreakdowns = null; // Detailed breakdown for Calculations tab
        
        
        // Enhanced Reference Points with pre-calculated multi-currency prices (V2 Dynamic Currency System)
        this.enhancedReferencePoints = null; // Built by ReferencePointsBuilder with instant O(1) currency switching
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
    async setPortfolioData(portfolioData, transactionData = null, currency = null, options = {}) {
        console.log('🚀 V2 Dynamic Currency System: Building enhanced reference points...');
        const startTime = performance.now();
        
        // Store basic data
        this.portfolioData = portfolioData;
        this.transactionData = transactionData;
        this.currency = currency;
        
        // Build enhanced reference points with pre-calculated multi-currency prices
        // Pass currency data from options to avoid duplicate database calls
        this.enhancedReferencePoints = await referencePointsBuilder.buildReferencePoints(
            portfolioData, 
            transactionData,
            options.currencyData || null
        );
        
        
        // Set initial currency (reference points start in original portfolio currency)
        if (currency) {
            referencePointsBuilder.changeCurrency(this.enhancedReferencePoints, currency);
        }
        
        const endTime = performance.now();
        console.log(`✨ V2 Enhanced reference points ready: ${this.enhancedReferencePoints.length} points in ${(endTime - startTime).toFixed(1)}ms`);
        
        // Cache enhanced reference points for debugging
        if (window.CacheManager && window.CacheManager.isAvailable()) {
            window.CacheManager.set('enhancedReferencePoints', this.enhancedReferencePoints, 'V2 Enhanced reference points with pre-calculated multi-currency prices');
        }
        
        config.debug('🔧 Calculator received data:', {
            portfolioEntries: portfolioData?.entries?.length || 0,
            transactionEntries: transactionData?.entries?.length || 0,
            hasTransactionData: !!transactionData,
            currency: currency,
            enhancedPointsCount: this.enhancedReferencePoints?.length || 0
        });
    }



    /**
     * V2 Dynamic Currency System: Change display currency instantly
     * Uses O(1) pointer updates on pre-calculated multi-currency prices
     */
    changeCurrency(newCurrency) {
        if (!this.enhancedReferencePoints || this.enhancedReferencePoints.length === 0) {
            console.warn('No enhanced reference points available for currency switching');
            return;
        }

        console.log(`🚀 V2 Instant currency switching to ${newCurrency}...`);
        const startTime = performance.now();
        
        // Use ReferencePointsBuilder for instant O(1) currency switching
        referencePointsBuilder.changeCurrency(this.enhancedReferencePoints, newCurrency);
        
        // Update calculator's current currency
        this.currency = newCurrency;
        
        const endTime = performance.now();
        console.log(`✨ V2 Currency switch completed in ${(endTime - startTime).toFixed(1)}ms`);
    }

    /**
     * Get enhanced reference points for UI components
     */
    getEnhancedReferencePoints() {
        return this.enhancedReferencePoints;
    }

    /**
     * Get available currencies from enhanced reference points
     */
    getAvailableCurrencies() {
        if (!this.enhancedReferencePoints || this.enhancedReferencePoints.length === 0) {
            return [];
        }
        return referencePointsBuilder.getAvailableCurrencies(this.enhancedReferencePoints);
    }

    /**
     * V2 System Status: Verify that V2 Dynamic Currency System is active
     */
    getV2SystemStatus() {
        return {
            isV2Active: !!this.enhancedReferencePoints,
            enhancedPointsCount: this.enhancedReferencePoints?.length || 0,
            currentCurrency: this.currency,
            availableCurrencies: this.getAvailableCurrencies(),
            hasHistoricalPrices: !!this.historicalPrices && this.historicalPrices.length > 0,
            historicalPricesCount: this.historicalPrices?.length || 0
        };
    }

    /**
     * V2 Enhanced Reference Points: Get reference points for UI components
     * Provides unified reference points for timeline, calculations, and other consumers
     */
    getReferencePoints() {
        if (!this.enhancedReferencePoints) {
            return [];
        }
        
        // Return enhanced reference points with current currency prices
        return this.enhancedReferencePoints.map(point => ({
            date: point.date,
            price: point.currentPrice,
            currency: point.currentCurrency,
            quantity: point.quantity,
            type: point.type,
            category: point.category,
            source: point.source
        }));
    }

    /**
     * Get the latest available currency rate from historical data
     */
    getLatestCurrencyRate(currencyRatesByDate) {
        const dates = Object.keys(currencyRatesByDate).sort((a, b) => new Date(b) - new Date(a));
        return dates.length > 0 ? currencyRatesByDate[dates[0]] : null;
    }

    /**
     * Set historical prices and enhance with current price data
     */
    /**
     * Calculate the timeline date range (earliest purchase to latest price date)
     */
    getTimelineRange() {
        if (!this.enhancedReferencePoints || this.enhancedReferencePoints.length === 0) {
            return null;
        }

        // Find earliest purchase date from enhanced reference points
        const purchaseDates = this.enhancedReferencePoints
            .filter(point => point.date && point.type === 'purchase')
            .map(point => this.parseUTCDate(point.date))
            .filter(date => date !== null);
            
        if (purchaseDates.length === 0) {
            return null;
        }
        
        const earliestPurchaseDate = purchaseDates.reduce((earliest, current) => 
            current < earliest ? current : earliest
        );

        // Find latest price date from historicalPrices (V2 system)
        let latestPriceDate = new Date();
        
        if (this.historicalPrices && this.historicalPrices.length > 0) {
            const sortedPrices = [...this.historicalPrices].sort((a, b) => new Date(b.date) - new Date(a.date));
            latestPriceDate = new Date(sortedPrices[0].date + 'T00:00:00.000Z');
            config.debug(`🎯 V2 Timeline range using enhanced reference points with latest date: ${sortedPrices[0].date}`);
        }

        return {
            startDate: earliestPurchaseDate,
            endDate: latestPriceDate,
            startDateStr: earliestPurchaseDate.toISOString().split('T')[0],
            endDateStr: latestPriceDate.toISOString().split('T')[0]
        };
    }


    /**
     * V2 Dynamic Currency System: Set historical prices for timeline
     * Simplified method that works with multiCurrencyPrices data
     */
    async setHistoricalPrices(historicalPrices) {
        this.historicalPrices = historicalPrices || [];
        
        // Enhance historical prices with AsOfDate reference points
        await this.enhanceHistoricalPrices();
        
        this.updateCurrentPrice();
        
        config.debug(`📈 V2 Historical prices set: ${this.historicalPrices.length} entries`);
        
        // Cache for debugging
        if (window.CacheManager && window.CacheManager.isAvailable()) {
            window.CacheManager.set('historicalPrices', this.historicalPrices, 'V2 Historical share price data for timeline');
        }
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
        
        // Initialize empty historical prices array if null (for Non-Shares companies)
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
     * Enhance historical prices with current portfolio data and manual prices
     * Now async to save AsOfDate to IndexedDB
     */
    async enhanceHistoricalPrices() {
        config.debug('🔧 enhanceHistoricalPrices called - checking conditions...');
        
        if (!this.enhancedReferencePoints) {
            config.debug('⚠️ No enhanced reference points - skipping enhancement');
            return;
        }
        
        config.debug('✅ Enhanced reference points exist:', this.enhancedReferencePoints.length);

        // Add portfolio "As of date" entry if it's more recent than latest historical price
        const asOfDate = this.extractAsOfDate();
        const marketPrice = this.extractMarketPrice();
        
        config.debug(`🔍 DEBUG: AsOfDate enhancement check:`, {
            asOfDate,
            marketPrice,
            currency: this.currency,
            willAdd: asOfDate && marketPrice
        });

        if (asOfDate && marketPrice) {
            // V2: Add AsOfDate reference point regardless of currency view
            // Enhanced reference points system handles multi-currency conversion properly
            const existingEntry = this.historicalPrices.find(p => p.date === asOfDate);
            
            // ALWAYS add AsOfDate entry for green diamond display (separate from timeline data)
            const asOfDateMarker = {
                date: asOfDate,
                price: marketPrice,
                source: 'AsOfDate'
            };
            
            // Remove any existing AsOfDate marker to avoid duplicates
            this.historicalPrices = this.historicalPrices.filter(p => p.source !== 'AsOfDate' || p.date !== asOfDate);
            
            // Add AsOfDate marker for green diamond display
            this.historicalPrices.push(asOfDateMarker);
            config.debug(`🟢 Added AsOfDate green diamond marker: ${asOfDate} = ${this.currency} ${marketPrice}`);
            
            // Handle timeline data separately
            if (!existingEntry) {
                config.debug(`📈 Portfolio AsOfDate ${asOfDate} (${this.currency} ${marketPrice}) not found in historical prices, will use AsOfDate price for timeline`);
                
                // Save to IndexedDB for persistence across sessions
                try {
                    await equateDB.appendHistoricalPrice(asOfDate, marketPrice);
                    config.debug(`✅ AsOfDate price saved to IndexedDB permanently: ${asOfDate} = ${this.currency} ${marketPrice}`);
                } catch (error) {
                    config.error('❌ Failed to save AsOfDate to IndexedDB:', error);
                }
            } else {
                // Historical price already exists for AsOfDate - keep the historical close price for timeline
                config.debug(`📈 Portfolio AsOfDate ${asOfDate} already has historical price ${this.currency} ${existingEntry.price}, keeping historical price for timeline (green diamond shows portfolio price ${this.currency} ${marketPrice})`);

                // Check price difference and log if >1 cent
                const priceDiff = Math.abs(existingEntry.price - marketPrice);
                if (priceDiff > 0.01) {
                    config.warn(`💰 Price difference detected for ${asOfDate}:`);
                    config.warn(`   📈 Historical Close Price: ${this.currency} ${existingEntry.price} (timeline)`);
                    config.warn(`   🟢 Portfolio AsOfDate Price: ${this.currency} ${marketPrice} (green diamond)`);
                    config.warn(`   📊 Difference: ${this.currency} ${priceDiff.toFixed(4)} (${(priceDiff/existingEntry.price*100).toFixed(2)}%)`);
                    config.warn(`   ℹ️  Timeline uses historical close, green diamond shows actual portfolio price`);
                }
            }
        } else {
            config.debug(`⚠️ AsOfDate enhancement skipped:`, {
                hasAsOfDate: !!asOfDate,
                hasMarketPrice: !!marketPrice,
                currency: this.currency
            });
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
            hasAsOfDate: this.historicalPrices.some(p => p.date === asOfDate),
            asOfDateEntry: this.historicalPrices.find(p => p.date === asOfDate),
            priceSource: this.priceSource,
            currency: this.currency
        });
        
    }

    /**
     * Extract "As of date" from portfolio file (cell B3)
     */
    extractAsOfDate() {
        if (!this.enhancedReferencePoints) {
            return null;
        }
        
        // Find AsOfDate from enhanced reference points
        const asOfDatePoint = this.enhancedReferencePoints.find(point => point.type === 'asOfDate');
        return asOfDatePoint ? asOfDatePoint.date : null;
    }

    /**
     * Extract market price from enhanced reference points
     */
    extractMarketPrice() {
        if (!this.enhancedReferencePoints) {
            return null;
        }
        
        // Find AsOfDate from enhanced reference points and get its current price
        const asOfDatePoint = this.enhancedReferencePoints.find(point => point.type === 'asOfDate');
        return asOfDatePoint ? asOfDatePoint.currentPrice : null;
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
            // Fallback: Use market price from enhanced reference points
            const asOfDatePoint = this.enhancedReferencePoints?.find(point => point.type === 'asOfDate');
            if (asOfDatePoint) {
                this.currentPrice = asOfDatePoint.currentPrice;
            }
        }
    }

    /**
     * Calculate all portfolio metrics including returns, CAGR, and current values
     * @returns {Object} Complete portfolio calculations object with all metrics
     * @throws {Error} If portfolio data or current price is missing
     */
    async calculate() {
        if (!this.enhancedReferencePoints || !this.currentPrice) {
            throw new Error('Missing enhanced reference points or current price');
        }

        // Use externally provided currency (from FileAnalyzer)
        const portfolioCurrency = this.currency || 'UNKNOWN';
        
        // Currency validation is now handled by FileAnalyzer during pre-parsing
        let currencyWarning = null;
        
        // Calculate core metrics using V2 Enhanced Reference Points
        
        const userInvestment = this.calculateUserInvestment();
        const companyMatch = this.calculateCompanyMatch();
        const freeShares = this.calculateFreeShares();
        const companyInvestment = companyMatch + freeShares; // Legacy total
        const dividendIncome = this.calculateDividendIncome();
        const totalSold = this.calculateTotalSold();
        const currentValue = this.calculateCurrentValue();
        
        const totalCurrentPortfolio = currentValue; // Current value of remaining shares
        const totalInvestment = userInvestment + companyInvestment + dividendIncome; // Sum of all investments
        const totalValue = currentValue + totalSold; // Current + Sold
        const totalReturn = totalValue - userInvestment; // Total Value - Your Investment (RETURN ON YOUR INVESTMENT)
        const returnPercentage = userInvestment > 0 ? ((totalReturn / userInvestment) * 100) : 0; // RETURN % ON YOUR INVESTMENT
        const returnOnTotalInvestment = this.calculateReturnOnTotalInvestment(totalValue, totalInvestment);
        const returnPercentageOnTotalInvestment = this.calculateReturnPercentageOnTotalInvestment(totalValue, totalInvestment);
        const availableShares = this.calculateAvailableShares();
        const blockedShares = this.calculateBlockedShares(); // Still uses old method for now (only used in PDF export)
        const xirrUserInvestment = await this.calculateXIRRUserInvestment();
        const xirrTotalInvestment = await this.calculateXIRRTotalInvestment();

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

        // Generate detailed calculation breakdowns for Calculations tab
        this.calculationBreakdowns = this.generateCalculationBreakdowns();
        this.calculations.calculationBreakdowns = this.calculationBreakdowns;

        return this.calculations;
    }

    /**
     * Recalculate with current settings
     */
    async recalculate() {
        config.debug('🔄 Recalculate called, has enhanced reference points:', !!this.enhancedReferencePoints);
        if (this.enhancedReferencePoints) {
            config.debug('🔄 Calling calculate() with current price:', this.currentPrice, 'source:', this.priceSource);
            return await this.calculate();
        }
        return null;
    }

    /**
     * Calculate total user investment
     * V2 Enhanced Reference Points: Uses pre-calculated multi-currency prices
     */
    calculateUserInvestment() {
        if (!this.enhancedReferencePoints) {
            return 0;
        }
        
        // All user investments ever made (including those already sold)
        return this.enhancedReferencePoints
            .filter(point => point.type === 'purchase') // Only purchase entries, not AsOfDate
            .filter(point => categorizeEntry(point) === 'userInvestment')
            .reduce((sum, point) => sum + (point.currentPrice * point.allocatedQuantity), 0);
    }

    /**
     * Calculate real company match investment
     * V2 Enhanced Reference Points: Uses pre-calculated multi-currency prices
     */
    calculateCompanyMatch() {
        if (!this.enhancedReferencePoints) {
            return 0;
        }
        
        // Company match contributions for Employee Share Purchase Plan
        return this.enhancedReferencePoints
            .filter(point => categorizeEntry(point) === 'companyMatch')
            .reduce((sum, point) => sum + (point.currentPrice * point.allocatedQuantity), 0);
    }

    /**
     * Calculate free shares investment
     * V2 Enhanced Reference Points: Uses pre-calculated multi-currency prices
     */
    calculateFreeShares() {
        if (!this.enhancedReferencePoints) {
            return 0;
        }
        
        // Award contributions for Free Share plan
        return this.enhancedReferencePoints
            .filter(point => categorizeEntry(point) === 'freeShares')
            .reduce((sum, point) => sum + (point.currentPrice * point.allocatedQuantity), 0);
    }

    /**
     * Calculate total company investment (Company match + Free shares)
     * V2 Enhanced Reference Points: Uses new parameter-less methods
     */
    calculateCompanyInvestment() {
        return this.calculateCompanyMatch() + this.calculateFreeShares();
    }

    /**
     * Calculate total dividend income
     * V2 Enhanced Reference Points: Uses pre-calculated multi-currency prices
     */
    calculateDividendIncome() {
        if (!this.enhancedReferencePoints) {
            return 0;
        }
        
        // All dividend reinvestments ever made (including those already sold)
        return this.enhancedReferencePoints
            .filter(point => categorizeEntry(point) === 'dividendIncome')
            .reduce((sum, point) => sum + (point.currentPrice * point.allocatedQuantity), 0);
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
    /**
     * Calculate current portfolio value 
     * V2 Enhanced Reference Points: Uses outstandingQuantity from enhanced reference points
     */
    calculateCurrentValue() {
        if (!this.enhancedReferencePoints) {
            return 0;
        }
        
        if (this.priceSource === 'manual' && this.currentPrice) {
            // Calculate using manual price: outstanding shares * manual price
            const totalValue = this.enhancedReferencePoints
                .filter(point => point.outstandingQuantity > 0)
                .reduce((sum, point) => sum + (point.outstandingQuantity * this.currentPrice), 0);
            config.debug('💰 V2 Manual price calculation result:', totalValue);
            return totalValue;
        } else if (this.currentPrice) {
            // Calculate using most recent price: outstanding shares * current price  
            const totalValue = this.enhancedReferencePoints
                .filter(point => point.outstandingQuantity > 0)
                .reduce((sum, point) => sum + (point.outstandingQuantity * this.currentPrice), 0);
            config.debug(`💰 V2 Current price calculation result: ${totalValue} (using price ${this.currentPrice})`);
            return totalValue;
        } else {
            // Fallback to market price from AsOfDate reference point
            const asOfDatePoint = this.enhancedReferencePoints.find(point => point.type === 'asOfDate');
            if (asOfDatePoint) {
                const totalValue = this.enhancedReferencePoints
                    .filter(point => point.outstandingQuantity > 0)
                    .reduce((sum, point) => sum + (point.outstandingQuantity * asOfDatePoint.currentPrice), 0);
                config.debug('💰 V2 Fallback to AsOfDate market price');
                return totalValue;
            }
            return 0;
        }
    }

    /**
     * Calculate Compound Annual Growth Rate (CAGR)
     */
    calculateCAGR(currentValue) {
        // Find the earliest investment date (including sold shares)
        const dates = this.enhancedReferencePoints
            .filter(point => point.date)
            .map(point => this.parseUTCDate(point.date))
            .filter(date => date !== null)
            .sort((a, b) => a - b);

        if (dates.length === 0) return 0;

        const startDate = dates[0];
        const endDate = new Date();
        const years = (endDate - startDate) / (365.25 * 24 * 60 * 60 * 1000);

        if (years <= 0) return 0;

        // Calculate user investment only (company match and dividends are returns, not investments)
        const userInvestment = this.calculateUserInvestment();

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
     * V2 Enhanced Reference Points: Uses outstandingQuantity from enhanced reference points
     */
    calculateAvailableShares() {
        if (!this.enhancedReferencePoints) {
            return 0;
        }
        
        // Total shares owned = sum of Outstanding Quantity column
        const total = this.enhancedReferencePoints
            .reduce((sum, point) => sum + (point.outstandingQuantity || 0), 0);
        return Math.round(total * 1000) / 1000; // Round to 3 decimals
    }

    /**
     * Calculate blocked shares by unlock date
     * V2 Enhanced Reference Points: Uses enhanced reference points (fallback to raw data if not available)
     */
    calculateBlockedShares() {
        const blocked = {
            total: 0,
            byYear: {}
        };

        // Use enhanced reference points if available, fallback to raw data for PDF export
        const dataSource = this.enhancedReferencePoints || (this.portfolioData?.entries || []);
        
        dataSource
            .filter(item => (item.outstandingQuantity || 0) > (item.availableQuantity || 0))
            .forEach(item => {
                const blockedQty = (item.outstandingQuantity || 0) - (item.availableQuantity || 0);
                
                if (item.availableFrom) {
                    const availableFromDate = this.parseUTCDate(item.availableFrom);
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
        

        if (!this.enhancedReferencePoints || !this.historicalPrices || this.historicalPrices.length === 0) {
            return [];
        }

        // DEBUG: Check what historical prices we actually have when timeline is generated
        const asOfDate = this.extractAsOfDate();
        config.debug('🔍 DEBUG getPortfolioTimeline - historicalPrices:', {
            count: this.historicalPrices.length,
            lastFew: this.historicalPrices.slice(-3),
            hasAsOfDate: asOfDate ? this.historicalPrices.some(p => p.date === asOfDate) : false,
            asOfDateEntry: asOfDate ? this.historicalPrices.find(p => p.date === asOfDate) : null
        });

        // Step 1: Get timeline date range
        const timelineRange = this.getTimelineRange();
        if (!timelineRange) {
            return [];
        }

        // Step 2: Create lookups by date (using ReferencePointsBuilder for optimization potential)
        // Group transactions by date
        const transactionsByDate = {};
        if (this.transactionData?.entries) {
            this.transactionData.entries.forEach(transaction => {
                const date = transaction.transactionDate;
                if (!transactionsByDate[date]) {
                    transactionsByDate[date] = [];
                }
                transactionsByDate[date].push(transaction);
            });
        }
        
        // Group purchases by date from enhanced reference points
        const purchasesByDate = {};
        if (this.enhancedReferencePoints) {
            this.enhancedReferencePoints
                .filter(point => point.type === 'purchase')
                .forEach(point => {
                    const date = point.date;
                    if (!purchasesByDate[date]) {
                        purchasesByDate[date] = [];
                }
                purchasesByDate[date].push(point);
            });
        }

        // Step 3: No price lookup needed - V2 uses actual prices directly from multiCurrencyPrices

        // Step 4: V2 COMPLIANT - Generate sparse timeline using only actual multiCurrencyPrices dates
        // NO daily interpolation - let Plotly connect the dots automatically for smooth lines
        const timeline = [];
        
        config.debug('🚀 V2 TIMELINE: Generating sparse timeline from', this.historicalPrices.length, 'actual price dates (no weekend filling)');
        
        let outstandingShares = 0; // Track actual shares currently owned (not cumulative)
        let cumulativeUserInvestment = 0; // Track total user money invested (never decreases)
        let cumulativeSaleProceeds = 0; // Track total money received from sales
        
        // Sort historical prices by date to ensure chronological processing
        const sortedPrices = [...this.historicalPrices].sort((a, b) => new Date(a.date) - new Date(b.date));
        
        // Build cumulative share tracking - we need to know shares at each price date
        const sharesAtDate = {};
        let runningShares = 0;
        let runningInvestment = 0;
        let runningProceeds = 0;
        
        // First pass: Calculate cumulative shares for all dates chronologically
        const allDates = new Set([
            ...Object.keys(purchasesByDate),
            ...Object.keys(transactionsByDate),
            ...sortedPrices.map(p => p.date)
        ]);
        
        const sortedAllDates = Array.from(allDates).sort();
        
        for (const dateStr of sortedAllDates) {
            let shareChange = 0;
            let userInvestmentChange = 0;
            let saleProceedsChange = 0;
            
            // Process purchases
            if (purchasesByDate[dateStr]) {
                purchasesByDate[dateStr].forEach(purchase => {
                    shareChange += purchase.allocatedQuantity;
                    
                    if (categorizeEntry(purchase) === 'userInvestment') {
                        userInvestmentChange += purchase.costBasis * purchase.allocatedQuantity;
                    }
                });
            }
            
            // Process sales
            if (transactionsByDate[dateStr]) {
                transactionsByDate[dateStr].forEach(transaction => {
                    if (this.isSellOrTransferOrderType(transaction.orderType)) {
                        const quantitySold = Math.abs(transaction.quantity);
                        shareChange -= quantitySold;
                        
                        const saleAmount = quantitySold * (transaction.executionPrice || 0);
                        saleProceedsChange += saleAmount;
                    }
                });
            }
            
            runningShares += shareChange;
            runningInvestment += userInvestmentChange;
            runningProceeds += saleProceedsChange;
            
            sharesAtDate[dateStr] = {
                shares: runningShares,
                investment: runningInvestment,
                proceeds: runningProceeds,
                hasTransaction: shareChange !== 0
            };
        }
        
        // Second pass: Create timeline entries ONLY for actual price dates
        for (const priceEntry of sortedPrices) {
            const dateStr = priceEntry.date;
            
            // Get share position at this date
            const shareInfo = sharesAtDate[dateStr] || { shares: 0, investment: 0, proceeds: 0, hasTransaction: false };
            
            // Find the most recent share position if no exact match
            if (!sharesAtDate[dateStr]) {
                // Find most recent date with share info
                const pastDates = Object.keys(sharesAtDate)
                    .filter(d => d <= dateStr)
                    .sort()
                    .reverse();
                
                if (pastDates.length > 0) {
                    const mostRecentDate = pastDates[0];
                    shareInfo.shares = sharesAtDate[mostRecentDate].shares;
                    shareInfo.investment = sharesAtDate[mostRecentDate].investment;
                    shareInfo.proceeds = sharesAtDate[mostRecentDate].proceeds;
                }
            }
            
            // Build reason string for this date
            let reason = '';
            if (purchasesByDate[dateStr]) {
                purchasesByDate[dateStr].forEach(purchase => {
                    reason += `Bought ${purchase.allocatedQuantity} shares (${purchase.contributionType}); `;
                });
            }
            if (transactionsByDate[dateStr]) {
                transactionsByDate[dateStr].forEach(transaction => {
                    if (this.isSellOrTransferOrderType(transaction.orderType)) {
                        const quantitySold = Math.abs(transaction.quantity);
                        const saleAmount = quantitySold * (transaction.executionPrice || 0);
                        const currencySymbol = currencyConverter.getCurrencySymbol(this.currency);
                        reason += `Sold ${quantitySold} shares for ${currencySymbol}${saleAmount.toFixed(2)}; `;
                    }
                });
            }
            
            // Calculate portfolio value using actual price from multiCurrencyPrices (no interpolation!)
            const currentPrice = priceEntry.price;
            const portfolioValue = shareInfo.shares * currentPrice;
            
            // Calculate profit/loss
            const profitLoss = portfolioValue + shareInfo.proceeds - shareInfo.investment;

            // Create timeline entry ONLY for actual price dates
            timeline.push({
                date: dateStr,
                sharesPrevDay: shareInfo.shares, // V2: We don't track daily changes in sparse timeline
                sharesCurrentDay: shareInfo.shares,
                portfolioValue: portfolioValue,
                reason: reason.trim() || null,
                profitLoss: profitLoss,
                cumulativeUserInvestment: shareInfo.investment,
                cumulativeSaleProceeds: shareInfo.proceeds,
                currentPrice: currentPrice,
                hasTransaction: shareInfo.hasTransaction
            });

            // DEBUG: Check if we're processing the asOfDate
            if (asOfDate && dateStr === asOfDate) {
                config.debug(`🎯 V2 DEBUG: Processing asOfDate ${asOfDate} in sparse timeline:`, {
                    portfolioValue: portfolioValue,
                    currentPrice: currentPrice,
                    outstandingShares: shareInfo.shares,
                    timelineLength: timeline.length
                });
            }
        }

        // DEBUG: Final timeline state
        config.debug('🔍 DEBUG final timeline:', {
            totalDays: timeline.length,
            lastFewDays: timeline.slice(-3).map(d => ({date: d.date, portfolioValue: d.portfolioValue})),
            hasAsOfDate: asOfDate ? timeline.some(d => d.date === asOfDate) : false,
            asOfDateEntry: asOfDate ? timeline.find(d => d.date === asOfDate) : null
        });
        
        // Cache portfolio timeline for debugging
        if (window.CacheManager && window.CacheManager.isAvailable()) {
            window.CacheManager.set('portfolioTimeline', timeline, 'V2 Sparse timeline data (no daily interpolation) - only actual multiCurrencyPrices dates');
        }
        
        return timeline;
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

    /**
     * Clear all data
     */
    clearData() {
        this.portfolioData = null;
        this.transactionData = null;
        this.historicalPrices = null;
        this.enhancedReferencePoints = null;
        this.currentPrice = null;
        this.priceSource = 'historical';
        this.calculations = null;
        
        // Clear cached data as well (but preserve hist.xlsx data)
        if (window.CacheManager && window.CacheManager.isAvailable()) {
            window.CacheManager.clearAll();
        }
        
        // Clear user portfolio/transaction data from IndexedDB (but preserve hist.xlsx data)
        if (typeof equateDB !== 'undefined' && equateDB.db) {
            // Clear user data stores only - preserve hist.xlsx related stores
            equateDB.clearPortfolioData();
            equateDB.clearManualPrices();
            equateDB.clearHistoricalPrices(); // This clears legacy format historical prices (not hist.xlsx cache)
            equateDB.clearHistoricalCurrencyData();
            equateDB.clearCurrencyQualityScores();
            config.debug('🗑️ User data cleared');
        }
    }


    /**
     * Calculate XIRR based only on user investment (excluding company contributions)
     * Uses direct Excel data for consistency
     */
    /**
     * Calculate XIRR for user investment only
     * V2 Enhanced Reference Points: Uses enhanced reference points for all calculations
     */
    async calculateXIRRUserInvestment() {
        try {
            if (!this.enhancedReferencePoints) {
                return 0;
            }
            
            const cashFlows = [];
            
            // Add all user purchases as negative cash flows (from enhanced reference points)
            this.enhancedReferencePoints.forEach(point => {
                if (point.date && categorizeEntry(point) === 'userInvestment') {
                    
                    const date = this.parseUTCDate(point.date);
                    if (date) {
                        const amount = -(point.currentPrice * point.allocatedQuantity);
                        cashFlows.push({
                            date: date,
                            amount: amount,
                            type: 'user_investment',
                            description: `User purchase: ${point.allocatedQuantity} shares`
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
            const currentPortfolioValue = this.calculateCurrentValue();
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
    /**
     * Calculate XIRR for total investment (user + company contributions)
     * V2 Enhanced Reference Points: Uses enhanced reference points for all calculations
     */
    async calculateXIRRTotalInvestment() {
        try {
            if (!this.enhancedReferencePoints) {
                return 0;
            }
            
            const cashFlows = [];
            
            // Add ALL contributions as negative cash flows (from enhanced reference points)
            this.enhancedReferencePoints.forEach(point => {
                if (point.date && point.costBasis && point.allocatedQuantity) {
                    const date = this.parseUTCDate(point.date);
                    if (date) {
                        const amount = -(point.currentPrice * point.allocatedQuantity);
                        let contributionDescription = '';
                        
                        // Categorize contribution type
                        if (categorizeEntry(point) === 'userInvestment') {
                            contributionDescription = `User purchase: ${point.allocatedQuantity} shares`;
                        } else if (categorizeEntry(point) === 'companyMatch') {
                            contributionDescription = `Company match: ${point.allocatedQuantity} shares`;
                        } else if (categorizeEntry(point) === 'freeShares') {
                            contributionDescription = `Free shares award: ${point.allocatedQuantity} shares`;
                        } else if (categorizeEntry(point) === 'dividendIncome') {
                            contributionDescription = `Dividend reinvestment: ${point.allocatedQuantity} shares`;
                        } else {
                            contributionDescription = `${point.contributionType} (${point.plan}): ${point.allocatedQuantity} shares`;
                        }
                        
                        cashFlows.push({
                            date: date,
                            amount: amount,
                            type: 'total_investment',
                            contributionType: point.contributionType,
                            plan: point.plan,
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
            const currentPortfolioValue = this.calculateCurrentValue();
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
                userId: this.enhancedReferencePoints?.find(p => p.source === 'portfolio')?.userId || 'unknown',
                totalReferencePoints: this.enhancedReferencePoints?.length || 0,
                purchasePoints: this.enhancedReferencePoints?.filter(p => p.type === 'purchase').length || 0,
                hasTransactions: !!this.transactionData
            },
            settings: {
                currency: this.currency || null,  // Use detected currency, never default to EUR
                decimals: 2
            }
        };
    }

    /**
     * Generate detailed calculation breakdowns for each metric card
     * @returns {Object} Breakdown data organized by card type
     */
    generateCalculationBreakdowns() {
        const breakdowns = {
            userInvestment: this.getBreakdownUserInvestment(),
            companyMatch: this.getBreakdownCompanyMatch(),
            freeShares: this.getBreakdownFreeShares(),
            dividendIncome: this.getBreakdownDividendIncome(),
            totalInvestment: this.getBreakdownTotalInvestment(),
            currentPortfolio: this.getBreakdownCurrentPortfolio(),
            totalSold: this.getBreakdownTotalSold(),
            xirrUserInvestment: this.getBreakdownXIRRUserInvestment(),
            xirrTotalInvestment: this.getBreakdownXIRRTotalInvestment()
        };

        return breakdowns;
    }

    /**
     * Get breakdown for Your Investment card
     */
    getBreakdownUserInvestment() {
        const breakdown = [];
        
        for (const point of this.enhancedReferencePoints) {
            if (categorizeEntry(point) === 'userInvestment' && 
                parseFloat(point.allocatedQuantity) > 0) {
                
                const date = this.parseUTCDate(point.date);
                const shares = parseFloat(point.allocatedQuantity);
                const investment = parseFloat(point.costBasis) * shares;
                const currentValue = shares * this.currentPrice;
                
                breakdown.push({
                    date: date ? date.toISOString().split('T')[0] : point.date,
                    action: 'Purchase',
                    category: 'Employee Purchase',
                    shares: shares,
                    investment: investment,
                    currentValue: currentValue,
                    cardType: 'userInvestment'
                });
            }
        }

        // Sort by date
        breakdown.sort((a, b) => new Date(a.date) - new Date(b.date));
        
        // Add total row
        const totalShares = breakdown.reduce((sum, item) => sum + item.shares, 0);
        const totalInvestment = breakdown.reduce((sum, item) => sum + item.investment, 0);
        const totalCurrentValue = breakdown.reduce((sum, item) => sum + item.currentValue, 0);
        
        breakdown.push({
            date: null,
            action: 'Total',
            category: null,
            shares: totalShares,
            investment: totalInvestment,
            currentValue: totalCurrentValue,
            isTotal: true
        });

        return breakdown;
    }

    /**
     * Get breakdown for Company Match card
     */
    getBreakdownCompanyMatch() {
        const breakdown = [];
        
        for (const point of this.enhancedReferencePoints) {
            if (categorizeEntry(point) === 'companyMatch' && 
                parseFloat(point.allocatedQuantity) > 0) {
                
                const date = this.parseUTCDate(point.date);
                const shares = parseFloat(point.allocatedQuantity);
                const investment = parseFloat(point.costBasis) * shares;
                const currentValue = shares * this.currentPrice;
                
                breakdown.push({
                    date: date ? date.toISOString().split('T')[0] : point.date,
                    action: 'Award',
                    category: 'Company Match Award',
                    shares: shares,
                    investment: investment,
                    currentValue: currentValue,
                    cardType: 'companyMatch'
                });
            }
        }

        breakdown.sort((a, b) => new Date(a.date) - new Date(b.date));
        
        const totalShares = breakdown.reduce((sum, item) => sum + item.shares, 0);
        const totalInvestment = breakdown.reduce((sum, item) => sum + item.investment, 0);
        const totalCurrentValue = breakdown.reduce((sum, item) => sum + item.currentValue, 0);
        
        breakdown.push({
            date: null,
            action: 'Total',
            category: null,
            shares: totalShares,
            investment: totalInvestment,
            currentValue: totalCurrentValue,
            isTotal: true
        });

        return breakdown;
    }

    /**
     * Get breakdown for Free Shares card
     */
    getBreakdownFreeShares() {
        const breakdown = [];
        
        for (const point of this.enhancedReferencePoints) {
            if (categorizeEntry(point) === 'freeShares' && 
                parseFloat(point.allocatedQuantity) > 0) {
                
                const date = this.parseUTCDate(point.date);
                const shares = parseFloat(point.allocatedQuantity);
                const investment = parseFloat(point.costBasis) * shares;
                const currentValue = shares * this.currentPrice;
                
                breakdown.push({
                    date: date ? date.toISOString().split('T')[0] : point.date,
                    action: 'Award',
                    category: 'Free Share Award',
                    shares: shares,
                    investment: investment,
                    currentValue: currentValue,
                    cardType: 'freeShares'
                });
            }
        }

        breakdown.sort((a, b) => new Date(a.date) - new Date(b.date));
        
        const totalShares = breakdown.reduce((sum, item) => sum + item.shares, 0);
        const totalInvestment = breakdown.reduce((sum, item) => sum + item.investment, 0);
        const totalCurrentValue = breakdown.reduce((sum, item) => sum + item.currentValue, 0);
        
        breakdown.push({
            date: null,
            action: 'Total',
            category: null,
            shares: totalShares,
            investment: totalInvestment,
            currentValue: totalCurrentValue,
            isTotal: true
        });

        return breakdown;
    }

    /**
     * Get breakdown for Dividend Income card
     */
    getBreakdownDividendIncome() {
        const breakdown = [];
        
        for (const point of this.enhancedReferencePoints) {
            if (categorizeEntry(point) === 'dividendIncome' && 
                parseFloat(point.allocatedQuantity) > 0) {
                
                const date = this.parseUTCDate(point.date);
                const shares = parseFloat(point.allocatedQuantity);
                const investment = parseFloat(point.costBasis) * shares;
                const currentValue = shares * this.currentPrice;
                
                breakdown.push({
                    date: date ? date.toISOString().split('T')[0] : point.date,
                    action: 'Dividend',
                    category: 'Dividend Reinvestment',
                    shares: shares,
                    investment: investment,
                    currentValue: currentValue,
                    cardType: 'dividendIncome'
                });
            }
        }

        breakdown.sort((a, b) => new Date(a.date) - new Date(b.date));
        
        const totalShares = breakdown.reduce((sum, item) => sum + item.shares, 0);
        const totalInvestment = breakdown.reduce((sum, item) => sum + item.investment, 0);
        const totalCurrentValue = breakdown.reduce((sum, item) => sum + item.currentValue, 0);
        
        breakdown.push({
            date: null,
            action: 'Total',
            category: null,
            shares: totalShares,
            investment: totalInvestment,
            currentValue: totalCurrentValue,
            isTotal: true
        });

        return breakdown;
    }

    /**
     * Get breakdown for Total Investment card (summary format)
     */
    getBreakdownTotalInvestment() {
        const userInvestment = this.calculateUserInvestment();
        const companyMatch = this.calculateCompanyMatch();
        const freeShares = this.calculateFreeShares();
        const dividendIncome = this.calculateDividendIncome();
        const totalInvestment = userInvestment + companyMatch + freeShares + dividendIncome;

        // Count transactions for each category
        const userTransactions = this.enhancedReferencePoints.filter(e => categorizeEntry(e) === 'userInvestment' && parseFloat(e.allocatedQuantity) > 0).length;
        const companyTransactions = this.enhancedReferencePoints.filter(e => categorizeEntry(e) === 'companyMatch').length;
        const freeShareTransactions = this.enhancedReferencePoints.filter(e => categorizeEntry(e) === 'freeShares').length;
        const dividendTransactions = this.enhancedReferencePoints.filter(e => categorizeEntry(e) === 'dividendIncome').length;

        // Calculate shares for each category
        const userShares = this.enhancedReferencePoints.filter(e => categorizeEntry(e) === 'userInvestment' && parseFloat(e.allocatedQuantity) > 0)
            .reduce((sum, e) => sum + parseFloat(e.allocatedQuantity), 0);
        const companyShares = this.enhancedReferencePoints.filter(e => categorizeEntry(e) === 'companyMatch')
            .reduce((sum, e) => sum + parseFloat(e.allocatedQuantity), 0);
        const freeShareShares = this.enhancedReferencePoints.filter(e => categorizeEntry(e) === 'freeShares')
            .reduce((sum, e) => sum + parseFloat(e.allocatedQuantity), 0);
        const dividendShares = this.enhancedReferencePoints.filter(e => categorizeEntry(e) === 'dividendIncome')
            .reduce((sum, e) => sum + parseFloat(e.allocatedQuantity), 0);

        const breakdown = [
            {
                category: 'Your Investment',
                transactions: userTransactions,
                shares: userShares,
                investment: userInvestment,
                percentage: totalInvestment > 0 ? ((userInvestment / totalInvestment) * 100) : 0
            },
            {
                category: 'Company Match',
                transactions: companyTransactions,
                shares: companyShares,
                investment: companyMatch,
                percentage: totalInvestment > 0 ? ((companyMatch / totalInvestment) * 100) : 0
            },
            {
                category: 'Free Shares',
                transactions: freeShareTransactions,
                shares: freeShareShares,
                investment: freeShares,
                percentage: totalInvestment > 0 ? ((freeShares / totalInvestment) * 100) : 0
            },
            {
                category: 'Dividend Income',
                transactions: dividendTransactions,
                shares: dividendShares,
                investment: dividendIncome,
                percentage: totalInvestment > 0 ? ((dividendIncome / totalInvestment) * 100) : 0
            }
        ];

        // Add total row
        breakdown.push({
            category: 'Total',
            transactions: userTransactions + companyTransactions + freeShareTransactions + dividendTransactions,
            shares: userShares + companyShares + freeShareShares + dividendShares,
            investment: totalInvestment,
            percentage: 100.0,
            isTotal: true
        });

        return breakdown;
    }

    /**
     * Get breakdown for Current Portfolio card (special dual format)
     */
    getBreakdownCurrentPortfolio() {
        // Use outstanding quantity for current portfolio (shares still owned, not originally allocated)
        const totalShares = this.enhancedReferencePoints.reduce((sum, point) => {
            return sum + parseFloat(point.outstandingQuantity || 0);
        }, 0);
        
        const totalInvestment = this.calculateUserInvestment() + 
                              this.calculateCompanyMatch() + 
                              this.calculateFreeShares() + 
                              this.calculateDividendIncome();
        
        const currentValue = totalShares * this.currentPrice;
        
        // Get portfolio AsOfDate for split calculation
        const asOfDate = this.extractAsOfDate() || new Date().toISOString().split('T')[0];
        const marketPriceAtDownload = this.extractMarketPrice();
        const portfolioDownloadValue = totalShares * marketPriceAtDownload;
        
        config.debug('🔍 Current Portfolio price comparison:', {
            totalShares,
            marketPriceAtDownload,
            currentPrice: this.currentPrice,
            portfolioDownloadValue,
            currentValue,
            asOfDate,
            firstPurchasePoint: this.enhancedReferencePoints?.find(p => p.type === 'purchase'),
            asOfDatePoint: this.enhancedReferencePoints?.find(p => p.type === 'asOfDate')
        });
        
        const profitLossSinceDownload = currentValue - portfolioDownloadValue;
        
        const breakdown = [
            {
                date: asOfDate,
                category: `Portfolio value at download date`,
                shares: totalShares,
                investment: totalInvestment,
                currentValue: portfolioDownloadValue
            },
            {
                date: new Date().toISOString().split('T')[0],
                category: `Portfolio value today`,
                shares: null, // Don't repeat shares
                investment: null,
                currentValue: currentValue
            },
            {
                date: null,
                category: `Profit/Loss since portfolio download`,
                shares: null,
                investment: null,
                currentValue: profitLossSinceDownload,
                isProfitLoss: true
            }
        ];


        return breakdown;
    }

    /**
     * Get breakdown for Total Sold card
     */
    getBreakdownTotalSold() {
        if (!this.transactionData || !this.transactionData.entries) {
            return [];
        }

        const breakdown = [];
        
        for (const transaction of this.transactionData.entries) {
            if (this.isSellOrTransferOrderType(transaction.orderType) && 
                transaction.status === 'Executed') {
                
                const date = this.parseUTCDate(transaction.transactionDate);
                const shares = Math.abs(parseFloat(transaction.quantity));
                const price = parseFloat(transaction.executionPrice);
                const proceeds = shares * price;
                
                breakdown.push({
                    date: date ? date.toISOString().split('T')[0] : transaction.transactionDate,
                    action: transaction.orderType,
                    category: transaction.orderType.includes('Sell') ? 'Share Sale' : 'Transfer Out',
                    shares: shares,
                    price: price,
                    proceeds: proceeds
                });
            }
        }

        breakdown.sort((a, b) => new Date(a.date) - new Date(b.date));
        
        const totalShares = breakdown.reduce((sum, item) => sum + item.shares, 0);
        const totalProceeds = breakdown.reduce((sum, item) => sum + item.proceeds, 0);
        const avgPrice = totalShares > 0 ? totalProceeds / totalShares : 0;
        
        breakdown.push({
            date: null,
            action: 'Total',
            category: null,
            shares: totalShares,
            price: avgPrice,
            proceeds: totalProceeds,
            isTotal: true
        });

        return breakdown;
    }

    /**
     * Get breakdown for XIRR User Investment card
     */
    getBreakdownXIRRUserInvestment() {
        const cashFlows = [];
        
        // Add ONLY user purchases as negative cash flows (investments)
        for (const point of this.enhancedReferencePoints) {
            if (point.date && 
                categorizeEntry(point) === 'userInvestment' && 
                parseFloat(point.allocatedQuantity) > 0) {
                
                const date = this.parseUTCDate(point.date);
                const investment = parseFloat(point.costBasis) * parseFloat(point.allocatedQuantity);
                
                cashFlows.push({
                    date: date ? date.toISOString().split('T')[0] : point.date,
                    category: 'Investment',
                    cashFlow: -investment,
                    description: 'Your Purchase'
                });
            }
        }

        // Add sale proceeds (positive)
        if (this.transactionData && this.transactionData.entries) {
            for (const transaction of this.transactionData.entries) {
                if (this.isSellOrTransferOrderType(transaction.orderType) && 
                    transaction.status === 'Executed') {
                    
                    const date = this.parseUTCDate(transaction.transactionDate);
                    const shares = Math.abs(parseFloat(transaction.quantity));
                    const price = parseFloat(transaction.executionPrice);
                    const proceeds = shares * price;
                    
                    cashFlows.push({
                        date: date ? date.toISOString().split('T')[0] : transaction.transactionDate,
                        category: 'Return',
                        cashFlow: proceeds,
                        description: 'Share Sale Proceeds'
                    });
                }
            }
        }

        // Add current portfolio value (positive) - total value of all outstanding shares
        const currentValue = this.calculateCurrentValue();
        
        if (currentValue > 0) {
            cashFlows.push({
                date: new Date().toISOString().split('T')[0],
                category: 'Return',
                cashFlow: currentValue,
                description: 'Current Portfolio Value'
            });
        }

        // Sort by date
        cashFlows.sort((a, b) => new Date(a.date) - new Date(b.date));

        return cashFlows;
    }

    /**
     * Get breakdown for XIRR Total Investment card
     */
    getBreakdownXIRRTotalInvestment() {
        const cashFlows = [];
        
        // Add all investment cash flows (negative)
        for (const point of this.enhancedReferencePoints) {
            if (parseFloat(point.allocatedQuantity) > 0) {
                const date = this.parseUTCDate(point.date);
                const investment = parseFloat(point.costBasis) * parseFloat(point.allocatedQuantity);
                let description = '';
                
                if (categorizeEntry(point) === 'userInvestment') {
                    description = 'Your Purchase';
                } else if (categorizeEntry(point) === 'companyMatch') {
                    description = 'Company Match';
                } else if (categorizeEntry(point) === 'freeShares') {
                    description = 'Free Shares';
                } else if (categorizeEntry(point) === 'dividendIncome') {
                    description = 'Dividend Reinvestment';
                } else {
                    description = 'Investment';
                }
                
                cashFlows.push({
                    date: date ? date.toISOString().split('T')[0] : point.date,
                    category: 'Investment',
                    cashFlow: -investment,
                    description: description
                });
            }
        }

        // Add sale proceeds (positive)
        if (this.transactionData && this.transactionData.entries) {
            for (const transaction of this.transactionData.entries) {
                if (this.isSellOrTransferOrderType(transaction.orderType) && 
                    transaction.status === 'Executed') {
                    
                    const date = this.parseUTCDate(transaction.transactionDate);
                    const shares = Math.abs(parseFloat(transaction.quantity));
                    const price = parseFloat(transaction.executionPrice);
                    const proceeds = shares * price;
                    
                    cashFlows.push({
                        date: date ? date.toISOString().split('T')[0] : transaction.transactionDate,
                        category: 'Return',
                        cashFlow: proceeds,
                        description: 'Sale Proceeds'
                    });
                }
            }
        }

        // Add current total portfolio value (positive) - same as XIRR User Investment
        const currentValue = this.calculateCurrentValue();
        
        if (currentValue > 0) {
            cashFlows.push({
                date: new Date().toISOString().split('T')[0],
                category: 'Return',
                cashFlow: currentValue,
                description: 'Total Portfolio Value'
            });
        }

        // Sort by date
        cashFlows.sort((a, b) => new Date(a.date) - new Date(b.date));

        return cashFlows;
    }

    /**
     * Get calculation breakdown for a specific card type
     * @param {string} cardType - Type of card breakdown to retrieve
     * @returns {Array} Breakdown data for the specified card
     */
    getCalculationBreakdown(cardType) {
        if (!this.calculationBreakdowns) {
            return [];
        }
        return this.calculationBreakdowns[cardType] || [];
    }
}

// Global calculator instance
const portfolioCalculator = new PortfolioCalculator();