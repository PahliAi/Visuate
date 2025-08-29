/**
 * Historical Price Generator for Non-EUR Portfolios
 * Converts EUR historical prices to local currency for accurate timeline charts
 * Uses currency ratios from portfolio/transaction reference points with interpolation
 */

class HistoricalPriceGenerator {
    constructor() {
        this.portfolioCurrency = null;
        this.eurHistoricalPrices = null;
        this.enhancedReferencePoints = [];
        this.generatedPrices = [];
    }

    /**
     * Generate local currency historical prices from EUR data
     * @param {Array} eurHistoricalPrices - EUR historical prices from hist.xlsx
     * @param {Object} portfolioData - Portfolio data with asOfDate, marketPrice, and entries
     * @param {Array} transactionEntries - Transaction entries (optional)
     * @param {string} portfolioCurrency - Target currency (e.g., 'ARS', 'USD')
     * @param {Array} historicalCurrencyData - Historical currency rates from hist.xlsx Currency sheet (optional)
     * @param {Array} allCurrencyPrices - Pre-computed multi-currency price data from company files (optional, for performance)
     * @param {Array} enhancedReferencePoints - Pre-built enhanced reference points from calculator (V2 optimization)
     * @returns {Promise<Array>} Generated historical prices in local currency
     * @throws {Error} If insufficient reference data or conversion fails
     */
    async generateLocalCurrencyPrices(eurHistoricalPrices, portfolioData, transactionEntries, portfolioCurrency, historicalCurrencyData = [], allCurrencyPrices = null, enhancedReferencePoints = null) {
        config.debug('🏭 HistoricalPriceGenerator starting generation:', {
            portfolioCurrency,
            eurPricesCount: eurHistoricalPrices?.length || 0,
            portfolioEntries: portfolioData?.entries?.length || 0,
            transactionEntries: transactionEntries?.length || 0,
            historicalCurrencyDataCount: historicalCurrencyData?.length || 0
        });

        this.portfolioCurrency = portfolioCurrency;
        
        // 🎯 PORTFOLIO-AWARE FILTERING: Filter EUR historical prices to portfolio timeline
        if (enhancedReferencePoints && enhancedReferencePoints.length > 0) {
            const firstPurchaseDate = enhancedReferencePoints[0].date; // Already sorted by date
            const originalCount = eurHistoricalPrices?.length || 0;
            
            this.eurHistoricalPrices = eurHistoricalPrices?.filter(price => price.date >= firstPurchaseDate) || [];
            
            const filteredCount = this.eurHistoricalPrices.length;
            const reductionPercent = originalCount > 0 ? ((originalCount - filteredCount) / originalCount * 100).toFixed(1) : 0;
            
            config.debug(`🎯 Portfolio-aware filtering: ${originalCount} → ${filteredCount} historical prices (${reductionPercent}% reduction, starting from ${firstPurchaseDate})`);
        } else {
            this.eurHistoricalPrices = eurHistoricalPrices;
            config.debug('📅 No enhanced reference points available, using all historical price data');
        }
        
        this.allCurrencyPrices = allCurrencyPrices; // Pre-computed multi-currency data for performance
        this.enhancedReferencePoints = [];
        this.generatedPrices = [];
        
        // 🚀 REVOLUTIONARY: Check for multi-currency data bypass (Phase 3)
        const multiCurrencyData = await equateDB.getMultiCurrencyPrices();
        
        if (multiCurrencyData.length > 0 && portfolioCurrency in multiCurrencyData[0]) {
            config.debug(`🚀 INSTANT BYPASS: Using pre-computed ${portfolioCurrency} prices from multi-currency database!`);
            config.debug(`💾 Found ${multiCurrencyData.length} pre-computed entries with ${portfolioCurrency} currency`);
            
            // Direct extraction - NO CONVERSION LOOP NEEDED!
            const directPrices = multiCurrencyData
                .map(entry => ({
                    date: entry.date,
                    price: entry[portfolioCurrency]
                }))
                .filter(entry => entry.price != null && entry.price > 0)
                .sort((a, b) => new Date(a.date) - new Date(b.date));
            
            config.debug(`✨ INSTANT RESULT: ${directPrices.length} ${portfolioCurrency} prices extracted in ~1ms (no conversion needed)`);
            return directPrices;
        }
        
        // V2 BINARY ARCHITECTURE: Switch to Synthetic Mode (flat-line timeline from reference points)
        config.debug(`🔄 SYNTHETIC MODE: ${portfolioCurrency} not found in multi-currency data, generating flat-line timeline from reference points`);

        // V2 SYNTHETIC MODE: Generate flat-line timeline from enhanced reference points
        if (enhancedReferencePoints && enhancedReferencePoints.length > 0) {
            config.debug('🚀 V2 SYNTHETIC: Using shared enhanced reference points from calculator');
            this.enhancedReferencePoints = enhancedReferencePoints
                .filter(point => point.currentCurrency === portfolioCurrency || point.originalCurrency === portfolioCurrency)
                .sort((a, b) => new Date(a.date) - new Date(b.date));
        } else {
            config.debug('🔄 V2 SYNTHETIC: Building enhanced reference points for flat-line timeline');
            this.enhancedReferencePoints = await referencePointsBuilder.buildReferencePoints(
                portfolioData, { entries: transactionEntries }
            );
        }
        
        if (this.enhancedReferencePoints.length === 0) {
            throw new Error(`No reference points available for synthetic timeline generation in ${portfolioCurrency}`);
        }
        
        // Generate synthetic flat-line timeline from reference points
        const syntheticPrices = this.generateSyntheticFlatLineTimeline();
        
        config.debug('✅ V2 SYNTHETIC timeline generated:', {
            currency: portfolioCurrency,
            referencePoints: this.enhancedReferencePoints.length,
            timelinePoints: syntheticPrices.length,
            dateRange: syntheticPrices.length > 0 ? {
                from: syntheticPrices[0].date,
                to: syntheticPrices[syntheticPrices.length - 1].date
            } : null
        });
        
        return syntheticPrices;
    }
    
    /**
     * Generate synthetic flat-line timeline from enhanced reference points
     * V2 BINARY ARCHITECTURE: Simple flat-line interpolation between reference points
     */
    generateSyntheticFlatLineTimeline() {
        if (this.enhancedReferencePoints.length === 0) return [];
        
        const timelinePoints = this.enhancedReferencePoints.map(point => ({
            date: point.date,
            price: point.currentPrice || point.prices?.[point.originalCurrency] || 0
        }));
        
        const timeline = [];
        const startDate = new Date(timelinePoints[0].date);
        const endDate = new Date(timelinePoints[timelinePoints.length - 1].date);
        
        let currentDate = new Date(startDate);
        let currentPriceIndex = 0;
        
        // Generate daily flat-line timeline
        while (currentDate <= endDate) {
            const dateStr = currentDate.toISOString().split('T')[0];
            
            // Find appropriate price (flat-line approach)
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


}

// Global instance
const historicalPriceGenerator = new HistoricalPriceGenerator();