/**
 * Calculations Tab Module for Equate Portfolio Analysis
 * Handles the accordion interface for displaying detailed calculation breakdowns
 */

class CalculationsTab {
    constructor() {
        this.isInitialized = false;
        this.allExpanded = true; // Default to expand all
        this.accordionContainer = null;
        this.expandAllBtn = null;
        this.exportBtn = null;
        this.translationManager = null;
    }

    /**
     * Initialize the calculations tab
     */
    init() {
        this.cacheElements();
        this.setupEventListeners();
        
        // Initialize translation manager reference
        if (typeof translationManager !== 'undefined') {
            this.translationManager = translationManager;
        }
        
        this.isInitialized = true;
        config.debug('✅ CalculationsTab initialized');
    }

    /**
     * Cache DOM elements
     */
    cacheElements() {
        this.accordionContainer = document.getElementById('accordionContainer');
        this.expandAllBtn = document.getElementById('expandAllBtn');
    }

    /**
     * Set up event listeners
     */
    setupEventListeners() {
        if (this.expandAllBtn) {
            this.expandAllBtn.addEventListener('click', () => this.toggleExpandAll());
        }

    }

    /**
     * Get translation with fallback
     */
    t(key, fallback = null) {
        if (this.translationManager) {
            return this.translationManager.t(key, fallback);
        }
        return fallback || key;
    }

    /**
     * Toggle expand/collapse all accordion items
     */
    toggleExpandAll() {
        const items = this.accordionContainer.querySelectorAll('.accordion-item');
        
        if (this.allExpanded) {
            // Collapse all
            items.forEach(item => item.classList.remove('expanded'));
            this.expandAllBtn.innerHTML = `<span>📖</span><span data-translate="expand_all">${this.t('expand_all', 'Expand All')}</span>`;
            this.allExpanded = false;
        } else {
            // Expand all
            items.forEach(item => item.classList.add('expanded'));
            this.expandAllBtn.innerHTML = `<span>📕</span><span data-translate="collapse_all">${this.t('collapse_all', 'Collapse All')}</span>`;
            this.allExpanded = true;
        }
        
        // No need to refresh charts - this is just a UI state change
        config.debug('📖 Accordion toggle completed (no chart refresh needed)');
    }


    /**
     * Render the calculations tab content
     */
    render(calculations) {
        if (!this.isInitialized) {
            this.init();
        }

        if (!calculations || !portfolioCalculator.calculationBreakdowns) {
            this.renderNoData();
            return;
        }

        this.renderAccordionItems(calculations);
        // Update button text for current language
        this.updateExpandAllButtonState();
        config.debug('📊 Calculations tab rendered');
    }

    /**
     * Render no data state
     */
    renderNoData() {
        if (this.accordionContainer) {
            this.accordionContainer.innerHTML = `
                <div class="no-calculations-data">
                    <p>${this.t('msg_no_calculation_breakdowns', 'No calculation breakdowns available. Please ensure portfolio data is loaded.')}</p>
                </div>
            `;
        }
    }

    /**
     * Render all accordion items
     */
    renderAccordionItems(calculations) {
        if (!this.accordionContainer) return;

        const accordionHTML = this.generateAccordionHTML(calculations);
        this.accordionContainer.innerHTML = accordionHTML;
        
        // Set up individual accordion item event listeners
        this.setupAccordionItemListeners();
        
        // Default to expand all items
        if (this.allExpanded) {
            const items = this.accordionContainer.querySelectorAll('.accordion-item');
            items.forEach(item => item.classList.add('expanded'));
        }
    }

    /**
     * Generate HTML for all accordion items using breakdowns from IndexedDB
     */
    generateAccordionHTML(calculations) {
        let html = '';

        // Get breakdown data that's already calculated by the portfolio calculator
        const breakdowns = portfolioCalculator.calculationBreakdowns;
        
        if (!breakdowns) {
            config.warn('No calculation breakdowns available');
            return `<p>${this.t('msg_no_calculation_breakdowns', 'No calculation breakdowns available. Please ensure portfolio data is loaded.')}</p>`;
        }

        // Define the order and configuration for accordion items following the mockup design
        const accordionConfig = [
            {
                id: 'userInvestment',
                title: this.t('your_investment_title', 'Your Investment'),
                value: this.formatCurrency(calculations.userInvestment),
                cssClass: 'highlight-investment',
                breakdown: breakdowns.userInvestment,
                showCurrentValue: true
            },
            {
                id: 'companyMatch',
                title: this.t('company_match_label', 'Company Match'),
                value: this.formatCurrency(calculations.companyMatch),
                cssClass: '',
                breakdown: breakdowns.companyMatch,
                showCurrentValue: true
            },
            {
                id: 'freeShares',
                title: this.t('free_shares_label', 'Free Shares'),
                value: this.formatCurrency(calculations.freeShares),
                cssClass: '',
                breakdown: breakdowns.freeShares,
                showCurrentValue: true
            },
            {
                id: 'dividendIncome',
                title: this.t('dividend_income', 'Dividend Income'),
                value: this.formatCurrency(calculations.dividendIncome),
                cssClass: '',
                breakdown: breakdowns.dividendIncome,
                showCurrentValue: true
            },
            {
                id: 'totalInvestment',
                title: this.t('total_investment', 'Total Investment'),
                value: this.formatCurrency(calculations.totalInvestment),
                cssClass: 'highlight-total-investment',
                breakdown: breakdowns.totalInvestment,
                showCurrentValue: false // Summary table format
            },
            {
                id: 'currentPortfolio',
                title: this.t('current_portfolio', 'Current Portfolio'),
                value: this.formatCurrency(calculations.currentValue),
                cssClass: '',
                breakdown: breakdowns.currentPortfolio,
                showCurrentValue: false // Special portfolio format
            },
            {
                id: 'totalSold',
                title: this.t('total_sold', 'Total Sold'),
                value: this.formatCurrency(calculations.totalSold),
                cssClass: '',
                breakdown: breakdowns.totalSold,
                showCurrentValue: false // Sold shares format
            },
            {
                id: 'xirrUserInvestment',
                title: this.t('xirr_user_investment', 'Annual Return (XIRR) on Your Investment'),
                value: this.formatPercentage(calculations.xirrUserInvestment / 100),
                cssClass: 'xirr-card',
                breakdown: breakdowns.xirrUserInvestment,
                showCurrentValue: false // XIRR cash flow format
            },
            {
                id: 'xirrTotalInvestment',
                title: this.t('xirr_total_investment', 'Annual Return (XIRR) on Total Investment'),
                value: this.formatPercentage(calculations.xirrTotalInvestment / 100),
                cssClass: 'xirr-card',
                breakdown: breakdowns.xirrTotalInvestment,
                showCurrentValue: false // XIRR cash flow format
            }
        ];

        // Generate HTML for each accordion item that has breakdown data
        accordionConfig.forEach(config => {
            // Show accordion item if value exists and breakdown has data (or for special cases)
            const hasValue = config.value && config.value !== '€ 0.00' && config.value !== '0.0%' && config.value !== '-';
            const hasBreakdown = config.breakdown && config.breakdown.length > 0;
            
            if (hasValue && (hasBreakdown || config.id.includes('xirr'))) {
                html += this.generateAccordionItem(config);
            }
        });

        return html;
    }

    /**
     * Generate HTML for a single accordion item
     */
    generateAccordionItem(config) {
        let tableHTML = '';
        
        // Add XIRR explanation for XIRR cards
        if (config.cssClass === 'xirr-card') {
            tableHTML = this.generateXIRRExplanation();
        }
        
        // Generate appropriate table based on breakdown data structure
        if (config.breakdown && config.breakdown.length > 0) {
            tableHTML += this.generateTableHTML(config.id, config.breakdown);
        } else {
            // Show message for cards with no breakdown data
            tableHTML += `<p>${this.t('msg_no_detailed_calculation_data', 'No detailed calculation data available for this metric.')}</p>`;
        }
        
        return `
            <div class="accordion-item ${config.cssClass} expanded" data-card="${config.id}">
                <div class="metric-card-header">
                    <div class="metric-info">
                        <div class="metric-title">${config.title}</div>
                        <div class="metric-value">${config.value}</div>
                    </div>
                    <div class="expand-indicator">▼</div>
                </div>
                <div class="accordion-content">
                    <div class="calculation-details">
                        ${tableHTML}
                    </div>
                </div>
            </div>
        `;
    }


    /**
     * Get readable category name from entry
     */
    getCategoryName(entry) {
        if (entry.contributionType === 'Purchase' && entry.plan === 'Employee Share Purchase Plan') return this.t('category_employee_purchase', 'Employee Purchase');
        if (entry.contributionType === 'Company match' && entry.plan === 'Employee Share Purchase Plan') return this.t('category_company_match', 'Company Match');
        if (entry.contributionType === 'Award' && entry.plan === 'Free Share') return this.t('free_shares_label', 'Free Shares');
        if (entry.plan === 'Allianz Dividend Reinvestment') return this.t('category_dividend_reinvestment', 'Dividend Reinvestment');
        if (entry.type === 'Sell') return this.t('category_share_sale', 'Share Sale');
        if (entry.type === 'Transfer') return this.t('category_transfer_out', 'Transfer Out');
        return entry.contributionType || entry.plan || this.t('label_unknown', 'Unknown');
    }

    /**
     * Generate XIRR explanation HTML
     */
    generateXIRRExplanation() {
        return `
            <div class="xirr-explanation">
                <strong>${this.t('xirr_explanation', 'XIRR (Extended Internal Rate of Return) calculates the annualized return rate considering the timing of your cash flows. It shows what annual interest rate would give the same return as your actual investment pattern.')}</strong>
            </div>
        `;
    }

    /**
     * Generate table HTML based on breakdown data structure
     */
    generateTableHTML(cardType, breakdown) {
        if (!breakdown || breakdown.length === 0) {
            return `<p>${this.t('msg_no_detailed_calculation_data', 'No calculation data available for this metric.')}</p>`;
        }

        const firstItem = breakdown[0];
        
        if (firstItem.hasOwnProperty('currentValue')) {
            return this.generateInvestmentTable(breakdown);
        } else if (firstItem.hasOwnProperty('percentage')) {
            return this.generateSummaryTable(breakdown);
        } else if (firstItem.hasOwnProperty('cashFlow')) {
            return this.generateCashFlowTable(breakdown);
        } else if (firstItem.hasOwnProperty('proceeds')) {
            return this.generateSoldTable(breakdown);
        } else if (firstItem.hasOwnProperty('profitLoss')) {
            return this.generatePortfolioTable(breakdown);
        }
        
        return `<p>${this.t('msg_unknown_breakdown_format', 'Unknown breakdown format.')}</p>`;
    }

    /**
     * Generate investment table (refined 5-column format with Current Value)
     */
    generateInvestmentTable(breakdown) {
        let html = `
            <table class="calculation-table">
                <thead>
                    <tr>
                        <th>${this.t('table_date', 'Date')}</th>
                        <th>${this.t('table_category', 'Category')}</th>
                        <th>${this.t('table_shares', 'Shares')}</th>
                        <th>${this.t('table_investment', 'Investment')}</th>
                        <th>${this.t('table_current_value', 'Current Value')}</th>
                    </tr>
                </thead>
                <tbody>
        `;

        breakdown.forEach(item => {
            const rowClass = item.isTotal ? 'total-row' : '';
            const categoryBadge = item.category ? `<span class="category-badge ${this.getCategoryBadgeClass(item.category)}">${item.category}</span>` : '';
            
            if (item.isTotal) {
                html += `
                    <tr class="${rowClass}">
                        <td colspan="2"><strong>${this.t('label_total', 'Total')}</strong></td>
                        <td><strong>${item.shares ? this.formatNumber(item.shares) : ''}</strong></td>
                        <td><strong>${item.investment ? this.formatCurrency(item.investment) : ''}</strong></td>
                        <td><strong>${item.currentValue ? this.formatCurrency(item.currentValue) : ''}</strong></td>
                    </tr>
                `;
            } else {
                // Handle profit/loss coloring for Current Portfolio
                let currentValueDisplay = '';
                if (item.currentValue) {
                    if (item.isProfitLoss) {
                        const color = item.currentValue < 0 ? '#e74c3c' : '#27ae60';
                        const sign = item.currentValue > 0 ? '+' : '';
                        currentValueDisplay = `<span style="color: ${color}; font-weight: bold;">${sign}${this.formatCurrency(item.currentValue)}</span>`;
                    } else {
                        currentValueDisplay = this.formatCurrency(item.currentValue);
                    }
                }
                
                html += `
                    <tr class="${rowClass}">
                        <td>${item.date || ''}</td>
                        <td>${categoryBadge}</td>
                        <td>${item.shares ? this.formatNumber(item.shares) : ''}</td>
                        <td>${item.investment ? this.formatCurrency(item.investment) : ''}</td>
                        <td>${currentValueDisplay}</td>
                    </tr>
                `;
            }
        });

        html += '</tbody></table>';
        return html;
    }

    /**
     * Generate summary table (for Total Investment)
     */
    generateSummaryTable(breakdown) {
        let html = `
            <table class="calculation-table">
                <thead>
                    <tr>
                        <th>${this.t('table_category', 'Category')}</th>
                        <th>${this.t('table_transactions', 'Transactions')}</th>
                        <th>${this.t('table_total_shares', 'Total Shares')}</th>
                        <th>${this.t('table_investment_amount', 'Investment Amount')}</th>
                        <th>${this.t('table_percentage', 'Percentage')}</th>
                    </tr>
                </thead>
                <tbody>
        `;

        breakdown.forEach(item => {
            const rowClass = item.isTotal ? 'total-row' : '';
            const categoryBadge = item.category ? `<span class="category-badge ${this.getCategoryBadgeClass(item.category)}">${item.category}</span>` : '';
            
            if (item.isTotal) {
                html += `
                    <tr class="${rowClass}">
                        <td><strong>${this.t('label_total', 'Total')}</strong></td>
                        <td><strong>${item.transactions || ''}</strong></td>
                        <td><strong>${item.shares ? this.formatNumber(item.shares) : ''}</strong></td>
                        <td><strong>${item.investment ? this.formatCurrency(item.investment) : ''}</strong></td>
                        <td><strong>${item.percentage ? item.percentage.toFixed(1) + '%' : ''}</strong></td>
                    </tr>
                `;
            } else {
                html += `
                    <tr class="${rowClass}">
                        <td>${categoryBadge}</td>
                        <td>${item.transactions || ''}</td>
                        <td>${item.shares ? this.formatNumber(item.shares) : ''}</td>
                        <td>${item.investment ? this.formatCurrency(item.investment) : ''}</td>
                        <td>${item.percentage ? item.percentage.toFixed(1) + '%' : ''}</td>
                    </tr>
                `;
            }
        });

        html += '</tbody></table>';
        return html;
    }

    /**
     * Generate cash flow table (for XIRR)
     */
    generateCashFlowTable(breakdown) {
        let html = `
            <table class="calculation-table">
                <thead>
                    <tr>
                        <th>${this.t('table_date', 'Date')}</th>
                        <th>${this.t('table_category', 'Category')}</th>
                        <th>${this.t('table_cash_flow', 'Cash Flow')}</th>
                        <th>${this.t('table_description', 'Description')}</th>
                    </tr>
                </thead>
                <tbody>
        `;

        breakdown.forEach(item => {
            if (item.isTotal) {
                html += `
                    <tr class="total-row">
                        <td colspan="2"><strong>${this.t('msg_xirr_calculation_result', 'XIRR Calculation Result')}</strong></td>
                        <td colspan="2"><strong>${item.xirrResult || this.t('msg_xirr_calculation_result', 'Calculation Result')}</strong></td>
                    </tr>
                `;
            } else {
                const cashFlowColor = item.cashFlow < 0 ? 'color: #e74c3c;' : 'color: #27ae60;';
                const categoryBadge = `<span class="category-badge ${this.getCategoryBadgeClass(item.category || 'Cash Flow')}">${item.category || this.t('table_cash_flow', 'Cash Flow')}</span>`;
                
                html += `
                    <tr>
                        <td>${item.date || ''}</td>
                        <td>${categoryBadge}</td>
                        <td style="${cashFlowColor}">${item.cashFlow < 0 ? '-' : '+'}${this.formatCurrency(Math.abs(item.cashFlow || 0))}</td>
                        <td>${item.description || ''}</td>
                    </tr>
                `;
            }
        });

        html += '</tbody></table>';
        return html;
    }

    /**
     * Generate sold table
     */
    generateSoldTable(breakdown) {
        let html = `
            <table class="calculation-table">
                <thead>
                    <tr>
                        <th>${this.t('table_date', 'Date')}</th>
                        <th>${this.t('table_category', 'Category')}</th>
                        <th>${this.t('table_shares_sold', 'Shares Sold')}</th>
                        <th>${this.t('table_sale_price', 'Sale Price')}</th>
                        <th>${this.t('table_sale_proceeds', 'Sale Proceeds')}</th>
                    </tr>
                </thead>
                <tbody>
        `;

        breakdown.forEach(item => {
            const rowClass = item.isTotal ? 'total-row' : '';
            const categoryBadge = item.category ? `<span class="category-badge ${this.getCategoryBadgeClass(item.category)}">${item.category}</span>` : '';
            
            if (item.isTotal) {
                html += `
                    <tr class="${rowClass}">
                        <td colspan="2"><strong>${this.t('label_total', 'Total')}</strong></td>
                        <td><strong>${item.shares ? this.formatNumber(item.shares) : ''}</strong></td>
                        <td><strong>${item.price ? this.formatCurrency(item.price) + ' ' + this.t('label_average', 'avg') : ''}</strong></td>
                        <td><strong>${item.proceeds ? this.formatCurrency(item.proceeds) : ''}</strong></td>
                    </tr>
                `;
            } else {
                html += `
                    <tr class="${rowClass}">
                        <td>${item.date || ''}</td>
                        <td>${categoryBadge}</td>
                        <td>${item.shares ? this.formatNumber(item.shares) : ''}</td>
                        <td>${item.price ? this.formatCurrency(item.price) : ''}</td>
                        <td>${item.proceeds ? this.formatCurrency(item.proceeds) : ''}</td>
                    </tr>
                `;
            }
        });

        html += '</tbody></table>';
        return html;
    }

    /**
     * Generate portfolio table (for Current Portfolio)
     */
    generatePortfolioTable(breakdown) {
        let html = `
            <table class="calculation-table">
                <thead>
                    <tr>
                        <th>${this.t('table_category', 'Category')}</th>
                        <th>${this.t('table_shares', 'Shares')}</th>
                        <th>${this.t('table_investment', 'Investment')}</th>
                        <th>${this.t('table_value_at_download', 'Value at Download')}</th>
                        <th>${this.t('table_current_value', 'Current Value')}</th>
                        <th>${this.t('table_profit_loss', 'Profit/Loss')}</th>
                    </tr>
                </thead>
                <tbody>
        `;

        breakdown.forEach(item => {
            const rowClass = item.isTotal ? 'total-row' : (item.isSubtotal ? 'subtotal-row' : '');
            const profitLossColor = item.profitLoss && item.profitLoss < 0 ? 'color: #e74c3c;' : 'color: #27ae60;';
            const categoryBadge = `<span class="category-badge profit-loss">${item.category || this.t('current_portfolio', 'Portfolio')}</span>`;
            
            if (item.isTotal) {
                html += `
                    <tr class="${rowClass}">
                        <td colspan="5"><strong>${this.t('msg_total_current_portfolio_value', 'Total Current Portfolio Value')}</strong></td>
                        <td><strong>${item.currentValue ? this.formatCurrency(item.currentValue) : ''}</strong></td>
                    </tr>
                `;
            } else {
                html += `
                    <tr class="${rowClass}">
                        <td>${categoryBadge}</td>
                        <td>${item.shares ? this.formatNumber(item.shares) : '-'}</td>
                        <td>${item.investment ? this.formatCurrency(item.investment) : '-'}</td>
                        <td>${item.valueAtDownload ? this.formatCurrency(item.valueAtDownload) : '-'}</td>
                        <td>${item.currentValue ? this.formatCurrency(item.currentValue) : '-'}</td>
                        <td style="${profitLossColor}">${item.profitLoss ? (item.profitLoss > 0 ? '+' : '') + this.formatCurrency(item.profitLoss) : '-'}</td>
                    </tr>
                `;
            }
        });

        html += '</tbody></table>';
        return html;
    }

    /**
     * Set up event listeners for individual accordion items
     */
    setupAccordionItemListeners() {
        const headers = this.accordionContainer.querySelectorAll('.metric-card-header');
        
        headers.forEach(header => {
            header.addEventListener('click', () => {
                const item = header.closest('.accordion-item');
                const isExpanded = item.classList.contains('expanded');
                
                // Toggle current accordion
                if (isExpanded) {
                    item.classList.remove('expanded');
                } else {
                    item.classList.add('expanded');
                }
                
                // Update global state
                this.updateExpandAllButtonState();
            });
        });
    }

    /**
     * Update expand all button state based on current accordion states
     */
    updateExpandAllButtonState() {
        const allItems = this.accordionContainer.querySelectorAll('.accordion-item');
        const expandedItems = this.accordionContainer.querySelectorAll('.accordion-item.expanded');
        
        if (expandedItems.length === allItems.length) {
            this.allExpanded = true;
            this.expandAllBtn.innerHTML = `<span>📕</span><span data-translate="collapse_all">${this.t('collapse_all', 'Collapse All')}</span>`;
        } else if (expandedItems.length === 0) {
            this.allExpanded = false;
            this.expandAllBtn.innerHTML = `<span>📖</span><span data-translate="expand_all">${this.t('expand_all', 'Expand All')}</span>`;
        } else {
            // Mixed state - show collapse all since some are expanded
            this.expandAllBtn.innerHTML = `<span>📕</span><span data-translate="collapse_all">${this.t('collapse_all', 'Collapse All')}</span>`;
        }
        
        // No need to refresh charts - button state change only
        config.debug('📖 Expand/collapse button state updated (no chart refresh needed)');
    }

    /**
     * Get CSS class for category badge
     */
    getCategoryBadgeClass(category) {
        if (!category) return '';
        
        const lowerCategory = category.toLowerCase();
        let cssClass = '';
        
        if (lowerCategory.includes('purchase')) cssClass = 'user-investment';
        else if (lowerCategory.includes('match')) cssClass = 'company-investment';
        else if (lowerCategory.includes('dividend')) cssClass = 'dividend';
        else if (lowerCategory.includes('free')) cssClass = 'free-share';
        else if (lowerCategory.includes('return')) cssClass = 'return';
        else if (lowerCategory.includes('sale')) cssClass = 'return';
        else if (lowerCategory.includes('transfer')) cssClass = 'return';
        else cssClass = '';
        
        // Removed verbose CSS class logging
        return cssClass;
    }

    /**
     * Format currency value using app's formatting method
     */
    formatCurrency(value) {
        if (value === null || value === undefined || isNaN(value)) return '-';
        const currency = portfolioCalculator.currency || '';
        return currencyConverter.formatCurrency(value, currency, 2);
    }

    /**
     * Format percentage value
     */
    formatPercentage(value) {
        if (value === null || value === undefined || isNaN(value)) return '-';
        return currencyConverter.formatPercentage(value * 100, 1);
    }

    /**
     * Format number value using app's formatting method
     */
    formatNumber(value) {
        if (value === null || value === undefined || isNaN(value)) return '-';
        
        // Use the app's formatNumber method if available, otherwise fallback
        if (typeof app !== 'undefined' && app.formatNumber) {
            return app.formatNumber(value, 2);
        } else {
            // Fallback formatting
            return value.toLocaleString('en-US', {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });
        }
    }

    /**
     * Show or hide calculations content based on data availability
     */
    toggleCalculationsDisplay(hasData) {
        const noDataMessage = document.getElementById('calculationsNoDataMessage');
        const calculationsContent = document.getElementById('calculationsContent');
        
        if (hasData) {
            if (noDataMessage) noDataMessage.style.display = 'none';
            if (calculationsContent) calculationsContent.style.display = 'block';
        } else {
            if (noDataMessage) noDataMessage.style.display = 'block';
            if (calculationsContent) calculationsContent.style.display = 'none';
        }
    }
}

// Global calculations tab instance
const calculationsTab = new CalculationsTab();