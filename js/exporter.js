/**
 * Export functionality for Equate Portfolio Analysis
 * Handles PDF export of portfolio data and analysis
 */

class PortfolioExporter {
    constructor() {
        this.calculations = null;
        this.portfolioData = null;
    }

    /**
     * Set data for export
     */
    setData(calculations, portfolioData) {
        this.calculations = calculations;
        this.portfolioData = portfolioData;
    }

    /**
     * Get currency symbol safely - no fallbacks to prevent wrong currency display
     */


    /**
     * Export portfolio analysis as PDF with custom section configuration
     */
    async exportCustomPDF(exportConfig = null) {
        if (!this.calculations || !this.portfolioData) {
            throw new Error('No data available for export');
        }

        try {
            config.debug('📋 Generating custom PDF-optimized report...');
            
            // Capture all chart images if available
            const chartImages = {};
            
            // Only capture charts for enabled sections
            const enabledSectionIds = exportConfig ? exportConfig.sections.map(s => s.id) : ['timelineChart', 'performanceBarChart', 'investmentPieChart'];
            
            // 1. Portfolio Timeline Chart
            if (enabledSectionIds.includes('timelineChart')) {
                const portfolioChartElement = document.getElementById('portfolioChart');
                if (portfolioChartElement && window.Plotly) {
                    config.debug('📊 Capturing portfolio timeline chart...');
                    try {
                        await new Promise(resolve => setTimeout(resolve, 100));
                        chartImages.portfolioChart = await window.Plotly.toImage(portfolioChartElement, {
                            format: 'png',
                            width: 800,
                            height: 400,
                            scale: 2
                        });
                        config.debug('✅ Portfolio timeline chart captured');
                    } catch (error) {
                        console.warn('⚠️ Could not capture portfolio timeline chart:', error.message);
                    }
                }
            }
            
            // 2. Performance Bar Chart
            if (enabledSectionIds.includes('performanceBarChart')) {
                const performanceChartElement = document.getElementById('performanceBarChart');
                if (performanceChartElement && window.Plotly) {
                    config.debug('📊 Capturing performance bar chart...');
                    try {
                        await new Promise(resolve => setTimeout(resolve, 100));
                        chartImages.performanceChart = await window.Plotly.toImage(performanceChartElement, {
                            format: 'png',
                            width: 800,
                            height: 400,
                            scale: 2
                        });
                        config.debug('✅ Performance bar chart captured');
                    } catch (error) {
                        console.warn('⚠️ Could not capture performance bar chart:', error.message);
                    }
                }
            }
            
            // 3. Investment Pie Chart
            if (enabledSectionIds.includes('investmentPieChart')) {
                const pieChartElement = document.getElementById('investmentPieChart');
                if (pieChartElement && window.Plotly) {
                    config.debug('📊 Capturing investment pie chart...');
                    try {
                        await new Promise(resolve => setTimeout(resolve, 100));
                        chartImages.pieChart = await window.Plotly.toImage(pieChartElement, {
                            format: 'png',
                            width: 800,
                            height: 400,
                            scale: 2
                        });
                        config.debug('✅ Investment pie chart captured');
                    } catch (error) {
                        console.warn('⚠️ Could not capture investment pie chart:', error.message);
                    }
                }
            }
            
            config.debug('📊 Chart capture summary:', {
                portfolioChart: !!chartImages.portfolioChart,
                performanceChart: !!chartImages.performanceChart,
                pieChart: !!chartImages.pieChart
            });
            
            // Generate custom filename for PDF
            const now = new Date();
            const dateStr = now.toISOString().split('T')[0].replace(/-/g, '');
            const timeStr = now.toTimeString().split(' ')[0].slice(0, 5).replace(/:/g, '.');
            const customTitle = `${dateStr}_${timeStr}_Visuate_Portfolio_Analysis_Report`;
            
            // Generate print-optimized HTML content with custom configuration
            const htmlContent = this.generatePrintOptimizedHTML(chartImages, customTitle, exportConfig);
            
            // Create a new window with the content
            const printWindow = window.open('', '_blank');
            printWindow.document.write(htmlContent);
            printWindow.document.close();
            
            // Wait for content to load, then trigger print dialog
            printWindow.onload = () => {
                setTimeout(() => {
                    printWindow.focus();
                    printWindow.print();
                    // Close the window after printing (user can cancel)
                    printWindow.addEventListener('afterprint', () => {
                        printWindow.close();
                    });
                }, 500);
            };
            
        } catch (error) {
            config.error('Custom PDF export error:', error);
            // Fallback to HTML export
            this.exportHTML();
        }
    }

    /**
     * Export portfolio analysis as PDF (default behavior)
     */
    async exportPDF() {
        // Use the enhanced exportCustomPDF with no config (shows all sections)
        return this.exportCustomPDF(null);
    }

    /**
     * Export as HTML (fallback for PDF)
     */
    exportHTML() {
        const htmlContent = this.generateHTMLReport();
        const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8;' });
        this.downloadFile(blob, this.generateFileName('html'));
    }


    /**
     * Generate HTML report content
     */
    generateHTMLReport() {
        return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Portfolio Analysis Report</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
        .header { text-align: center; border-bottom: 2px solid #333; padding-bottom: 20px; margin-bottom: 30px; }
        .metrics-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-bottom: 30px; }
        .metric-card { border: 1px solid #ddd; border-radius: 8px; padding: 15px; text-align: center; }
        .metric-card.highlight { background-color: #f8f9fa; border-color: #007bff; }
        .metric-label { color: #666; font-size: 12px; text-transform: uppercase; margin-bottom: 5px; }
        .metric-value { font-size: 20px; font-weight: bold; color: #333; }
        .positive { color: #27ae60; }
        .negative { color: #e74c3c; }
        .price-info { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; }
        .chart-notice { background: #e7f3ff; border: 1px solid #b3d9ff; border-radius: 8px; padding: 20px; text-align: center; color: #0066cc; margin: 20px 0; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .section { margin-bottom: 30px; }
        .section h2 { color: #333; border-bottom: 1px solid #ddd; padding-bottom: 10px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Portfolio Analysis Report</h1>
        <p>Generated: ${new Date().toLocaleString()}</p>
        <p>User ID: ${this.portfolioData.userId}</p>
    </div>

    <div class="section">
        <h2>Price Information</h2>
        <div class="price-info">
            <p><strong>Current Price:</strong> ${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.currentPrice.toFixed(2)}</p>
            <p><strong>Price Source:</strong> ${this.calculations.priceSource}</p>
            <p><strong>Price Date:</strong> ${this.calculations.priceDate}</p>
        </div>
    </div>

    <div class="section">
        <h2>Portfolio Summary</h2>
        <div class="metrics-grid">
            <!-- First Row: Investment Details -->
            <div class="metric-card highlight">
                <div class="metric-label">Your Investment</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.userInvestment.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Company Match</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.companyInvestment.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Dividend Income</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.dividendIncome.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Investment</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.totalInvestment.toFixed(2)}</div>
            </div>
            
            <!-- Second Row: Portfolio Performance -->
            <div class="metric-card">
                <div class="metric-label">Current Portfolio</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.currentValue.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Sold</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.totalSold.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Value</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.totalValue.toFixed(2)}</div>
            </div>
            <div class="metric-card highlight">
                <div class="metric-label">Total Return</div>
                <div class="metric-value ${this.calculations.totalReturn >= 0 ? 'positive' : 'negative'}">
                    ${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.totalReturn.toFixed(2)}
                </div>
            </div>
            
            <!-- Third Row: Additional Metrics -->
            <div class="metric-card">
                <div class="metric-label">Return Percentage</div>
                <div class="metric-value ${this.calculations.returnPercentage >= 0 ? 'positive' : 'negative'}">
                    ${this.calculations.returnPercentage.toFixed(2)}%
                </div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Annual Growth (CAGR)</div>
                <div class="metric-value">${this.calculations.annualGrowth.toFixed(2)}%</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Available Shares</div>
                <div class="metric-value">${this.calculations.availableShares}</div>
            </div>
        </div>
    </div>

    <div class="section">
        <h2>Share Information</h2>
        <p><strong>Total Shares:</strong> ${this.calculations.totalShares}</p>
        
        ${this.generateBlockedSharesTable()}
    </div>

    <div class="section">
        <h2>Portfolio Performance Timeline</h2>
        <div class="chart-notice">
            📊 <strong>Interactive Chart Available in App</strong><br>
            View the complete portfolio timeline chart with transaction markers and interactive features in the web application.
        </div>
    </div>

    <div class="section">
        <h2>Portfolio Details</h2>
        ${this.generatePortfolioTable()}
    </div>

    <div style="margin-top: 40px; text-align: center; color: #666; font-size: 12px;">
        <p>Generated by Equate Portfolio Analysis Tool</p>
        <p>Data processed locally - no information was transmitted to external servers</p>
        <p><em>For interactive charts and full features, visit the web application</em></p>
    </div>
</body>
</html>`;
    }

    /**
     * Generate blocked shares table HTML
     */
    generateBlockedSharesTable() {
        if (!this.calculations.blockedShares.byYear || Object.keys(this.calculations.blockedShares.byYear).length === 0) {
            return '<p>No blocked shares found.</p>';
        }

        let html = '<h3>Blocked Shares by Unlock Year</h3>';
        html += '<table><thead><tr><th>Unlock Year</th><th>Shares</th></tr></thead><tbody>';
        
        Object.entries(this.calculations.blockedShares.byYear)
            .sort(([a], [b]) => parseInt(a) - parseInt(b))
            .forEach(([year, shares]) => {
                html += `<tr><td>${year}</td><td>${shares}</td></tr>`;
            });
        
        html += '</tbody></table>';
        return html;
    }

    /**
     * Generate portfolio table HTML
     */
    generatePortfolioTable() {
        let html = '<table><thead><tr>';
        html += '<th>Date</th><th>Plan</th><th>Type</th><th>Cost Basis</th>';
        html += '<th>Outstanding</th><th>Available</th><th>Purchase Amount</th>';
        html += '</tr></thead><tbody>';

        this.portfolioData.entries.forEach(entry => {
            html += '<tr>';
            html += `<td>${entry.allocationDate || '-'}</td>`;
            html += `<td>${entry.plan}</td>`;
            html += `<td>${entry.contributionType}</td>`;
            const currencySymbol = currencyConverter.getCurrencySymbol(this.calculations.currency);
            html += `<td>${currencySymbol} ${entry.costBasis.toFixed(2)}</td>`;
            html += `<td>${entry.outstandingQuantity}</td>`;
            html += `<td>${entry.availableQuantity}</td>`;
            html += `<td>${currencySymbol} ${(entry.outstandingQuantity * entry.costBasis).toFixed(2)}</td>`;
            html += '</tr>';
        });

        html += '</tbody></table>';
        return html;
    }

    /**
     * Get user's section order from localStorage or use default
     */
    getUserSectionOrder() {
        const savedOrder = localStorage.getItem('sectionOrder');
        return savedOrder ? JSON.parse(savedOrder) : ['cards', 'timeline', 'bar', 'pie', 'calculations'];
    }

    /**
     * Map new section IDs to legacy section IDs for backward compatibility
     */
    mapSectionIdToLegacy(sectionId) {
        const mapping = {
            'portfolioCards': 'cards',
            'timelineChart': 'timeline', 
            'performanceBarChart': 'bar',
            'investmentPieChart': 'pie'
        };
        
        // All calculation sections map to 'calculations' for the main section generation
        const calculationSections = [
            'yourInvestment', 'companyMatch', 'freeShares', 'dividendIncome',
            'totalInvestment', 'currentPortfolio', 'totalSold', 
            'xirrYourInvestment', 'xirrTotalInvestment'
        ];
        
        if (calculationSections.includes(sectionId)) {
            return 'calculations';
        }
        
        return mapping[sectionId] || sectionId;
    }

    /**
     * Generate section HTML content
     */
    generateSectionHTML(sectionType, chartImages, exportConfig = null) {
        const sectionMapping = {
            'cards': () => this.generateCardsSection(),
            'timeline': () => this.generateTimelineSection(chartImages.portfolioChart),
            'bar': () => this.generateBarSection(chartImages.performanceChart),
            'pie': () => this.generatePieSection(chartImages.pieChart),
            'calculations': () => this.generateCalculationsSection(exportConfig)
        };

        const generator = sectionMapping[sectionType];
        return generator ? generator() : '';
    }

    /**
     * Generate cards section HTML
     */
    generateCardsSection() {
        return `
    <div class="section">
        <h2>Portfolio Summary</h2>
        <div class="metrics-grid">
            <!-- First Row: Investment Details -->
            <div class="metric-card highlight">
                <div class="metric-label">Your Investment</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.userInvestment.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Company Match</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${(this.calculations.companyMatch || 0).toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Free Shares</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${(this.calculations.freeShares || 0).toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Dividend Income</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.dividendIncome.toFixed(2)}</div>
            </div>
            
            <!-- Second Row: Portfolio Performance -->
            <div class="metric-card highlight" style="background-color: #fff8e6; border-color: #f39c12;">
                <div class="metric-label">Total Investment</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.totalInvestment.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Current Portfolio</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.currentValue.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Sold</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.totalSold.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Value</div>
                <div class="metric-value">${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.totalValue.toFixed(2)}</div>
            </div>
            
            <!-- Third Row: Returns -->
            <div class="metric-card highlight">
                <div class="metric-label">Return on Your Investment</div>
                <div class="metric-value ${this.calculations.totalReturn >= 0 ? 'positive' : 'negative'}">
                    ${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.totalReturn.toFixed(2)}
                </div>
            </div>
            <div class="metric-card highlight" style="background-color: #fff8e6; border-color: #f39c12;">
                <div class="metric-label">Return on Total Investment</div>
                <div class="metric-value ${(this.calculations.returnOnTotalInvestment || 0) >= 0 ? 'positive' : 'negative'}">
                    ${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${(this.calculations.returnOnTotalInvestment || 0).toFixed(2)}
                </div>
            </div>
            <div class="metric-card highlight">
                <div class="metric-label">Return % on Your Investment</div>
                <div class="metric-value ${this.calculations.returnPercentage >= 0 ? 'positive' : 'negative'}">
                    ${this.calculations.returnPercentage.toFixed(2)}%
                </div>
            </div>
            <div class="metric-card highlight" style="background-color: #fff8e6; border-color: #f39c12;">
                <div class="metric-label">Return % on Total Investment</div>
                <div class="metric-value ${(this.calculations.returnPercentageOnTotalInvestment || 0) >= 0 ? 'positive' : 'negative'}">
                    ${(this.calculations.returnPercentageOnTotalInvestment || 0).toFixed(2)}%
                </div>
            </div>
            
            <!-- Fourth Row: Additional Metrics -->
            <div class="metric-card">
                <div class="metric-label">Annual Growth (CAGR)</div>
                <div class="metric-value">${this.calculations.annualGrowth.toFixed(2)}%</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Available Shares</div>
                <div class="metric-value">${this.calculations.availableShares}</div>
            </div>
        </div>
    </div>`;
    }

    /**
     * Generate timeline section HTML
     */
    generateTimelineSection(chartImage) {
        return `
    <div class="section">
        <h2>Portfolio Performance Timeline</h2>
        ${chartImage ? `
        <div class="chart-image">
            <img src="${chartImage}" alt="Portfolio Performance Timeline Chart" />
            <p style="font-size: 10px; color: #666; margin-top: 10px;">
                <em>Portfolio value progression over time with transaction markers</em>
            </p>
        </div>
        ` : `
        <div class="chart-notice">
            📊 <strong>Interactive Timeline Chart Available in Web App</strong><br>
            Visit the web application to view the complete portfolio timeline chart with transaction markers and interactive features.
        </div>
        `}
    </div>`;
    }

    /**
     * Generate bar chart section HTML
     */
    generateBarSection(chartImage) {
        return `
    <div class="section">
        <h2>Performance Overview</h2>
        ${chartImage ? `
        <div class="chart-image">
            <img src="${chartImage}" alt="Performance Overview Chart" />
            <p style="font-size: 10px; color: #666; margin-top: 10px;">
                <em>Stacked bar chart showing investment breakdown and returns comparison</em>
            </p>
        </div>
        ` : `
        <div class="chart-notice">
            📊 <strong>Interactive Performance Chart Available in Web App</strong><br>
            Visit the web application to view the interactive performance breakdown chart.
        </div>
        `}
    </div>`;
    }

    /**
     * Generate pie chart section HTML
     */
    generatePieSection(chartImage) {
        return `
    <div class="section">
        <h2>Investment Sources</h2>
        ${chartImage ? `
        <div class="chart-image">
            <img src="${chartImage}" alt="Investment Sources Chart" />
            <p style="font-size: 10px; color: #666; margin-top: 10px;">
                <em>Distribution of investment sources with amounts and percentages</em>
            </p>
        </div>
        ` : `
        <div class="chart-notice">
            📊 <strong>Interactive Investment Sources Chart Available in Web App</strong><br>
            Visit the web application to view the interactive investment sources breakdown chart.
        </div>
        `}
    </div>`;
    }

    /**
     * Generate print-optimized HTML content for PDF export
     */
    generatePrintOptimizedHTML(chartImages = {}, customTitle = "Portfolio_Analysis_Report", exportConfig = null) {
        return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${customTitle}</title>
    <style>
        @media print {
            body { margin: 0; padding: 20px; font-size: 12px; }
            .no-print { display: none !important; }
            .page-break { page-break-before: always; }
            @page { 
                margin: 0.5in; 
                @bottom-left { content: ""; }
                @bottom-right { content: ""; }
                @bottom-center { content: ""; }
                @top-left { content: ""; }
                @top-right { content: ""; }
                @top-center { content: ""; }
            }
        }
        body { font-family: Arial, sans-serif; max-width: 100%; margin: 0 auto; padding: 20px; line-height: 1.4; }
        .header { display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2px solid #333; padding-bottom: 20px; margin-bottom: 20px; }
        .header-left h1 { margin: 0; font-size: 24px; }
        .header-right { text-align: right; font-size: 11px; }
        .header-right p { margin: 2px 0; }
        .metrics-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 25px; }
        .metric-card { border: 1px solid #ddd; border-radius: 6px; padding: 12px; text-align: center; }
        .metric-card.highlight { background-color: #f8f9fa; border-color: #007bff; }
        .metric-label { color: #666; font-size: 10px; text-transform: uppercase; margin-bottom: 4px; }
        .metric-value { font-size: 16px; font-weight: bold; color: #333; }
        .positive { color: #27ae60; }
        .negative { color: #e74c3c; }
        .chart-notice { background: #e7f3ff; border: 1px solid #b3d9ff; border-radius: 6px; padding: 15px; text-align: center; color: #0066cc; margin: 15px 0; font-size: 11px; }
        .chart-image { text-align: center; margin: 20px 0; }
        .chart-image img { max-width: 100%; height: auto; border: 1px solid #ddd; border-radius: 6px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 10px; }
        th, td { border: 1px solid #ddd; padding: 6px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        .section { margin-bottom: 25px; }
        .section h2 { color: #333; border-bottom: 1px solid #ddd; padding-bottom: 8px; font-size: 16px; }
        .footer { margin-top: 30px; text-align: center; color: #666; font-size: 10px; border-top: 1px solid #eee; padding-top: 15px; }
        .footer-logo { margin-bottom: 10px; }
        .footer-logo img { height: 40px; width: auto; }
    </style>
</head>
<body>
    <div class="header">
        <div class="header-left">
            <h1>Portfolio Analysis Report</h1>
        </div>
        <div class="header-right">
            <p><strong>User ID:</strong> ${this.portfolioData.userId}</p>
            <p><strong>Current Price:</strong> ${currencyConverter.getCurrencySymbol(this.calculations.currency)} ${this.calculations.currentPrice.toFixed(2)}</p>
            <p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>
        </div>
    </div>

    ${(() => {
        let sectionsHTML = '';
        
        if (exportConfig && exportConfig.sections) {
            // Use custom configuration
            const enabledSections = exportConfig.sections.filter(s => s.enabled);
            const mainSections = enabledSections.filter(s => s.category === 'main');
            const hasCalculations = enabledSections.some(s => s.category === 'calculation');
            
            // Generate main sections (charts and cards) in their custom order
            mainSections.forEach((section, index) => {
                const sectionId = this.mapSectionIdToLegacy(section.id);
                sectionsHTML += this.generateSectionHTML(sectionId, chartImages, exportConfig);
                
                // Add page break after every 2 main sections
                if ((index + 1) % 2 === 0) {
                    sectionsHTML += '<div class="page-break"></div>\n';
                }
            });
            
            // Page break is handled inside generateCalculationsSection for consistency
        } else {
            // Default behavior - use user section order
            const sectionOrder = this.getUserSectionOrder();
            
            if (sectionOrder.length >= 2) {
                sectionsHTML += this.generateSectionHTML(sectionOrder[0], chartImages, exportConfig);
                sectionsHTML += this.generateSectionHTML(sectionOrder[1], chartImages, exportConfig);
                sectionsHTML += '<div class="page-break"></div>\n';
            }
            
            if (sectionOrder.length >= 4) {
                sectionsHTML += this.generateSectionHTML(sectionOrder[2], chartImages, exportConfig);
                sectionsHTML += this.generateSectionHTML(sectionOrder[3], chartImages, exportConfig);
                sectionsHTML += '<div class="page-break"></div>\n';
            }
        }
        
        return sectionsHTML;
    })()}

    ${this.calculations.blockedShares.byYear && Object.keys(this.calculations.blockedShares.byYear).length > 0 ? `
    <div class="section">
        <h2>Share Information</h2>
        ${this.generateBlockedSharesTable()}
    </div>
    ` : ''}

    <div class="section">
        <h2>Portfolio Details</h2>
        ${this.generatePortfolioTable()}
    </div>

    ${this.generateCalculationsSection(exportConfig)}

    <div class="footer">
        <div class="footer-logo">
            <img src="visuate-rb.png" alt="Visuate++" />
        </div>
        <p><strong>Generated by Visuate++ Portfolio Analysis Tool</strong></p>
        <p>Data processed locally - no information was transmitted to external servers</p>
        <p><em>For interactive charts and full features, visit the web application</em></p>
    </div>
</body>
</html>`;
    }

    /**
     * Generate filename for export
     */
    generateFileName(extension) {
        const now = new Date();
        const dateStr = now.toISOString().split('T')[0].replace(/-/g, '');
        const timeStr = now.toTimeString().split(' ')[0].slice(0, 5).replace(/:/g, '.');
        return `${dateStr}_${timeStr}_Visuate_Portfolio_Analysis_Report.${extension}`;
    }

    /**
     * Download file to user's device
     */
    downloadFile(blob, filename) {
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(link.href);
    }


    /**
     * Generate calculations section HTML for PDF export
     */
    generateCalculationsSection(exportConfig = null) {
        if (!this.calculations || !this.calculations.calculationBreakdowns) {
            return `
    <div class="section">
        <h2>Calculation Breakdown</h2>
        <p>No calculation breakdown data available.</p>
    </div>
            `;
        }

        const breakdowns = this.calculations.calculationBreakdowns;
        const currency = this.calculations.currency; // Currency code for formatting
        
        // Check if we should show calculations (either no config or config has enabled calculations)
        let showCalculations = true;
        if (exportConfig && exportConfig.sections) {
            showCalculations = exportConfig.sections.some(s => s.category === 'calculation' && s.enabled);
        }
        
        if (!showCalculations) {
            return ''; // Don't show calculations section if no calculation sections are enabled
        }
        
        let sectionsHTML = `
    <div class="page-break"></div>
    <div class="section">
        <h2>Calculation Breakdown</h2>
        <p style="font-size: 12px; color: #666; margin-bottom: 20px;">
            Detailed breakdown of how each portfolio metric is calculated, showing all transactions and their current values.
        </p>
        `;

        // Generate breakdown for each card that has data
        const cardConfigs = [
            { id: 'yourInvestment', title: 'Your Investment', breakdown: breakdowns.userInvestment },
            { id: 'companyMatch', title: 'Company Match', breakdown: breakdowns.companyMatch },
            { id: 'freeShares', title: 'Free Shares', breakdown: breakdowns.freeShares },
            { id: 'dividendIncome', title: 'Dividend Income', breakdown: breakdowns.dividendIncome },
            { id: 'totalInvestment', title: 'Total Investment', breakdown: breakdowns.totalInvestment },
            { id: 'currentPortfolio', title: 'Current Portfolio', breakdown: breakdowns.currentPortfolio },
            { id: 'totalSold', title: 'Total Sold', breakdown: breakdowns.totalSold },
            { id: 'xirrYourInvestment', title: 'Annual Return (XIRR) on Your Investment', breakdown: breakdowns.xirrUserInvestment },
            { id: 'xirrTotalInvestment', title: 'Annual Return (XIRR) on Total Investment', breakdown: breakdowns.xirrTotalInvestment }
        ];

        // Filter and reorder cards based on export configuration
        let enabledCardConfigs = cardConfigs;
        if (exportConfig && exportConfig.sections) {
            const enabledCalculations = exportConfig.sections
                .filter(s => s.category === 'calculation' && s.enabled);
            
            // Preserve the custom order from exportConfig.sections
            enabledCardConfigs = [];
            enabledCalculations.forEach(section => {
                const matchingConfig = cardConfigs.find(config => config.id === section.id);
                if (matchingConfig) {
                    enabledCardConfigs.push(matchingConfig);
                }
            });
        }

        enabledCardConfigs.forEach(config => {
            if (config.breakdown && config.breakdown.length > 0) {
                sectionsHTML += this.generateCalculationTable(config.title, config.breakdown, currency);
            }
        });

        sectionsHTML += `
    </div>
        `;

        return sectionsHTML;
    }

    /**
     * Generate HTML table for a specific calculation breakdown
     */
    generateCalculationTable(title, breakdown, currency) {
        if (!breakdown || breakdown.length === 0) return '';

        // Determine table structure based on breakdown type
        const firstItem = breakdown[0];
        let tableHTML = `
        <div class="section calculation-section" style="margin-bottom: 25px; page-break-inside: avoid; break-inside: avoid;">
            <h3 style="color: #333; font-size: 14px; margin-bottom: 10px; border-bottom: 1px solid #ddd; padding-bottom: 5px;">${title}</h3>
        `;

        if (firstItem.hasOwnProperty('currentValue')) {
            // Investment breakdown table (5-column)
            tableHTML += `
            <table class="calculation-table" style="width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 15px; page-break-inside: avoid; break-inside: avoid;">
                <thead>
                    <tr style="background: #f5f5f5;">
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Date</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Category</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Shares</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Investment</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Current Value</th>
                    </tr>
                </thead>
                <tbody>
            `;

            breakdown.forEach(item => {
                const rowStyle = item.isTotal ? 'background: #e8f5e8; font-weight: bold;' : '';
                tableHTML += `
                    <tr style="${rowStyle}">
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.date || ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.category || ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.shares ? this.formatNumber(item.shares, 2) : ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.investment ? currencyConverter.formatCurrency(item.investment, currency, 2) : ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.currentValue ? currencyConverter.formatCurrency(item.currentValue, currency, 2) : ''}</td>
                    </tr>
                `;
            });

        } else if (firstItem.hasOwnProperty('percentage')) {
            // Summary table (Total Investment)
            tableHTML += `
            <table class="calculation-table" style="width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 15px; page-break-inside: avoid; break-inside: avoid;">
                <thead>
                    <tr style="background: #f5f5f5;">
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Category</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Transactions</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Shares</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Investment</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Percentage</th>
                    </tr>
                </thead>
                <tbody>
            `;

            breakdown.forEach(item => {
                const rowStyle = item.isTotal ? 'background: #e8f5e8; font-weight: bold;' : '';
                tableHTML += `
                    <tr style="${rowStyle}">
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.category}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.transactions}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${this.formatNumber(item.shares, 2)}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${currencyConverter.formatCurrency(item.investment, currency, 2)}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.percentage.toFixed(1)}%</td>
                    </tr>
                `;
            });

        } else if (firstItem.hasOwnProperty('cashFlow')) {
            // XIRR cash flow table
            tableHTML += `
            <div style="background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 4px; padding: 10px; margin-bottom: 10px; font-size: 10px;">
                <strong>XIRR (Extended Internal Rate of Return)</strong> calculates the annualized return rate considering the timing of your cash flows.
            </div>
            <table class="calculation-table" style="width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 15px; page-break-inside: avoid; break-inside: avoid;">
                <thead>
                    <tr style="background: #f5f5f5;">
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Date</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Category</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Cash Flow</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Description</th>
                    </tr>
                </thead>
                <tbody>
            `;

            breakdown.forEach(item => {
                const cashFlowColor = item.cashFlow < 0 ? 'color: #dc3545;' : 'color: #28a745;';
                const cashFlowText = (item.cashFlow < 0 ? '-' : '+') + currencyConverter.formatCurrency(Math.abs(item.cashFlow), currency, 2);
                
                tableHTML += `
                    <tr>
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.date}</td>
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.category}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right; ${cashFlowColor}">${cashFlowText}</td>
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.description}</td>
                    </tr>
                `;
            });

        } else if (firstItem.hasOwnProperty('proceeds')) {
            // Total sold table
            tableHTML += `
            <table class="calculation-table" style="width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 15px; page-break-inside: avoid; break-inside: avoid;">
                <thead>
                    <tr style="background: #f5f5f5;">
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Date</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Category</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Shares</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Price</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Proceeds</th>
                    </tr>
                </thead>
                <tbody>
            `;

            breakdown.forEach(item => {
                const rowStyle = item.isTotal ? 'background: #e8f5e8; font-weight: bold;' : '';
                const priceText = item.price ? (item.isTotal ? currencyConverter.formatCurrency(item.price, currency, 2) + ' avg' : currencyConverter.formatCurrency(item.price, currency, 2)) : '';
                
                tableHTML += `
                    <tr style="${rowStyle}">
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.date || ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.category || ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.shares ? this.formatNumber(item.shares, 2) : ''}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${priceText}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.proceeds ? currencyConverter.formatCurrency(item.proceeds, currency, 2) : ''}</td>
                    </tr>
                `;
            });

        } else if (firstItem.hasOwnProperty('profitLoss')) {
            // Current portfolio table
            tableHTML += `
            <table class="calculation-table" style="width: 100%; border-collapse: collapse; font-size: 11px; margin-bottom: 15px; page-break-inside: avoid; break-inside: avoid;">
                <thead>
                    <tr style="background: #f5f5f5;">
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Category</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Shares</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Investment</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Value at Download</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Current Value</th>
                        <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Profit/Loss</th>
                    </tr>
                </thead>
                <tbody>
            `;

            breakdown.forEach(item => {
                const rowStyle = item.isTotal ? 'background: #e8f5e8; font-weight: bold;' : (item.isSubtotal ? 'background: #fff3cd; font-weight: 600;' : '');
                const profitLossColor = item.profitLoss && item.profitLoss < 0 ? 'color: #dc3545;' : 'color: #28a745;';
                const profitLossText = item.profitLoss ? (item.profitLoss > 0 ? '+' : '') + currencyConverter.formatCurrency(item.profitLoss, currency, 2) : '-';
                
                tableHTML += `
                    <tr style="${rowStyle}">
                        <td style="border: 1px solid #ddd; padding: 8px;">${item.category}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.shares ? this.formatNumber(item.shares, 2) : '-'}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.investment ? currencyConverter.formatCurrency(item.investment, currency, 2) : '-'}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.valueAtDownload ? currencyConverter.formatCurrency(item.valueAtDownload, currency, 2) : '-'}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">${item.currentValue ? currencyConverter.formatCurrency(item.currentValue, currency, 2) : '-'}</td>
                        <td style="border: 1px solid #ddd; padding: 8px; text-align: right; ${profitLossColor}">${profitLossText}</td>
                    </tr>
                `;
            });
        }

        tableHTML += `
                </tbody>
            </table>
        </div>
        `;

        return tableHTML;
    }

    /**
     * Format currency value for display
     */

    /**
     * Format number with specified decimal places
     */
    formatNumber(value, decimals = 2) {
        if (value === null || value === undefined || isNaN(value)) return '-';
        return parseFloat(value).toFixed(decimals);
    }
}

// Global exporter instance
const portfolioExporter = new PortfolioExporter();