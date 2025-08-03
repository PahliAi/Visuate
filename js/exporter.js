/**
 * Export functionality for Equate Portfolio Analysis
 * Handles CSV and PDF export of portfolio data and analysis
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
     * Export portfolio analysis as CSV
     */
    exportCSV() {
        if (!this.calculations || !this.portfolioData) {
            throw new Error('No data available for export');
        }

        const csvData = this.generateCSVData();
        const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
        this.downloadFile(blob, this.generateFileName('csv'));
    }

    /**
     * Export portfolio analysis as PDF
     */
    async exportPDF() {
        if (!this.calculations || !this.portfolioData) {
            throw new Error('No data available for export');
        }

        try {
            config.debug('📋 Generating PDF-optimized report...');
            
            // Capture all chart images if available
            const chartImages = {};
            
            // 1. Portfolio Timeline Chart
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
            
            // 2. Performance Bar Chart
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
            
            // 3. Investment Pie Chart
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
            
            // Generate print-optimized HTML content with all charts
            const htmlContent = this.generatePrintOptimizedHTML(chartImages, customTitle);
            
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
            config.error('PDF export error:', error);
            // Fallback to HTML export
            this.exportHTML();
        }
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
     * Generate CSV data
     */
    generateCSVData() {
        const lines = [];
        
        // Header
        lines.push('Equate Portfolio Analysis Report');
        lines.push(`Generated: ${new Date().toLocaleString()}`);
        lines.push(`User ID: ${this.portfolioData.userId}`);
        lines.push('');

        // Summary metrics
        lines.push('PORTFOLIO SUMMARY');
        lines.push('Metric,Value');
        lines.push(`Your Investment,"€${this.calculations.userInvestment.toFixed(2)}"`);
        lines.push(`Company Match,"€${this.calculations.companyInvestment.toFixed(2)}"`);
        lines.push(`Dividend Income,"€${this.calculations.dividendIncome.toFixed(2)}"`);
        lines.push(`Current Value,"€${this.calculations.currentValue.toFixed(2)}"`);
        lines.push(`Total Return,"€${this.calculations.totalReturn.toFixed(2)}"`);
        lines.push(`Return Percentage,"${this.calculations.returnPercentage.toFixed(2)}%"`);
        lines.push(`Annual Growth (CAGR),"${this.calculations.annualGrowth.toFixed(2)}%"`);
        lines.push(`Available Shares,${this.calculations.availableShares}`);
        lines.push(`Total Shares,${this.calculations.totalShares}`);
        lines.push('');

        // Price information
        lines.push('PRICE INFORMATION');
        lines.push(`Current Price,"€${this.calculations.currentPrice.toFixed(2)}"`);
        lines.push(`Price Source,${this.calculations.priceSource}`);
        lines.push(`Price Date,${this.calculations.priceDate}`);
        lines.push('');

        // Blocked shares breakdown
        if (this.calculations.blockedShares.byYear && Object.keys(this.calculations.blockedShares.byYear).length > 0) {
            lines.push('BLOCKED SHARES BY UNLOCK YEAR');
            lines.push('Year,Shares');
            Object.entries(this.calculations.blockedShares.byYear)
                .sort(([a], [b]) => parseInt(a) - parseInt(b))
                .forEach(([year, shares]) => {
                    lines.push(`${year},${shares}`);
                });
            lines.push('');
        }

        // Portfolio details
        lines.push('PORTFOLIO DETAILS');
        lines.push('Allocation Date,Plan,Contribution Type,Cost Basis,Outstanding Qty,Available Qty,Purchase Amount');
        
        this.portfolioData.entries.forEach(entry => {
            lines.push([
                entry.allocationDate || '',
                `"${entry.plan}"`,
                `"${entry.contributionType}"`,
                entry.costBasis.toFixed(2),
                entry.outstandingQuantity,
                entry.availableQuantity,
                entry.purchaseAmount.toFixed(2)
            ].join(','));
        });

        return lines.join('\n');
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
            <p><strong>Current Price:</strong> €${this.calculations.currentPrice.toFixed(2)}</p>
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
                <div class="metric-value">€${this.calculations.userInvestment.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Company Match</div>
                <div class="metric-value">€${this.calculations.companyInvestment.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Dividend Income</div>
                <div class="metric-value">€${this.calculations.dividendIncome.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Investment</div>
                <div class="metric-value">€${this.calculations.totalInvestment.toFixed(2)}</div>
            </div>
            
            <!-- Second Row: Portfolio Performance -->
            <div class="metric-card">
                <div class="metric-label">Current Portfolio</div>
                <div class="metric-value">€${this.calculations.currentValue.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Sold</div>
                <div class="metric-value">€${this.calculations.totalSold.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Value</div>
                <div class="metric-value">€${this.calculations.totalValue.toFixed(2)}</div>
            </div>
            <div class="metric-card highlight">
                <div class="metric-label">Total Return</div>
                <div class="metric-value ${this.calculations.totalReturn >= 0 ? 'positive' : 'negative'}">
                    €${this.calculations.totalReturn.toFixed(2)}
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
            html += `<td>€${entry.costBasis.toFixed(2)}</td>`;
            html += `<td>${entry.outstandingQuantity}</td>`;
            html += `<td>${entry.availableQuantity}</td>`;
            html += `<td>€${entry.purchaseAmount.toFixed(2)}</td>`;
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
        return savedOrder ? JSON.parse(savedOrder) : ['cards', 'timeline', 'bar', 'pie'];
    }

    /**
     * Generate section HTML content
     */
    generateSectionHTML(sectionType, chartImages) {
        const sectionMapping = {
            'cards': () => this.generateCardsSection(),
            'timeline': () => this.generateTimelineSection(chartImages.portfolioChart),
            'bar': () => this.generateBarSection(chartImages.performanceChart),
            'pie': () => this.generatePieSection(chartImages.pieChart)
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
                <div class="metric-value">€${this.calculations.userInvestment.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Company Match</div>
                <div class="metric-value">€${(this.calculations.companyMatch || 0).toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Free Shares</div>
                <div class="metric-value">€${(this.calculations.freeShares || 0).toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Dividend Income</div>
                <div class="metric-value">€${this.calculations.dividendIncome.toFixed(2)}</div>
            </div>
            
            <!-- Second Row: Portfolio Performance -->
            <div class="metric-card highlight" style="background-color: #fff8e6; border-color: #f39c12;">
                <div class="metric-label">Total Investment</div>
                <div class="metric-value">€${this.calculations.totalInvestment.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Current Portfolio</div>
                <div class="metric-value">€${this.calculations.currentValue.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Sold</div>
                <div class="metric-value">€${this.calculations.totalSold.toFixed(2)}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Value</div>
                <div class="metric-value">€${this.calculations.totalValue.toFixed(2)}</div>
            </div>
            
            <!-- Third Row: Returns -->
            <div class="metric-card highlight">
                <div class="metric-label">Return on Your Investment</div>
                <div class="metric-value ${this.calculations.totalReturn >= 0 ? 'positive' : 'negative'}">
                    €${this.calculations.totalReturn.toFixed(2)}
                </div>
            </div>
            <div class="metric-card highlight" style="background-color: #fff8e6; border-color: #f39c12;">
                <div class="metric-label">Return on Total Investment</div>
                <div class="metric-value ${(this.calculations.returnOnTotalInvestment || 0) >= 0 ? 'positive' : 'negative'}">
                    €${(this.calculations.returnOnTotalInvestment || 0).toFixed(2)}
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
    generatePrintOptimizedHTML(chartImages = {}, customTitle = "Portfolio_Analysis_Report") {
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
            <p><strong>Current Price:</strong> €${this.calculations.currentPrice.toFixed(2)}</p>
            <p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>
        </div>
    </div>

    ${(() => {
        const sectionOrder = this.getUserSectionOrder();
        return `
    ${this.generateSectionHTML(sectionOrder[0], chartImages)}
    ${this.generateSectionHTML(sectionOrder[1], chartImages)}
    <div class="page-break"></div>
    
    ${this.generateSectionHTML(sectionOrder[2], chartImages)}
    ${this.generateSectionHTML(sectionOrder[3], chartImages)}
    <div class="page-break"></div>
        `;
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
     * Export portfolio timeline data as CSV
     */
    exportTimelineCSV(timelineData) {
        if (!timelineData || timelineData.length === 0) {
            throw new Error('No timeline data available for export');
        }

        const lines = [];
        lines.push('Portfolio Timeline Data');
        lines.push(`Generated: ${new Date().toLocaleString()}`);
        lines.push('');
        lines.push('Date,Share Price,Total Shares,Portfolio Value,Has Transaction');

        timelineData.forEach(point => {
            lines.push([
                point.date,
                point.price.toFixed(2),
                point.shares,
                point.value.toFixed(2),
                point.hasTransaction ? 'Yes' : 'No'
            ].join(','));
        });

        const csvData = lines.join('\n');
        const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
        this.downloadFile(blob, `portfolio-timeline-${new Date().toISOString().split('T')[0]}.csv`);
    }
}

// Global exporter instance
const portfolioExporter = new PortfolioExporter();