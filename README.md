# Equate - Portfolio Analysis

## Overview

Equate (also known as Visuate ++) is a 100% client-side portfolio analysis tool for EquatePlus data. All processing happens locally in your browser - no data leaves your computer. The application features a modern service-based architecture with comprehensive multi-language and multi-currency support.

## Getting Started

### Step 1: Download Your Files from EquatePlus

1. **Portfolio File (Required)**
   - Log into EquatePlus website
   - Go to **Overview → Plans & Trading**
   - Click download button next to "YOUR PORTFOLIO - Estimated Gross Value"
   - Save the `PortfolioDetails_[UserID].xlsx` file

2. **Transaction File (Optional)**
   - Only needed if you've sold shares
   - Go to **Library → Transactions & Records** 
   - Click download button next to "Transaction history"
   - Save the `CompletedTransactions_[UserID].xlsx` file

### Step 2: Run the Application

**Option A: Simple (Double-click)**
- Double-click `index.html` to open in your browser

**Option B: Local Server (Recommended)**
```bash
# Using Python
python3 -m http.server 8000

# Using Node.js  
npx http-server -p 8000
```
Then open: `http://localhost:8000`

### Step 3: Upload and Analyze

The application features four main tabs:

1. **📋 How to** - Step-by-step instructions with screenshots
2. **📤 Upload** - File upload interface 
3. **📊 Results** - Portfolio analysis and charts
4. **🧮 Calculations** - Detailed calculation breakdowns

**Upload Process:**
1. **Upload Portfolio File**
   - Drag and drop or click to select your `PortfolioDetails.xlsx` file
   - Green checkmark indicates successful upload

2. **Upload Transaction File** (if you have one)
   - Drag and drop or click to select your `CompletedTransactions.xlsx` file
   - Green checkmark indicates successful upload

3. **Generate Analysis**
   - Click the "Generate Analysis" button
   - Wait for processing to complete
   - Results and Calculations tabs will automatically become available

## Understanding Your Results

### Portfolio Summary Cards
- **Your Investment** - Total money you've invested
- **Company Match** - Free shares and matching contributions from your employer
- **Dividend Income** - Reinvested dividends received
- **Current Portfolio Value** - Today's market value
- **Total Return** - Your profit/loss amount
- **Return Percentage** - Your profit/loss as a percentage
- **Annual Return (XIRR)** - Internal rate of return for your investment and total portfolio
- **Available Shares** - Shares you can sell (not blocked)

### Charts and Visualizations
- **Timeline Chart** - Shows portfolio value growth over time with stock price trendline (click legend to show/hide individual lines)
- **Performance Breakdown** - Bar chart showing investment sources and returns  
- **Investment Distribution** - Pie chart of how your money was invested
- **Application Flow** - Interactive flowchart showing the complete application workflow

### Portfolio Details Table
- **Allocation Date** - When shares were granted
- **Cost Basis** - Price you paid per share
- **Outstanding Shares** - Total shares you own
- **Available Shares** - Shares you can sell (not blocked)
- **Current Value** - Market value of each allocation

## Advanced Features

### Currency Switching
- **Dynamic Currency Selection**: Switch between detected currencies using the currency selector dropdown
- **Real-time Conversion**: All values automatically convert when switching currencies
- **Historical Accuracy**: Uses historical exchange rates for precise conversions
- **Preference Persistence**: Selected currency is saved for future sessions

### Manual Price Updates
To see current profit/loss with today's share price:
1. Find current share price from EquatePlus or your broker
2. Enter price in the "Current Price" field in the Results tab
3. All calculations update automatically across all tabs
4. Manual prices are saved for future sessions

### Detailed Calculations Tab
- **🧮 Calculations Tab**: Comprehensive breakdown of all financial calculations
- **Accordion Interface**: Expandable/collapsible sections for each metric
- **Interactive Controls**: Expand/collapse all sections with one click
- **Calculation Transparency**: See exactly how each metric is calculated
- **Multi-language Support**: All calculation details adapt to selected language

### PDF Export System
1. **Custom Export Builder** - Click "Export PDF" to access advanced options
2. **Section Selection** - Choose which sections to include in your report
3. **Chart Integration** - All charts are embedded in high-quality format
4. **Professional Formatting** - Print-ready reports with proper styling

### Section Reordering
- Drag the section icons in the Results tab to reorder charts and tables
- Layout preference is saved automatically and persists across sessions
- Affects both display order and PDF export order

### Multi-Language System
- **Automatic Detection**: Application automatically detects your file language
- **13 Language Support**: English, German, Dutch, French, Spanish, Italian, Polish, Turkish, Portuguese, Czech, Romanian, Croatian, Indonesian, Chinese
- **Dynamic Switching**: UI language can be changed using the language selector
- **Complete Translation**: All tabs, buttons, and calculation details are fully translated
- **Persistent Preferences**: Language preference is saved and persists across browser sessions

## Data Privacy & Storage

### Privacy
- **100% Local Processing** - All calculations happen in your browser
- **No Data Transmission** - Files never leave your computer
- **No Account Required** - No registration or login needed

### Data Storage  
- **Browser Cache** - Data stored locally for faster future loading
- **Persistent Sessions** - Results and preferences saved between browser sessions
- **Clear Data** - Use "Clear All Data" button to remove stored information
- **Offline Capable** - Works without internet connection after first load

## Currency Support

### Advanced Currency System
- **25+ Supported Currencies**: EUR, USD, GBP, CAD, ARS, and many more
- **Automatic Detection**: Currencies are automatically detected from Excel cell formatting
- **Dynamic Switching**: Switch between any detected currencies using the dropdown selector
- **Historical Exchange Rates**: Uses accurate historical rates for conversions
- **Real-time Updates**: All metrics update instantly when switching currencies

### Currency Architecture
- **Portfolio Always Wins**: Portfolio file currency takes precedence for validation
- **File Consistency**: Portfolio and transaction files must use the same base currency
- **Smart Detection**: Multiple detection methods including formatting, headers, and values
- **Fallback Options**: Manual currency selection if auto-detection fails

### Currency Selector Features
- **Clean Interface**: Simple dropdown showing "Currency - Name" format (quality scores removed)
- **Sorted Display**: Currencies sorted alphabetically for easy selection
- **Persistent Choice**: Selected currency saved and restored between sessions
- **Validation Feedback**: Clear error messages for currency mismatches

## Troubleshooting

### File Upload Issues
- **"File not recognized"** - Ensure Excel files are unmodified from EquatePlus
- **"UserID mismatch"** - Portfolio and transaction files must be from same EquatePlus account
- **Currency mismatch** - Files must use same currency (EUR, USD, etc.)

### Display Issues  
- **Charts not loading** - Refresh page or try local server method
- **Missing data** - Ensure files are complete Excel downloads from EquatePlus
- **Performance slow** - Use local server for better performance

### Browser Issues
- **JavaScript required** - Enable JavaScript in browser settings
- **Modern browser needed** - Use Chrome, Firefox, Safari, or Edge
- **CORS errors** - Use local server instead of file:// protocol

### Export Issues
- **Export not working** - Ensure analysis completed successfully first
- **PDF missing charts** - Charts may take a moment to render before export
- **Export customization** - Use "Customize Export" to select specific sections for your report

## Technical Requirements

### Browser Support
- Chrome 80+, Firefox 75+, Safari 13+, Edge 80+
- JavaScript must be enabled
- IndexedDB support required (for data storage)

### File Format Support  
- Excel files (.xlsx) from EquatePlus only
- Files must be unmodified from EquatePlus download
- Maximum file size: 10MB per file

### Performance
- Processes files up to 1000+ portfolio entries
- Analysis typically completes in 2-5 seconds
- Local server recommended for optimal performance

---

## Quick Reference

### Essential Steps
1. Download `PortfolioDetails.xlsx` from EquatePlus
2. Download `CompletedTransactions.xlsx` if you've sold shares  
3. Open Equate application
4. Upload files and click "Generate Analysis"
5. Review results and export if needed

### Key Features
- ✅ **Complete Privacy**: 100% local processing, no data transmission
- ✅ **Four-Tab Interface**: How to, Upload, Results, and detailed Calculations
- ✅ **Dynamic Currency Switching**: Real-time conversion between 25+ currencies
- ✅ **Multi-Language Support**: 13 languages with automatic detection and persistent preferences
- ✅ **Interactive Charts**: Timeline, performance bars, and pie charts with toggleable visibility
- ✅ **Detailed Calculations Tab**: Comprehensive breakdown with accordion interface
- ✅ **Advanced XIRR Calculations**: Annualized returns for investment analysis
- ✅ **Professional PDF Export**: Customizable reports with chart integration
- ✅ **Service-Based Architecture**: Modern EventBus coordination system
- ✅ **Manual Price Override**: Update share prices with automatic recalculation
- ✅ **Intelligent Caching**: Fast loading with automatic data validation
- ✅ **Section Reordering**: Drag-and-drop interface customization

For technical details, see `TECHNICAL_OVERVIEW.md`

## License

**PolyForm Noncommercial License 1.0.0** - This software is licensed for noncommercial use only.

🔐 **License Authority**: [Master License Record](https://github.com/PahliAi/Equate/tree/license-master-v1.0)  
📅 **Originally Established**: August 14, 2025 (cryptographically verified)  
📄 **Full License Text**: See [LICENSE](LICENSE) file in this repository  
🏷️ **Git Proof**: Commit `8e4f5bd43eb` with permanent verification

**Permitted Uses**: Personal use, research, education, nonprofit organizations  
**Restricted Uses**: Commercial use, resale, commercial distribution  

For commercial licensing inquiries, please contact the repository maintainer.

---

*This software is protected under PolyForm Noncommercial License. The original license establishment is permanently recorded with full git history proof in our development repository.*