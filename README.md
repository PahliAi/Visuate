# Equate User Manual

## Overview

Equate analyzes your EquatePlus share portfolio to calculate profits, losses, and performance metrics. All processing happens locally in your browser - no data leaves your computer.

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

1. **Upload Portfolio File**
   - Drag and drop or click to select your `PortfolioDetails.xlsx` file
   - Green checkmark indicates successful upload

2. **Upload Transaction File** (if you have one)
   - Drag and drop or click to select your `CompletedTransactions.xlsx` file
   - Green checkmark indicates successful upload

3. **Generate Analysis**
   - Click the "Generate Analysis" button
   - Wait for processing to complete
   - Results tab will automatically appear

## Understanding Your Results

### Portfolio Summary Cards
- **Your Investment** - Total money you've invested
- **Company Match** - Free shares and matching contributions from your employer
- **Dividend Income** - Reinvested dividends received
- **Current Portfolio Value** - Today's market value
- **Total Return** - Your profit/loss amount
- **Return Percentage** - Your profit/loss as a percentage
- **Annual Growth (CAGR)** - Compound annual growth rate

### Charts and Visualizations
- **Timeline Chart** - Shows portfolio value growth over time
- **Performance Breakdown** - Bar chart showing investment sources and returns  
- **Investment Distribution** - Pie chart of how your money was invested

### Portfolio Details Table
- **Allocation Date** - When shares were granted
- **Cost Basis** - Price you paid per share
- **Outstanding Shares** - Total shares you own
- **Available Shares** - Shares you can sell (not blocked)
- **Current Value** - Market value of each allocation

## Advanced Features

### Manual Price Updates
To see current profit/loss with today's share price:
1. Find current share price from EquatePlus or your broker
2. Enter price in the "Current Price" field
3. All calculations update automatically
4. Manual prices are saved for future sessions

### Exporting Reports
1. **CSV Export** - Click "Export CSV" for spreadsheet data
2. **PDF Export** - Click "Export PDF" for complete report with charts

### Section Reordering
- Drag the section icons to reorder charts and tables
- Layout preference is saved automatically

### Multiple Languages
- Application automatically detects your file language
- Supports 13 languages: English, German, Dutch, French, Spanish, Italian, Polish, Turkish, Portuguese, Czech, Romanian, Croatian, Indonesian
- UI language can be changed using the language selector

## Data Privacy & Storage

### Privacy
- **100% Local Processing** - All calculations happen in your browser
- **No Data Transmission** - Files never leave your computer
- **No Account Required** - No registration or login needed

### Data Storage  
- **Browser Cache** - Data stored locally for faster future loading
- **Persistent Sessions** - Results saved between browser sessions
- **Clear Data** - Use "Clear All Data" button to remove stored information
- **Offline Capable** - Works without internet connection after first load

## Currency Support

### Automatic Detection
- Currencies are automatically detected from your Excel files
- Common currencies: EUR (€), USD ($), GBP (£), ARS ($), CAD (C$), and 20+ others

### Currency Validation
- **Portfolio and transaction files must use the same currency**
- **Currency mismatch error** - Transaction file rejected if currencies differ
- **Portfolio always wins** - Portfolio currency takes precedence

### Mixed Currency Files
- If multiple currencies detected in one file, you'll be prompted to choose
- System shows numbers only if currency cannot be determined

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
- **CSV formatting** - Open CSV files in Excel or Google Sheets for proper formatting

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
- ✅ Complete privacy (local processing only)
- ✅ Multi-language support (13 languages)
- ✅ Multi-currency support (25+ currencies)
- ✅ Interactive charts and visualizations
- ✅ PDF and CSV export capabilities
- ✅ Manual price override functionality
- ✅ Automatic data caching and persistence

For technical details, see `TECHNICAL_OVERVIEW.md`