# Equate - Portfolio Analysis Tool

A local web application for analyzing EquatePlus portfolio data. All processing happens in your browser.

## How to Use

### Step 1: Download Files from EquatePlus

1. Log in to the EquatePlus website
2. Go to your account and select 'Preferences'
3. Set language to English
4. Go to tab 'Overview - Plans & Trading' and click the button next to 'YOUR PORTFOLIO - Estimated Gross Value'
5. If you have sold shares: go to tab 'Library - Transactions & Records' and click the button next to 'Transaction history'
6. Close EquatePlus website

### Step 2: Upload and Analyze

1. Open `index.html` in your web browser (or serve locally for best results)
2. Go to the 'Upload' tab
3. Upload your PortfolioDetails Excel file (required)
4. Upload CompletedTransactions file (optional, only if you've sold shares)
5. Click "Generate Analysis"

## File Requirements

### Required Files
- **PortfolioDetails_[UserID].xlsx** - Your current portfolio holdings

### Optional Files  
- **CompletedTransactions_[UserID].xlsx** - Transaction history for sold shares
- **hist.xlsx** - Historical share price data (included in project)

## What You'll See

The analysis shows:
- Your total investment amount
- Company match and free shares
- Dividend income
- Current portfolio value
- Profit/loss (absolute and percentage)
- Annual growth rate (CAGR)
- Available vs blocked shares

## Updating Share Price

To get current profit/loss calculations:
1. Find the current share price from EquatePlus or your broker
2. Enter it in the "Price" input field in the Results tab
3. Click "Update Price" for instant recalculation

## Data Storage

- Your uploaded data is stored locally in your browser
- No data is sent to external servers
- Analysis persists between browser sessions
- Use "Clear Data" button to remove stored information

## Exporting Results

1. Go to Results tab after analysis
2. Click "Export PDF" to save analysis as HTML file
3. Data includes all metrics and calculations  

## Running Locally

For best results, serve the files locally:

```bash
# Using Python
python -m http.server 8000

# Using Node.js
npx http-server -p 8000
```

Then open: `http://localhost:8000`

## Common Issues

- **File not recognized**: Make sure Excel files are unmodified from EquatePlus
- **UserID mismatch**: Both files must be from the same user account  
- **Charts not loading**: Try refreshing the page or using a local server
- **Export not working**: Make sure analysis has completed first

## Browser Requirements

Works with modern browsers (Chrome, Firefox, Safari, Edge). Requires JavaScript enabled.

## Files Structure

- `index.html` - Main application
- `js/` - Application logic and calculations
- `css/` - Styling and themes
- `hist.xlsx` - Historical share price data