import pandas as pd
import yfinance as yf
import requests
from datetime import datetime, timedelta
import os
import json

EXCEL_FILE = "hist_base.xlsx"

def read_stocks_from_metadata():
    """Read stock configuration dynamically from hist_base.xlsx metadata"""
    if not os.path.exists(EXCEL_FILE):
        print(f"⚠️ {EXCEL_FILE} not found, using fallback configuration")
        return {
            "Allianz Share": "ALV.DE",
            "IBM Share": "IBM", 
            "Unilever Share": "UL"
        }, {"IBM Share", "Unilever Share"}
    
    try:
        base_data = pd.read_excel(EXCEL_FILE, sheet_name='Shares', header=None)
        
        # Read metadata rows (headers are row 0, metadata in rows 1-2)
        instruments = base_data.iloc[0][1:].dropna().tolist()  # Row 0: Instrument names (headers)
        companies = base_data.iloc[1][1:].dropna().tolist()    # Row 1: Company names (metadata)  
        tickers = base_data.iloc[2][1:].dropna().tolist()      # Row 2: Ticker symbols (metadata)
        
        # Build dynamic STOCKS dict
        stocks = {}
        usd_stocks = set()
        
        for instrument, company, ticker in zip(instruments, companies, tickers):
            stocks[instrument] = str(ticker)  # Ensure ticker is string
            
            # Determine if stock needs USD to EUR conversion
            # Allianz (ALV.DE) is EUR-based, others are typically USD
            ticker_str = str(ticker)
            if not ticker_str.endswith('.DE') and not ticker_str.endswith('.L') and not ticker_str.endswith('.PA'):
                usd_stocks.add(instrument)
        
        print(f"📊 Dynamically loaded {len(stocks)} stocks from metadata:")
        for instrument, ticker in stocks.items():
            currency_note = "(USD→EUR)" if instrument in usd_stocks else "(EUR)"
            print(f"   • {instrument}: {ticker} {currency_note}")
        
        return stocks, usd_stocks
        
    except Exception as e:
        print(f"⚠️ Error reading metadata: {e}")
        print("   Using fallback configuration")
        return {
            "Allianz Share": "ALV.DE",
            "IBM Share": "IBM", 
            "Unilever Share": "UL"
        }, {"IBM Share", "Unilever Share"}

# Initialize stocks configuration
STOCKS, USD_STOCKS = read_stocks_from_metadata()

# ⏳ Hoeveel dagen terug kijken om te bepalen of update nodig is
DAYS_THRESHOLD = 1

# Currency exchange rates API (EUR base) - gets all 160+ supported currencies
CURRENCY_API_URL = "https://api.exchangerate-api.com/v4/latest/EUR"

def get_yfinance_ticker(currency_code):
    """
    Dynamically generate yfinance ticker symbol for EUR base currency pairs.
    e.g., 'USD' -> 'EURUSD=X'
    """
    return f"EUR{currency_code}=X"

def get_usd_to_eur_rate(currency_df, date_str):
    """
    Get USD to EUR exchange rate for a specific date.
    Returns None if rate not available.
    """
    if currency_df.empty:
        return None
    
    # Find the rate for the exact date or the most recent previous date
    currency_df_temp = currency_df.copy()
    currency_df_temp['Date'] = pd.to_datetime(currency_df_temp['Date'], format='%d-%m-%Y', errors='coerce')
    target_date = pd.to_datetime(date_str, format='%d-%m-%Y', errors='coerce')
    
    # Filter to dates up to and including the target date
    available_dates = currency_df_temp[currency_df_temp['Date'] <= target_date]
    
    if available_dates.empty:
        return None
    
    # Get the most recent date
    most_recent = available_dates.loc[available_dates['Date'].idxmax()]
    
    if 'USD' in most_recent and pd.notna(most_recent['USD']):
        # Convert USD rate to EUR (since EUR is base=1, USD rate is EUR/USD, we need 1/USD for USD/EUR)
        return 1.0 / float(most_recent['USD'])
    
    return None

def convert_usd_to_eur(usd_price, exchange_rate):
    """
    Convert USD price to EUR using the exchange rate.
    """
    if exchange_rate is None or pd.isna(usd_price):
        return usd_price
    
    # Round the final EUR result to 2 decimals for consistency
    return round(float(usd_price) * exchange_rate, 2)

def write_excel_with_metadata(shares_df, currency_df, excel_file):
    """
    Write Excel file preserving metadata structure:
    Row 0: Date | Allianz Share | IBM Share | Unilever Share (headers)
    Row 1: Company | Allianz | IBM | Unilever (metadata)
    Row 2: Ticker | ALV.DE | IBM | UL (metadata)  
    Row 3+: Actual date data
    """
    try:
        # Read existing metadata if file exists
        metadata_rows = []
        if os.path.exists(excel_file):
            try:
                base_data = pd.read_excel(excel_file, sheet_name='Shares', header=None)
                if len(base_data) >= 3:
                    # Extract metadata rows (assuming they exist)
                    metadata_rows = [
                        base_data.iloc[1].tolist(),  # Company row
                        base_data.iloc[2].tolist()   # Ticker row
                    ]
            except Exception as e:
                print(f"   ⚠️ Could not read existing metadata: {e}")
        
        # Create metadata rows if not found
        if not metadata_rows:
            print("   📝 Creating default metadata rows")
            # Default metadata based on current STOCKS configuration
            instruments = ['Date'] + list(STOCKS.keys())
            companies = ['Company'] + [inst.split(' ')[0] for inst in STOCKS.keys()]  # Allianz, IBM, Unilever
            tickers = ['Ticker'] + list(STOCKS.values())  # ALV.DE, IBM, UL
            metadata_rows = [companies, tickers]
        
        # Clean data first to avoid NaN/INF issues
        shares_clean = shares_df.fillna('')  # Replace NaN with empty strings
        
        # Write to Excel with metadata preservation
        with pd.ExcelWriter(excel_file, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
            workbook = writer.book
            
            # Create Shares sheet with metadata
            shares_worksheet = workbook.add_worksheet('Shares')
            writer.sheets['Shares'] = shares_worksheet
            
            # Write headers (row 0)
            headers = shares_df.columns.tolist()
            for col_num, header in enumerate(headers):
                shares_worksheet.write(0, col_num, header)
            
            # Write metadata rows (rows 1-2)
            for row_idx, metadata_row in enumerate(metadata_rows):
                for col_num, value in enumerate(metadata_row):
                    if col_num < len(headers):  # Don't exceed column count
                        shares_worksheet.write(row_idx + 1, col_num, value)
            
            # Write actual data starting from row 3
            for row_idx, (_, row) in enumerate(shares_clean.iterrows()):
                for col_num, value in enumerate(row):
                    shares_worksheet.write(row_idx + 3, col_num, value)
            
            # Write Currency sheet normally (no metadata needed)
            if not currency_df.empty:
                currency_df.to_excel(writer, sheet_name='Currency', index=False)
        
        print(f"   ✅ Saved with metadata structure preserved")
        
    except Exception as e:
        print(f"   ❌ Error writing Excel with metadata: {e}")
        # Fallback to simple write
        with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
            shares_df.to_excel(writer, sheet_name='Shares', index=False)
            if not currency_df.empty:
                currency_df.to_excel(writer, sheet_name='Currency', index=False)

def analyze_data_quality(shares_df, currency_df, days=5):
    """
    Analyze data quality for the last N days and generate report.
    Returns comprehensive data quality metrics and last N days of data.
    """
    # Convert Date columns to datetime for analysis
    shares_df = shares_df.copy()
    currency_df = currency_df.copy()
    
    shares_df['Date'] = pd.to_datetime(shares_df['Date'], format='%d-%m-%Y', errors='coerce')
    currency_df['Date'] = pd.to_datetime(currency_df['Date'], format='%d-%m-%Y', errors='coerce')
    
    # Get last N days
    end_date = datetime.now()
    start_date = end_date - timedelta(days=days)
    
    # Filter for recent data
    recent_shares = shares_df[shares_df['Date'] >= start_date].sort_values('Date')
    recent_currency = currency_df[currency_df['Date'] >= start_date].sort_values('Date')
    
    # Data quality analysis
    quality_report = {
        'analysis_period': f"{start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')}",
        'shares_data': {},
        'currency_data': {},
        'last_5_days_shares': [],
        'last_5_days_currency': [],
        'data_gaps': [],
        'overall_health': 'HEALTHY'
    }
    
    # Analyze each stock
    for stock_name in STOCKS.keys():
        if stock_name in shares_df.columns:
            # Overall statistics
            total_records = len(shares_df)
            non_null_records = shares_df[stock_name].notna().sum()
            coverage = (non_null_records / total_records) * 100 if total_records > 0 else 0
            
            # Find first and last dates with data
            stock_data = shares_df[shares_df[stock_name].notna()]
            first_date = stock_data['Date'].min() if not stock_data.empty else None
            last_date = stock_data['Date'].max() if not stock_data.empty else None
            
            # Check for recent data gaps
            recent_stock_data = recent_shares[stock_name]
            recent_gaps = recent_stock_data.isna().sum()
            
            quality_report['shares_data'][stock_name] = {
                'total_records': non_null_records,
                'coverage_pct': round(coverage, 1),
                'first_date': first_date.strftime('%d-%m-%Y') if first_date else None,
                'last_date': last_date.strftime('%d-%m-%Y') if last_date else None,
                'recent_gaps': recent_gaps,
                'staleness_days': (end_date - last_date).days if last_date else 999
            }
            
            # Check if data is stale (>3 days for Allianz, >5 days for others)
            staleness_threshold = 3 if 'Allianz' in stock_name else 5
            if last_date and (end_date - last_date).days > staleness_threshold:
                quality_report['data_gaps'].append(f"{stock_name} data is stale ({(end_date - last_date).days} days)")
                quality_report['overall_health'] = 'WARNING'
    
    # Analyze currency data
    if not currency_df.empty:
        total_currency_records = len(currency_df)
        latest_currency = currency_df['Date'].max() if not currency_df.empty else None
        
        # Key currencies to monitor
        key_currencies = ['USD', 'GBP', 'CAD', 'CNY', 'JPY']
        
        quality_report['currency_data'] = {
            'total_records': total_currency_records,
            'latest_date': latest_currency.strftime('%d-%m-%Y') if latest_currency else None,
            'staleness_days': (end_date - latest_currency).days if latest_currency else 999,
            'key_currencies': {}
        }
        
        # Check key currency availability
        for currency in key_currencies:
            if currency in currency_df.columns:
                currency_records = currency_df[currency].notna().sum()
                coverage = (currency_records / total_currency_records) * 100 if total_currency_records > 0 else 0
                quality_report['currency_data']['key_currencies'][currency] = {
                    'records': currency_records,
                    'coverage_pct': round(coverage, 1)
                }
        
        # Check currency staleness (>1 day)
        if latest_currency and (end_date - latest_currency).days > 1:
            quality_report['data_gaps'].append(f"Currency data is stale ({(end_date - latest_currency).days} days)")
            quality_report['overall_health'] = 'WARNING'
    
    # Prepare last 5 days data for email
    for _, row in recent_shares.iterrows():
        day_data = {'Date': row['Date'].strftime('%d-%m-%Y')}
        for stock_name in STOCKS.keys():
            if stock_name in row:
                day_data[stock_name] = f"€{row[stock_name]:.2f}" if pd.notna(row[stock_name]) else "N/A"
        quality_report['last_5_days_shares'].append(day_data)
    
    for _, row in recent_currency.iterrows():
        day_data = {'Date': row['Date'].strftime('%d-%m-%Y')}
        for currency in ['USD', 'GBP', 'CAD', 'CNY']:
            if currency in row:
                day_data[currency] = f"{row[currency]:.4f}" if pd.notna(row[currency]) else "N/A"
        quality_report['last_5_days_currency'].append(day_data)
    
    # Set overall health based on gaps
    if len(quality_report['data_gaps']) > 2:
        quality_report['overall_health'] = 'CRITICAL'
    
    return quality_report

def generate_html_email_report(quality_report, operation_summary):
    """
    Generate HTML email report with data quality analysis.
    """
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
            .header {{ background: #2c3e50; color: white; padding: 20px; text-align: center; }}
            .status {{ padding: 10px; margin: 10px 0; border-radius: 5px; text-align: center; }}
            .healthy {{ background: #d4edda; color: #155724; }}
            .warning {{ background: #fff3cd; color: #856404; }}
            .critical {{ background: #f8d7da; color: #721c24; }}
            .section {{ margin: 20px 0; }}
            .table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
            .table th, .table td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
            .table th {{ background: #f8f9fa; }}
            .metric {{ display: inline-block; margin: 10px; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }}
            ul {{ padding-left: 20px; }}
            .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 0.9em; }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>📈 Equate Stock & Currency Update Report</h1>
            <p>Daily automated data update - {datetime.now().strftime('%d %B %Y at %H:%M UTC')}</p>
        </div>
        
        <div class="status {'healthy' if quality_report['overall_health'] == 'HEALTHY' else 'warning' if quality_report['overall_health'] == 'WARNING' else 'critical'}">
            <h2>System Status: {quality_report['overall_health']}</h2>
        </div>
        
        <div class="section">
            <h2>📊 Operation Summary</h2>
            <div class="metric">
                <strong>New Stock Records:</strong> {operation_summary.get('new_stock_records', 0)}
            </div>
            <div class="metric">
                <strong>Currency Updates:</strong> {operation_summary.get('currency_updates', 0)}
            </div>
            <div class="metric">
                <strong>Gaps Filled:</strong> {operation_summary.get('gaps_filled', 0)}
            </div>
            <div class="metric">
                <strong>USD Conversions:</strong> {operation_summary.get('conversions', 0)}
            </div>
        </div>
    """
    
    # Data Quality Issues
    if quality_report['data_gaps']:
        html_content += f"""
        <div class="section">
            <h2>⚠️ Data Quality Issues</h2>
            <ul>
        """
        for gap in quality_report['data_gaps']:
            html_content += f"<li>{gap}</li>"
        html_content += "</ul></div>"
    
    # Last 5 Days Stock Prices
    if quality_report['last_5_days_shares']:
        html_content += """
        <div class="section">
            <h2>📈 Stock Prices - Last 5 Days</h2>
            <table class="table">
                <tr><th>Date</th><th>Allianz (EUR)</th><th>IBM (EUR)</th><th>Unilever (EUR)</th></tr>
        """
        for day in quality_report['last_5_days_shares'][-5:]:  # Last 5 days
            html_content += f"""
                <tr>
                    <td>{day['Date']}</td>
                    <td>{day.get('Allianz Share', 'N/A')}</td>
                    <td>{day.get('IBM Share', 'N/A')}</td>
                    <td>{day.get('Unilever Share', 'N/A')}</td>
                </tr>
            """
        html_content += "</table></div>"
    
    # Last 5 Days Currency Rates
    if quality_report['last_5_days_currency']:
        html_content += """
        <div class="section">
            <h2>💱 Exchange Rates - Last 5 Days (Base: EUR)</h2>
            <table class="table">
                <tr><th>Date</th><th>USD</th><th>GBP</th><th>CAD</th><th>CNY</th></tr>
        """
        for day in quality_report['last_5_days_currency'][-5:]:  # Last 5 days
            html_content += f"""
                <tr>
                    <td>{day['Date']}</td>
                    <td>{day.get('USD', 'N/A')}</td>
                    <td>{day.get('GBP', 'N/A')}</td>
                    <td>{day.get('CAD', 'N/A')}</td>
                    <td>{day.get('CNY', 'N/A')}</td>
                </tr>
            """
        html_content += "</table></div>"
    
    # Data Coverage Summary
    html_content += """
    <div class="section">
        <h2>📊 Data Coverage Summary</h2>
        <table class="table">
            <tr><th>Asset</th><th>Total Records</th><th>Coverage %</th><th>First Date</th><th>Last Date</th><th>Staleness (days)</th></tr>
    """
    
    for stock_name, data in quality_report['shares_data'].items():
        html_content += f"""
            <tr>
                <td>{stock_name}</td>
                <td>{data['total_records']}</td>
                <td>{data['coverage_pct']}%</td>
                <td>{data['first_date'] or 'N/A'}</td>
                <td>{data['last_date'] or 'N/A'}</td>
                <td>{data['staleness_days']}</td>
            </tr>
        """
    
    # Currency data row
    if quality_report['currency_data']:
        currency_data = quality_report['currency_data']
        html_content += f"""
            <tr>
                <td>Currency Rates</td>
                <td>{currency_data['total_records']}</td>
                <td>N/A</td>
                <td>N/A</td>
                <td>{currency_data['latest_date'] or 'N/A'}</td>
                <td>{currency_data['staleness_days']}</td>
            </tr>
        """
    
    html_content += """
        </table>
    </div>
    
    <div class="footer">
        <p>This automated report is generated by the Equate portfolio analysis system.</p>
        <p>GitHub Actions workflow: <code>update-stock.yml</code></p>
        <p>For issues or questions, check the repository: <a href="https://github.com/PahliAi/Equate">Equate on GitHub</a></p>
    </div>
    
    </body>
    </html>
    """
    
    return html_content

def generate_success_summary(quality_report, operation_summary):
    """
    Generate a concise success summary for GitHub Actions output.
    This will be visible in the workflow logs and GitHub's email notifications.
    """
    health_emoji = "🟢" if quality_report['overall_health'] == 'HEALTHY' else "🟡" if quality_report['overall_health'] == 'WARNING' else "🔴"
    
    summary = f"""
📈 EQUATE DATA UPDATE SUCCESS REPORT - {datetime.now().strftime('%d %B %Y')}
{'='*60}

{health_emoji} SYSTEM STATUS: {quality_report['overall_health']}

📊 OPERATION SUMMARY:
   • New Stock Records: {operation_summary.get('new_stock_records', 0)}
   • Currency Updates: {operation_summary.get('currency_updates', 0)}  
   • Gaps Filled: {operation_summary.get('gaps_filled', 0)}
   • USD Conversions: {operation_summary.get('conversions', 0)}

"""

    # Add data quality issues if any
    if quality_report['data_gaps']:
        summary += "⚠️  DATA QUALITY ISSUES:\n"
        for gap in quality_report['data_gaps']:
            summary += f"   • {gap}\n"
        summary += "\n"

    # Last 5 days stock prices
    if quality_report['last_5_days_shares']:
        summary += "📈 STOCK PRICES (LAST 5 DAYS):\n"
        summary += "   Date       | Allianz  | IBM      | Unilever\n"
        summary += "   -----------|----------|----------|----------\n"
        for day in quality_report['last_5_days_shares'][-5:]:
            summary += f"   {day['Date']} | {day.get('Allianz Share', 'N/A'):>8} | {day.get('IBM Share', 'N/A'):>8} | {day.get('Unilever Share', 'N/A'):>8}\n"
        summary += "\n"

    # Last 5 days currency rates  
    if quality_report['last_5_days_currency']:
        summary += "💱 EXCHANGE RATES (LAST 5 DAYS, BASE: EUR):\n"
        summary += "   Date       | USD    | GBP    | CAD    | CNY\n"
        summary += "   -----------|--------|--------|--------|--------\n"
        for day in quality_report['last_5_days_currency'][-5:]:
            summary += f"   {day['Date']} | {day.get('USD', 'N/A'):>6} | {day.get('GBP', 'N/A'):>6} | {day.get('CAD', 'N/A'):>6} | {day.get('CNY', 'N/A'):>6}\n"
        summary += "\n"

    # Data coverage summary
    summary += "📊 DATA COVERAGE SUMMARY:\n"
    for stock_name, data in quality_report['shares_data'].items():
        staleness = f"{data['staleness_days']}d" if data['staleness_days'] < 999 else "N/A"
        summary += f"   • {stock_name}: {data['total_records']} records ({data['coverage_pct']}%), last: {data['last_date'] or 'N/A'} ({staleness})\n"
    
    if quality_report['currency_data']:
        currency_data = quality_report['currency_data']
        staleness = f"{currency_data['staleness_days']}d" if currency_data['staleness_days'] < 999 else "N/A"
        summary += f"   • Currency Rates: {currency_data['total_records']} records, last: {currency_data['latest_date'] or 'N/A'} ({staleness})\n"

    summary += f"\n✅ Update completed successfully at {datetime.now().strftime('%H:%M UTC')}"
    
    return summary

def fill_missing_stock_data(shares_df, currency_df, lookback_days=10):
    """
    Look back 'lookback_days' days from the most recent date and fill any gaps.
    Much simpler: one API call per stock for the entire period.
    Also handles USD to EUR conversion for newly filled data only.
    """
    if shares_df.empty:
        return shares_df, {'filled': 0, 'skipped': 0}
    
    # Find the most recent date in the dataframe
    shares_df['Date_parsed'] = pd.to_datetime(shares_df['Date'], format='%d-%m-%Y', errors='coerce')
    last_date = shares_df['Date_parsed'].max()
    start_date = last_date - timedelta(days=lookback_days)
    
    print(f"🔍 Checking for gaps from {start_date.strftime('%d-%m-%Y')} to {last_date.strftime('%d-%m-%Y')}")
    
    fill_stats = {'filled': 0, 'skipped': 0}
    filled_positions = []  # Track what we filled for USD conversion
    
    # For each stock, fetch data for the entire lookback period
    for stock_name, ticker in STOCKS.items():
        if stock_name not in shares_df.columns:
            continue
        
        try:
            print(f"   📊 Fetching {stock_name} ({ticker}) data for gap filling...")
            
            # Download data for the entire lookback period (one API call per stock)
            stock_data = yf.download(ticker, start=start_date, end=last_date + timedelta(days=1), auto_adjust=True)
            
            if stock_data.empty:
                print(f"   ⚠️ No data returned for {stock_name}")
                continue
                
            # Handle MultiIndex columns
            if isinstance(stock_data.columns, pd.MultiIndex):
                stock_data.columns = stock_data.columns.droplevel(1)
            
            # Check each date in our dataframe within the lookback period
            recent_period_mask = shares_df['Date_parsed'] >= start_date
            recent_records = shares_df[recent_period_mask]
            
            gaps_filled_for_stock = 0
            
            for idx, row in recent_records.iterrows():
                # Only fill if there's a gap (NaN value)
                if pd.notna(row[stock_name]):
                    continue
                    
                date_str = row['Date']
                target_date = row['Date_parsed'].strftime('%Y-%m-%d')
                
                # Check if we have data for this exact date
                if target_date in stock_data.index.strftime('%Y-%m-%d'):
                    close_price = round(float(stock_data.loc[target_date]['Close']), 2)
                    shares_df.loc[idx, stock_name] = close_price
                    fill_stats['filled'] += 1
                    gaps_filled_for_stock += 1
                    filled_positions.append((idx, stock_name, date_str, close_price))
                    print(f"   ✅ Filled {stock_name} for {date_str} = {close_price}")
                else:
                    # Find the most recent available date before the gap
                    available_dates = stock_data.index[stock_data.index <= row['Date_parsed']]
                    if not available_dates.empty:
                        most_recent_date = available_dates.max()
                        close_price = round(float(stock_data.loc[most_recent_date]['Close']), 2)
                        shares_df.loc[idx, stock_name] = close_price
                        fill_stats['filled'] += 1
                        gaps_filled_for_stock += 1
                        filled_positions.append((idx, stock_name, date_str, close_price))
                        print(f"   ✅ Filled {stock_name} for {date_str} = {close_price} (from {most_recent_date.strftime('%d-%m-%Y')})")
                    else:
                        fill_stats['skipped'] += 1
                        print(f"   ⚠️ No data available for {stock_name} around {date_str}")
            
            if gaps_filled_for_stock == 0:
                print(f"   ✅ No gaps found in {stock_name}")
                        
        except Exception as e:
            fill_stats['skipped'] += 1
            print(f"   ❌ Error processing {stock_name}: {e}")
    
    # Convert USD stocks to EUR for newly filled positions only
    if filled_positions:
        conversion_count = 0
        print("💱 Converting newly filled USD prices to EUR...")
        for idx, stock_name, date_str, original_price in filled_positions:
            if stock_name in USD_STOCKS:
                usd_eur_rate = get_usd_to_eur_rate(currency_df, date_str)
                if usd_eur_rate is not None:
                    eur_price = convert_usd_to_eur(original_price, usd_eur_rate)
                    shares_df.loc[idx, stock_name] = eur_price
                    conversion_count += 1
                    print(f"   💱 Converted {stock_name} for {date_str}: ${original_price} → €{eur_price}")
        
        if conversion_count > 0:
            print(f"   ✅ Converted {conversion_count} newly filled USD prices to EUR")
    
    # Clean up temporary column
    shares_df = shares_df.drop('Date_parsed', axis=1)
    
    print(f"📊 Gap filling complete: {fill_stats['filled']} filled, {fill_stats['skipped']} skipped")
    return shares_df, fill_stats

def fill_missing_currency_data(currency_df, lookback_days=10):
    """
    Look back 'lookback_days' days from the most recent date and fill any currency gaps.
    Uses yfinance to get ACTUAL HISTORICAL exchange rates for each specific date.
    """
    if currency_df.empty:
        return currency_df, {'filled': 0, 'skipped': 0}
    
    # Find the most recent date in the dataframe
    currency_df['Date_parsed'] = pd.to_datetime(currency_df['Date'], format='%d-%m-%Y', errors='coerce')
    last_date = currency_df['Date_parsed'].max()
    start_date = last_date - timedelta(days=lookback_days)
    
    print(f"🔍 Checking for currency gaps from {start_date.strftime('%d-%m-%Y')} to {last_date.strftime('%d-%m-%Y')}")
    
    fill_stats = {'filled': 0, 'skipped': 0}
    existing_currencies = [col for col in currency_df.columns if col != 'Date' and col != 'Date_parsed']
    
    
    try:
        # Find all missing dates: 10 days before last_date AND all weekdays after last_date
        missing_dates = []
        
        # 1. Check 10 days BEFORE last date
        current_date = start_date.date()
        end_date = last_date.date()
        
        while current_date <= end_date:
            # Skip weekends
            if current_date.weekday() < 5:  # Monday=0, Friday=4
                date_str = current_date.strftime('%d-%m-%Y')
                has_date = any(currency_df['Date_parsed'].dt.strftime('%d-%m-%Y') == date_str)
                if not has_date:
                    missing_dates.append(current_date)
            current_date += timedelta(days=1)
        
        # 2. Check ALL weekdays AFTER last date up to today
        current_date = (last_date + timedelta(days=1)).date()
        today_date = datetime.now().date()
        
        while current_date <= today_date:
            # Skip weekends
            if current_date.weekday() < 5:  # Monday=0, Friday=4
                missing_dates.append(current_date)
            current_date += timedelta(days=1)
        
        if not missing_dates:
            print(f"   ✅ No currency gaps found")
            return currency_df.drop('Date_parsed', axis=1), fill_stats
        
        # OPTIMIZED: Download data for all missing dates at once for key currencies only
        if missing_dates:
            print(f"   💱 Getting historical rates for {len(missing_dates)} missing dates...")
            
            # Get date range for efficient batch download
            start_date = min(missing_dates)
            end_date = max(missing_dates) + timedelta(days=1)
            
            # Get ALL currencies from existing data (all will be tried with yfinance)
            available_currencies = existing_currencies
            
            # Batch download for ALL available currencies (not just 5)
            forex_data_cache = {}
            for currency_code in available_currencies:  # Fill ALL currencies, not just 5
                try:
                    ticker = get_yfinance_ticker(currency_code)
                    print(f"     📈 Downloading {currency_code} data...")
                    forex_data = yf.download(ticker, start=start_date, end=end_date, auto_adjust=True)
                    
                    if not forex_data.empty:
                        # Handle MultiIndex columns
                        if isinstance(forex_data.columns, pd.MultiIndex):
                            forex_data.columns = forex_data.columns.droplevel(1)
                        forex_data_cache[currency_code] = forex_data
                        print(f"     ✅ Got {len(forex_data)} days of {currency_code} data")
                    
                except Exception as e:
                    print(f"     ❌ Error downloading {currency_code}: {e}")
                    continue
            
            # Now fill missing dates using cached data
            for missing_date in missing_dates:
                date_str = missing_date.strftime('%d-%m-%Y')
                new_row = {'Date': date_str}
                currencies_filled = 0
                
                # Get rates from cached data
                for currency_code, forex_data in forex_data_cache.items():
                    try:
                        date_key = missing_date.strftime('%Y-%m-%d')
                        if date_key in forex_data.index.strftime('%Y-%m-%d'):
                            rate = round(float(forex_data.loc[date_key]['Close']), 4)
                            new_row[currency_code] = rate
                            currencies_filled += 1
                    except Exception:
                        continue
                
                # Only add the row if we got some currency data
                if currencies_filled > 0:
                    new_row_df = pd.DataFrame([new_row])
                    currency_df = pd.concat([currency_df, new_row_df], ignore_index=True)
                    fill_stats['filled'] += 1
                    print(f"   ✅ Filled {currencies_filled} currencies for {date_str}")
                else:
                    fill_stats['skipped'] += 1
                    print(f"   ⚠️ No currency data available for {date_str}")
        
    except Exception as e:
        print(f"   ❌ Error in currency gap filling: {e}")
        fill_stats['skipped'] += len(missing_dates) if 'missing_dates' in locals() else 1
    
    # Clean up temporary column and sort by date, ensuring consistent format
    currency_df = currency_df.drop('Date_parsed', axis=1)
    
    # Parse all dates to datetime and convert back to consistent dd-mm-yyyy format
    currency_df['Date_parsed'] = pd.to_datetime(currency_df['Date'], errors='coerce')
    currency_df = currency_df.sort_values('Date_parsed').reset_index(drop=True)
    currency_df['Date'] = currency_df['Date_parsed'].dt.strftime('%d-%m-%Y')
    currency_df = currency_df.drop('Date_parsed', axis=1)
    
    print(f"💱 Currency gap filling complete: {fill_stats['filled']} filled, {fill_stats['skipped']} skipped")
    return currency_df, fill_stats

def generate_company_specific_files():
    """
    Generate hist_Company.xlsx files with 30-currency Share sheet
    Each company gets a separate file with pre-computed prices in all major currencies
    Returns: dict with 'companies_found', 'files_created', 'files_failed'
    """
    print("\n📊 Generating company-specific multi-currency files...")
    
    company_stats = {'companies_found': 0, 'files_created': 0, 'files_failed': 0}
    
    try:
        # Load base data with Row 1 as headers (Company names)
        base_shares_eur = pd.read_excel('hist_base.xlsx', sheet_name='Shares', header=1)
        currency_ratios = pd.read_excel('hist_base.xlsx', sheet_name='Currency')
        
        # Now columns are: ['Company', 'Allianz', 'IBM', 'Unilever'] from Row 1
        # Get company names (skip first column 'Company')
        companies = [col for col in base_shares_eur.columns if col != 'Company']
        company_stats['companies_found'] = len(companies)
        
        print(f"   📋 Found {len(companies)} companies: {companies}")
        
        # Skip ticker row to get actual data (Row 3+ in original Excel)
        data_rows = base_shares_eur.iloc[1:].reset_index(drop=True)
        print(f"   📊 Data rows available: {len(data_rows)}")
        
        # Get available currencies from Currency sheet (ONLY existing ones + EUR)
        available_currencies = [col for col in currency_ratios.columns if col != 'Date']
        currencies = ['EUR'] + available_currencies  # EUR first, then existing currencies
        print(f"   💱 Target currencies: {currencies[:5]}... ({len(currencies)} total)")
        print(f"   💱 Using ONLY currencies from hist_base.xlsx Currency sheet")
        
        # Generate file for each company
        for company in companies:
            company_file = f"hist_{company}.xlsx"
            
            print(f"   🏢 Processing {company} → {company_file}")
            
            # Create multi-currency data frame
            multi_currency_data = pd.DataFrame()
            
            # Use dates from the data rows (skip ticker row)
            multi_currency_data['Date'] = data_rows['Company']  # 'Company' column has dates after header=1
            
            # Get base EUR prices for this company
            if company in data_rows.columns:
                base_eur_prices = data_rows[company]
                print(f"   📈 Using column '{company}' with {len(base_eur_prices)} price entries")
            else:
                print(f"   ❌ Column '{company}' not found in base data")
                available_cols = list(data_rows.columns)
                print(f"   📋 Available columns: {available_cols}")
                company_stats['files_failed'] += 1
                continue
            
            # Start with EUR as base currency - ensure numeric type
            multi_currency_data['EUR'] = pd.to_numeric(base_eur_prices, errors='coerce')
            
            # Use efficient pandas merge for currency conversions
            # Merge on Date to align share prices with currency rates
            merged_data = multi_currency_data.merge(currency_ratios, on='Date', how='left')
            
            # Generate prices for all other currencies using vectorized operations
            for currency in available_currencies:
                if currency in merged_data.columns:
                    # Convert currency column to numeric, handling any non-numeric values
                    merged_data[currency] = pd.to_numeric(merged_data[currency], errors='coerce')
                    
                    # Vectorized conversion: EUR_price × currency_rate = target_price
                    # Only where both values exist (non-null)
                    valid_mask = (merged_data['EUR'].notna() & 
                                 merged_data[currency].notna() & 
                                 (merged_data[currency] > 0))
                    
                    multi_currency_data[currency] = None  # Initialize with None
                    if valid_mask.sum() > 0:  # Only proceed if we have valid data
                        converted_values = (merged_data.loc[valid_mask, 'EUR'] * 
                                          merged_data.loc[valid_mask, currency]).round(2)
                        multi_currency_data.loc[valid_mask, currency] = converted_values
            
            # Write company-specific file
            try:
                with pd.ExcelWriter(company_file, engine='openpyxl') as writer:
                    multi_currency_data.to_excel(writer, sheet_name='Share', index=False)
                
                # Count valid data
                valid_dates = multi_currency_data['Date'].count()
                valid_eur = multi_currency_data['EUR'].count()
                
                company_stats['files_created'] += 1
                print(f"   ✅ Generated {company_file}")
                print(f"      📅 {valid_dates} dates, {valid_eur} EUR prices")
                print(f"      💱 {len(currencies)} currencies pre-computed")
            except Exception as file_error:
                company_stats['files_failed'] += 1
                print(f"   ❌ Failed to create {company_file}: {file_error}")
    
    except Exception as e:
        print(f"   ❌ Error generating company files: {e}")
        import traceback
        traceback.print_exc()
        company_stats['files_failed'] = company_stats['companies_found']  # Mark all as failed
    
    print(f"   📊 Company file generation complete: {company_stats['files_created']} created, {company_stats['files_failed']} failed")
    return company_stats

def main():
    print("📈 Multi-Stock Daily Update Script")
    print("=" * 50)
    
    # 1. Bestaande data inladen
    shares_df_existing = pd.DataFrame()
    currency_df_existing = pd.DataFrame()

    if os.path.exists(EXCEL_FILE):
        try:
            with pd.ExcelFile(EXCEL_FILE) as xl:
                if 'Shares' in xl.sheet_names:
                    shares_df_existing = pd.read_excel(EXCEL_FILE, sheet_name='Shares')
                    
                    # CRITICAL FIX: Filter out metadata rows BEFORE any processing
                    if not shares_df_existing.empty and 'Date' in shares_df_existing.columns:
                        # Parse dates and filter out NaT (metadata rows like "Ticker")
                        shares_df_existing['Date_temp'] = pd.to_datetime(shares_df_existing['Date'], format='%d-%m-%Y', errors='coerce')
                        original_count = len(shares_df_existing)
                        shares_df_existing = shares_df_existing[shares_df_existing['Date_temp'].notna()].copy()
                        shares_df_existing = shares_df_existing.drop('Date_temp', axis=1)
                        filtered_count = len(shares_df_existing)
                        if original_count > filtered_count:
                            print(f"📊 Filtered out {original_count - filtered_count} metadata rows (Ticker, etc.)")
                    
                    # Ensure all stock columns are numeric
                    for stock_name in STOCKS.keys():
                        if stock_name in shares_df_existing.columns:
                            shares_df_existing[stock_name] = pd.to_numeric(shares_df_existing[stock_name], errors='coerce')
                    print(f"📊 Bestaande shares records: {len(shares_df_existing)}")
                    
                if 'Currency' in xl.sheet_names:
                    currency_df_existing = pd.read_excel(EXCEL_FILE, sheet_name='Currency')
                    print(f"💱 Bestaande currency records: {len(currency_df_existing)}")
        except Exception as e:
            print(f"⚠️ Fout bij lezen bestaand bestand: {e}")

    if shares_df_existing.empty:
        shares_df_existing = pd.DataFrame(columns=["Date"] + list(STOCKS.keys()))
    if currency_df_existing.empty:
        currency_df_existing = pd.DataFrame(columns=["Date"])

    # 2. Laatste datum bepalen
    last_date_shares = None
    last_date_currency = None

    if not shares_df_existing.empty:
        shares_df_existing['Date'] = pd.to_datetime(shares_df_existing['Date'], format='%d-%m-%Y', errors='coerce')
        last_date_shares = shares_df_existing['Date'].max()

    if not currency_df_existing.empty:
        currency_df_existing['Date'] = pd.to_datetime(currency_df_existing['Date'], format='%d-%m-%Y', errors='coerce')
        last_date_currency = currency_df_existing['Date'].max()

    # Use the earliest date to ensure both datasets are updated
    last_date = min(filter(None, [last_date_shares, last_date_currency])) if any([last_date_shares, last_date_currency]) else datetime.now() - timedelta(days=DAYS_THRESHOLD + 1)

    print(f"📅 Laatste datum in bestand: {last_date.strftime('%d-%m-%Y') if last_date else 'Geen data'}")

    # 2.5. Fill missing stock data gaps (look back up to 10 days from most recent date)
    gap_stats = {'filled': 0, 'skipped': 0}
    currency_gap_stats = {'filled': 0, 'skipped': 0}
    
    if not shares_df_existing.empty:
        shares_df_existing, gap_stats = fill_missing_stock_data(shares_df_existing, currency_df_existing, lookback_days=10)
    
    # 2.6. Fill missing currency data gaps (look back up to 10 days from most recent date)
    if not currency_df_existing.empty:
        currency_df_existing, currency_gap_stats = fill_missing_currency_data(currency_df_existing, lookback_days=10)
        
    if gap_stats['filled'] > 0 or currency_gap_stats['filled'] > 0:
            # Save the updated data immediately after gap filling
            print("💾 Saving gap-filled data...")
            shares_df_save = shares_df_existing.copy()
            shares_df_save['Date'] = shares_df_save['Date'].dt.strftime('%d-%m-%Y') if hasattr(shares_df_save['Date'], 'dt') else shares_df_save['Date'].astype(str)
            
            currency_df_save = currency_df_existing.copy()
            if not currency_df_save.empty:
                currency_df_save['Date'] = currency_df_save['Date'].dt.strftime('%d-%m-%Y') if hasattr(currency_df_save['Date'], 'dt') else currency_df_save['Date'].astype(str)
            
            # Note: USD to EUR conversion will be handled later in the main flow
            # to avoid double-converting existing data
            
            # Save to Excel with metadata preservation
            write_excel_with_metadata(shares_df_save, currency_df_save, EXCEL_FILE)
            
            print(f"✅ Gap-filled data saved to {EXCEL_FILE}")

    # 3. Check of update nodig is
    if (datetime.now() - last_date).days < DAYS_THRESHOLD:
        print(f"✅ Data is al up-to-date (laatste: {last_date.strftime('%d-%m-%Y')}) — geen API-call.")
        
        # Convert dates back to proper format for analysis
        shares_df_for_analysis = shares_df_existing.copy()
        currency_df_for_analysis = currency_df_existing.copy()
        
        if not shares_df_existing.empty:
            shares_df_for_analysis['Date'] = shares_df_existing['Date'].dt.strftime('%d-%m-%Y')
        if not currency_df_existing.empty:
            currency_df_for_analysis['Date'] = currency_df_existing['Date'].dt.strftime('%d-%m-%Y')
        
        # Generate company files even when data is current
        print(f"\n🏢 Starting company-specific file generation...")
        company_file_stats = generate_company_specific_files()
        
        # Use centralized reporting function (no update scenario) - mark as "data current"
        generate_final_report_and_exit(shares_df_for_analysis, currency_df_for_analysis, 0, 0, gap_stats.get('filled', 0) + currency_gap_stats.get('filled', 0), 0, company_file_stats, data_was_current=True)

    # 4. Nieuwe stock data ophalen voor alle stocks
    print(f"📈 Nieuwe stock data ophalen vanaf {(last_date + timedelta(days=1)).strftime('%d-%m-%Y')}...")
    
    new_stock_records = []
    total_new_records = 0
    
    for stock_name, ticker in STOCKS.items():
        try:
            print(f"   📊 {stock_name} ({ticker})...")
            stock_data = yf.download(ticker, start=last_date + timedelta(days=1), auto_adjust=True)
            
            if not stock_data.empty:
                # Handle MultiIndex columns from yfinance
                if isinstance(stock_data.columns, pd.MultiIndex):
                    stock_data.columns = stock_data.columns.droplevel(1)
                
                # Process each date
                for date, row in stock_data.iterrows():
                    date_str = date.strftime('%d-%m-%Y')
                    close_price = round(float(row['Close']), 2)
                    
                    # Find or create record for this date
                    existing_record = None
                    for record in new_stock_records:
                        if record['Date'] == date_str:
                            existing_record = record
                            break
                    
                    if existing_record is None:
                        existing_record = {'Date': date_str}
                        new_stock_records.append(existing_record)
                    
                    existing_record[stock_name] = close_price
                
                print(f"   ✅ {len(stock_data)} records gevonden")
                total_new_records = max(total_new_records, len(stock_data))
            else:
                print(f"   ⚠️ Geen data voor {stock_name}")
                
        except Exception as e:
            print(f"   ❌ Fout bij {stock_name}: {e}")

    # 5. Nieuwe currency data ophalen (backfill missing days)
    print(f"💱 Currency data ophalen...")
    new_currency_data = pd.DataFrame()
    currency_added_count = 0
    
    try:
        response = requests.get(CURRENCY_API_URL, timeout=10)
        currency_data = response.json()
        currency_rates = currency_data['rates']
        
        # Get existing currencies from file structure
        existing_currencies = [col for col in currency_df_existing.columns if col != 'Date']
        
        # Only add TODAY's currency data using yfinance for consistent 4-decimal precision
        today_str = datetime.now().strftime('%d-%m-%Y')
        today_weekday = datetime.now().weekday()
        today_date = datetime.now().date()
        
        # Only add if today is a weekday and we don't already have today's data
        if today_weekday < 5:  # Monday=0, Friday=4
            # Check if we already have today's data (handle both datetime and string formats)
            if hasattr(currency_df_existing['Date'], 'dt'):
                has_today = not currency_df_existing.empty and any(
                    currency_df_existing['Date'].dt.strftime('%d-%m-%Y') == today_str
                )
            else:
                has_today = not currency_df_existing.empty and any(
                    currency_df_existing['Date'].astype(str) == today_str
                )
            
            if not has_today:
                print(f"   💱 Getting TODAY's currency data from yfinance for consistency: {today_str}")
                currency_row = {'Date': today_str}
                currencies_added = 0
                
                # Get today's rates using yfinance for 4-decimal precision
                for currency_code in existing_currencies:
                    try:
                        ticker = get_yfinance_ticker(currency_code)
                        forex_data = yf.download(ticker, start=today_date, end=today_date + timedelta(days=1), auto_adjust=True)
                        
                        if not forex_data.empty:
                            if isinstance(forex_data.columns, pd.MultiIndex):
                                forex_data.columns = forex_data.columns.droplevel(1)
                            
                            rate = round(float(forex_data['Close'].iloc[0]), 4)
                            currency_row[currency_code] = rate
                            currencies_added += 1
                    except Exception:
                        # Fallback to API rate if yfinance fails
                        if currency_code in currency_rates:
                            rate = round(float(currency_rates[currency_code]), 4)
                            currency_row[currency_code] = rate
                            currencies_added += 1
                
                if currencies_added > 0:
                    currency_rows = [currency_row]
                    currency_added_count = 1
                    print(f"   ✅ Added {currencies_added} currencies for TODAY with 4-decimal precision")
                else:
                    currency_rows = []
                    currency_added_count = 0
                    print(f"   ❌ Could not get today's currency data")
            else:
                print(f"   💱 Today's currency data already exists: {today_str}")
        else:
            print(f"   💱 Today is weekend, no currency data to add")
        
        if 'currency_rows' in locals() and currency_rows:
            new_currency_data = pd.DataFrame(currency_rows)
            print(f"💱 Currency data toegevoegd voor {currency_added_count} dagen ({len(existing_currencies)} currencies each)")
        else:
            print(f"💱 Currency data is up to date")
            
    except Exception as e:
        print(f"⚠️ Fout bij ophalen currency data: {e}")

    # 6. Data samenvoegen
    print(f"📊 {len(new_stock_records)} nieuwe stock records + {len(new_currency_data)} currency updates gevonden")

    if not new_stock_records and new_currency_data.empty:
        print("⚠️ Geen nieuwe data beschikbaar.")
        return

    # Convert new stock records to DataFrame
    new_shares_data = pd.DataFrame(new_stock_records)
    
    # Shares data merging
    if not shares_df_existing.empty:
        shares_df_existing['Date'] = shares_df_existing['Date'].dt.strftime('%d-%m-%Y')

    shares_df_updated = pd.concat([shares_df_existing, new_shares_data], ignore_index=True)
    shares_df_updated = shares_df_updated.drop_duplicates(subset=["Date"], keep="last")

    # Currency data merging
    if not currency_df_existing.empty:
        # Only convert if Date column is still datetime (gap filling may have converted to strings)
        if hasattr(currency_df_existing['Date'], 'dt'):
            currency_df_existing['Date'] = currency_df_existing['Date'].dt.strftime('%d-%m-%Y')
        else:
            currency_df_existing['Date'] = currency_df_existing['Date'].astype(str)

    currency_df_updated = pd.concat([currency_df_existing, new_currency_data], ignore_index=True)
    currency_df_updated = currency_df_updated.drop_duplicates(subset=["Date"], keep="last")
    
    # 6.3. NOW fill any gaps between the oldest date and today with REAL historical rates
    print("💱 Checking for currency gaps to fill with historical rates...")
    currency_df_updated, final_currency_gap_stats = fill_missing_currency_data(currency_df_updated, lookback_days=10)
    
    if final_currency_gap_stats['filled'] > 0:
        print(f"💱 Filled {final_currency_gap_stats['filled']} currency gaps with historical rates")
        # Update total gaps filled
        currency_gap_stats['filled'] += final_currency_gap_stats['filled']
        currency_gap_stats['skipped'] += final_currency_gap_stats['skipped']

    # 6.5. Convert USD stock prices to EUR ONLY for newly fetched data
    print("💱 Converting newly fetched USD stock prices to EUR...")
    conversion_stats = {'converted': 0, 'skipped_no_rate': 0, 'skipped_missing_price': 0}
    
    # CRITICAL FIX: Only convert NEW data from new_stock_records, not existing EUR data
    for new_record in new_stock_records:
        date_str = new_record['Date']
        
        # Get the USD to EUR rate for this date
        usd_eur_rate = get_usd_to_eur_rate(currency_df_updated, date_str)
        
        # Convert each USD stock price to EUR for THIS new record only
        for stock_name in USD_STOCKS:
            if stock_name in new_record and pd.notna(new_record[stock_name]):
                if usd_eur_rate is not None:
                    # Convert the USD price in the new record
                    eur_price = convert_usd_to_eur(new_record[stock_name], usd_eur_rate)
                    new_record[stock_name] = eur_price
                    conversion_stats['converted'] += 1
                else:
                    conversion_stats['skipped_no_rate'] += 1
            elif stock_name in new_record:
                conversion_stats['skipped_missing_price'] += 1
    
    # Rebuild shares_df_updated with converted new records
    if new_stock_records:
        new_shares_data = pd.DataFrame(new_stock_records)
        shares_df_updated = pd.concat([shares_df_existing, new_shares_data], ignore_index=True)
        shares_df_updated = shares_df_updated.drop_duplicates(subset=["Date"], keep="last")
    
    if conversion_stats['converted'] > 0:
        print(f"   ✅ Converted {conversion_stats['converted']} USD prices to EUR")
    if conversion_stats['skipped_no_rate'] > 0:
        print(f"   ⚠️ Skipped {conversion_stats['skipped_no_rate']} prices (no exchange rate available)")
    if conversion_stats['skipped_missing_price'] > 0:
        print(f"   ⚠️ Skipped {conversion_stats['skipped_missing_price']} missing prices")

    # 7. Opslaan in Excel - direct approach
    # Ensure all data is clean before writing
    shares_df_updated = shares_df_updated.copy()
    currency_df_updated = currency_df_updated.copy()

    # Make sure Date columns are strings and numeric columns are proper floats
    shares_df_updated['Date'] = shares_df_updated['Date'].astype(str)
    for col in shares_df_updated.columns:
        if col != 'Date':
            shares_df_updated[col] = pd.to_numeric(shares_df_updated[col], errors='coerce')

    currency_df_updated['Date'] = currency_df_updated['Date'].astype(str)
    for col in currency_df_updated.columns:
        if col != 'Date':
            currency_df_updated[col] = pd.to_numeric(currency_df_updated[col], errors='coerce')

    # Write to Excel with metadata preservation
    write_excel_with_metadata(shares_df_updated, currency_df_updated, EXCEL_FILE)

    print(f"💾 Data opgeslagen in {EXCEL_FILE}:")
    print(f"   📈 Shares: {len(shares_df_updated)} records")
    print(f"   💱 Currency: {len(currency_df_updated)} records")
    
    # Show stock coverage
    for stock_name in STOCKS.keys():
        if stock_name in shares_df_updated.columns:
            non_null = shares_df_updated[stock_name].notna().sum()
            coverage = (non_null / len(shares_df_updated)) * 100
            print(f"   📊 {stock_name}: {non_null} records ({coverage:.1f}% coverage)")

    # 8. Generate company-specific multi-currency files (Sprint 2)
    print(f"\n🏢 Starting company-specific file generation...")
    company_file_stats = generate_company_specific_files()
    
    # 9. Data quality analysis and success email  
    generate_final_report_and_exit(shares_df_updated, currency_df_updated, len(new_stock_records), len(new_currency_data), gap_stats.get('filled', 0) + currency_gap_stats.get('filled', 0), conversion_stats.get('converted', 0), company_file_stats, data_was_current=False)

def generate_final_report_and_exit(shares_df, currency_df, new_records, currency_updates, gaps_filled, conversions, company_file_stats=None, data_was_current=False):
    """
    Generate final data quality report and exit with failure for email notification.
    Centralized function to avoid code duplication.
    """
    print("\n📊 Analyzing data quality and preparing success report...")
    
    # Prepare operation summary
    operation_summary = {
        'new_stock_records': new_records,
        'currency_updates': currency_updates,
        'gaps_filled': gaps_filled,
        'conversions': conversions
    }
    
    # Analyze data quality for last 5 days
    quality_report = analyze_data_quality(shares_df, currency_df, days=5)
    
    print(f"   📊 System health: {quality_report['overall_health']}")
    print(f"   📈 Stock data coverage: {len(quality_report['shares_data'])} stocks")
    print(f"   💱 Currency data: {quality_report['currency_data']['total_records'] if quality_report['currency_data'] else 0} records")
    
    if quality_report['data_gaps']:
        print("   ⚠️ Data quality issues detected:")
        for gap in quality_report['data_gaps']:
            print(f"      - {gap}")
    
    # Generate and display success summary
    success_summary = generate_success_summary(quality_report, operation_summary)
    print(success_summary)
    
    # Check for functional failures
    functional_failures = []
    
    # Functional failure logic depends on context:
    # - If data was current: only check company files
    # - If updates were attempted: check shares, currencies, AND company files
    
    if not data_was_current:
        # Updates were attempted - check if we got any results
        if new_records == 0:
            functional_failures.append("No share price updates obtained despite attempting to fetch new data")
        if currency_updates == 0 and gaps_filled == 0:
            functional_failures.append("No currency updates obtained despite attempting to fetch new data")
    
    # Always check company file generation (regardless of whether data was current)
    if company_file_stats:
        if company_file_stats.get('files_failed', 0) > 0:
            failed_count = company_file_stats.get('files_failed', 0)
            total_count = company_file_stats.get('companies_found', 0)
            functional_failures.append(f"Company file generation failures: {failed_count} out of {total_count} files failed")
        
        if company_file_stats.get('files_created', 0) == 0:
            functional_failures.append("No company-specific hist files were created")
    else:
        functional_failures.append("Company file statistics not available - function may have failed")
    
    # Determine exit code based on functional success
    if functional_failures:
        print(f"\n❌ FUNCTIONAL FAILURE DETECTED!")
        print(f"   While script executed without technical errors, core business objectives failed:")
        for failure in functional_failures:
            print(f"   • {failure}")
        
        # Output to GitHub Actions
        if os.getenv('GITHUB_ACTIONS'):
            print(f"::error title=Functional Failure::Core business objectives not met")
            for failure in functional_failures:
                print(f"::error title=Business Failure::{failure}")
            
            print("\n" + "="*60)
            print("📧 FUNCTIONAL FAILURE REPORT")
            print("="*60)
            print("Script executed without technical errors but failed to meet business objectives.")
            print("This requires immediate attention to fix data pipeline issues.")
            print("=" * 60)
        
        # Exit with failure code for functional failures
        exit(1)
    else:
        print(f"\n✅ Data processing completed successfully!")
        print(f"   🏥 Overall system health: {quality_report['overall_health']}")
        
        # Output summary to GitHub Actions (will be included in email notifications)
        if os.getenv('GITHUB_ACTIONS'):
            status_msg = "Data Already Current" if data_was_current else "Data Update Success"
            print(f"::notice title={status_msg}::Equate data processed successfully")
            if quality_report['overall_health'] != 'HEALTHY':
                print(f"::warning title=Data Quality Warning::System status is {quality_report['overall_health']}")
            if quality_report['data_gaps']:
                for gap in quality_report['data_gaps']:
                    print(f"::warning title=Data Gap::{gap}")
            
            print("\n" + "="*60)
            print("📧 SUCCESS REPORT GENERATED FOR EMAIL NOTIFICATION")
            print("="*60)
            print("The script will complete successfully and commit changes.")
            print("Data processing completed successfully.")
            print("Check the logs above for the complete data quality report.")
            print("=" * 60)
            
            # Success - no forced exit needed for git commits to work
            exit(0)

if __name__ == "__main__":
    main()