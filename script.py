import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False
    print("Warning: yfinance not installed. Stock price lookup will be unavailable.")
    print("Install with: pip install yfinance")

try:
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

def parse_percentage(value_str):
    """Parse percentage string (e.g., '30.9%') to float."""
    if pd.isna(value_str) or str(value_str).strip() == '':
        return 0

    value_str = str(value_str).strip()
    # Remove percentage sign if present
    value_str = value_str.rstrip('%')

    try:
        return float(value_str)
    except ValueError:
        print(f"Warning: Could not parse percentage value: {value_str}")
        return 0

def parse_date(date_str):
    """Parse date string in various formats."""
    if pd.isna(date_str) or str(date_str).strip() == '':
        return None

    date_str = str(date_str).strip()

    # Try different date formats
    date_formats = [
        '%d-%b-%Y',    # 19-NOV-2025
        '%m/%d/%Y',    # 11/19/2025
        '%d/%m/%Y',    # 19/11/2025 (if needed)
        '%Y-%m-%d',    # 2025-11-19
        '%b %d, %Y',   # Nov 19, 2025
    ]

    for fmt in date_formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue

    # If all formats fail, try to extract date parts
    try:
        # Handle cases like "11/19/2025 00:00:00"
        if ' ' in date_str:
            date_part = date_str.split(' ')[0]
            for fmt in ['%m/%d/%Y', '%Y-%m-%d']:
                try:
                    return datetime.strptime(date_part, fmt)
                except:
                    continue
    except:
        pass

    print(f"Warning: Could not parse date: {date_str}")
    return None

def get_financial_year(date_obj):
    """
    Get financial year (April 1 - March 31) for a given date.
    For example, April 1, 2025 - March 31, 2026 is FY 2025-26.
    """
    if date_obj is None:
        return None

    if date_obj.month >= 4:  # April onwards
        return f"FY{date_obj.year}-{date_obj.year + 1}"
    else:  # January to March
        return f"FY{date_obj.year - 1}-{date_obj.year}"

def get_capital_gains_tax_rate(grant_date, sale_date):
    """
    Calculate capital gains tax rate based on holding period.
    LTCG (Long-Term): 12.5% for holdings > 1 year
    STCG (Short-Term): 20% for holdings <= 1 year
    Returns (rate, tax_type)
    """
    if grant_date is None or sale_date is None:
        return None, None

    holding_period = sale_date - grant_date
    holding_days = holding_period.days

    # More than 365 days (1 year)
    if holding_days > 365:
        return 0.125, "LTCG"  # 12.5%
    else:
        return 0.20, "STCG"   # 20%

def get_exchange_rate(date_str):
    """Get historical USD to INR exchange rate for a given date."""
    if not YFINANCE_AVAILABLE:
        return None

    try:
        # Parse the date
        parsed_date = parse_date(date_str)
        if parsed_date is None:
            return None

        # Get USDINR exchange rate from yfinance
        ticker = yf.Ticker('USDINR=X')

        # Get historical data around the date
        start_date = (parsed_date - pd.Timedelta(days=5)).strftime('%Y-%m-%d')
        end_date = (parsed_date + pd.Timedelta(days=5)).strftime('%Y-%m-%d')

        hist = ticker.history(start=start_date, end=end_date)

        if len(hist) > 0:
            # Use closest date
            closest_date = hist.index[hist.index.get_indexer([parsed_date], method='nearest')[0]]
            return hist.loc[closest_date, 'Close']
    except Exception as e:
        pass  # Silently fail and return None

    return None

def get_stock_price(symbol, date_str):
    """Get historical stock price for a given symbol and date."""
    if not YFINANCE_AVAILABLE:
        return None

    try:
        # Parse the date
        parsed_date = parse_date(date_str)
        if parsed_date is None:
            return None

        # Add ticker to make it valid (e.g., PTC -> PTC)
        ticker = yf.Ticker(symbol)

        # Get historical data around the date
        start_date = (parsed_date - pd.Timedelta(days=5)).strftime('%Y-%m-%d')
        end_date = (parsed_date + pd.Timedelta(days=5)).strftime('%Y-%m-%d')

        hist = ticker.history(start=start_date, end=end_date)

        # Find the closest trading date to the actual date
        if parsed_date.strftime('%Y-%m-%d') in hist.index.strftime('%Y-%m-%d'):
            # Exact date match
            return hist.loc[parsed_date.strftime('%Y-%m-%d'), 'Close']
        elif len(hist) > 0:
            # Use closest date
            closest_date = hist.index[hist.index.get_indexer([parsed_date], method='nearest')[0]]
            return hist.loc[closest_date, 'Close']
    except Exception as e:
        print(f"Could not fetch price for {symbol} on {date_str}: {str(e)}")
        return None

def process_restricted_stock(df, symbol_for_price='PTC', grant_type='RSU'):
    """
    Process Restricted Stock data and return grants dictionary.

    Parameters:
    -----------
    df : DataFrame
        Input data
    symbol_for_price : str
        Stock ticker symbol for historical price lookup (default: 'PTC')
    grant_type : str
        Type of grant ('RSU' or 'ESPP')

    Returns:
    --------
    dict : Dictionary of processed grants
    """
    # Standardize column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Remove completely empty rows
    df = df.dropna(how='all')

    # Reset index for easier processing
    df = df.reset_index(drop=True)

    # Dictionary to store grant information
    grants = {}
    current_grant = None
    grant_counter = 0

    # Process each row
    for idx, row in df.iterrows():
        record_type = str(row['Record Type']).strip() if pd.notna(row.get('Record Type')) else ''

        # Handle Grant records
        if record_type == 'Grant':
            grant_counter += 1
            symbol = str(row['Symbol']).strip() if pd.notna(row.get('Symbol')) else ''
            grant_date_str = str(row['Grant Date']).strip() if pd.notna(row.get('Grant Date')) else ''

            # Create unique grant ID - use Grant Number if available, otherwise use date + counter
            grant_number = str(row.get('Grant Number', '')).strip() if pd.notna(row.get('Grant Number')) else ''
            if grant_number:
                grant_id = grant_number
            else:
                grant_id = f"{grant_date_str}_{grant_counter}"

            # Parse grant date
            grant_date = parse_date(grant_date_str)

            # Get grant date stock price for capital gains calculation
            grant_price = None
            if YFINANCE_AVAILABLE and symbol:
                grant_price = get_stock_price(symbol, grant_date_str)

            # Initialize grant dictionary
            current_grant = {
                'grant_id': grant_id,
                'grant_type': grant_type,
                'symbol': symbol,
                'grant_date': grant_date,
                'grant_date_str': grant_date_str,
                'grant_price': grant_price,
                'granted_qty': float(row['Granted Qty.']) if pd.notna(row.get('Granted Qty.')) else 0,
                'withheld_qty': float(row['Withheld Qty.']) if pd.notna(row.get('Withheld Qty.')) else 0,
                'vested_qty': float(row['Vested Qty.']) if pd.notna(row.get('Vested Qty.')) else 0,
                'sellable_qty': float(row['Sellable Qty.']) if pd.notna(row.get('Sellable Qty.')) else 0,
                'unvested_qty': float(row['Unvested Qty.']) if pd.notna(row.get('Unvested Qty.')) else 0,
                'released_qty': float(row['Released Qty']) if pd.notna(row.get('Released Qty')) else 0,
                'est_market_value': float(row['Est. Market Value']) if pd.notna(row.get('Est. Market Value')) else 0,
                'events': [],  # List of events (vest, release, sell)
                'vest_schedules': [],  # List of vest schedules
                'tax_withholdings': [],  # List of tax withholdings
                'sales': [],  # List of sales
                'capital_gains_tax': [],  # List of capital gains taxes
                'total_tax_withheld': 0,
                'total_capital_gains_tax': 0,
                'total_sold_qty': 0,
                'total_sale_proceeds': 0,
                'sale_dates': [],
                'validation_issues': []
            }

            grants[grant_id] = current_grant

        # Handle Event records (grant, vest, release, sell)
        elif record_type == 'Event' and current_grant is not None:
            event_date_str = str(row['Date']).strip() if pd.notna(row.get('Date')) else ''
            event_type = str(row['Event Type']).strip() if pd.notna(row.get('Event Type')) else ''
            qty_or_amount = float(row['Qty. or Amount']) if pd.notna(row.get('Qty. or Amount')) else 0

            event_date = parse_date(event_date_str)

            event_info = {
                'date': event_date,
                'date_str': event_date_str,
                'type': event_type,
                'quantity': qty_or_amount
            }

            current_grant['events'].append(event_info)

            # Track sales separately
            if 'sold' in event_type.lower():
                sale_price = float(row['Sale Price']) if pd.notna(row.get('Sale Price')) else None

                # Try to fetch historical stock price if not available
                if sale_price is None and YFINANCE_AVAILABLE:
                    sale_price = get_stock_price(symbol_for_price, event_date_str)

                # Get exchange rate on sale date
                exchange_rate = None
                if YFINANCE_AVAILABLE:
                    exchange_rate = get_exchange_rate(event_date_str)

                # Calculate capital gains tax based on holding period
                capital_gain = 0
                capital_gains_tax = 0
                tax_rate = 0
                tax_type = "N/A"

                if sale_price is not None and current_grant['grant_price'] is not None:
                    capital_gain = (sale_price - current_grant['grant_price']) * qty_or_amount

                    # Determine tax rate based on holding period
                    tax_rate, tax_type = get_capital_gains_tax_rate(current_grant['grant_date'], event_date)

                    if tax_rate is not None:
                        capital_gains_tax = capital_gain * tax_rate
                        current_grant['total_capital_gains_tax'] += capital_gains_tax

                        # Track capital gain tax separately
                        current_grant['capital_gains_tax'].append({
                            'date': event_date,
                            'date_str': event_date_str,
                            'grant_price': current_grant['grant_price'],
                            'sale_price': sale_price,
                            'quantity': qty_or_amount,
                            'capital_gain': capital_gain,
                            'holding_days': (event_date - current_grant['grant_date']).days,
                            'tax_type': tax_type,
                            'tax_rate': tax_rate,
                            'tax_amount': capital_gains_tax
                        })

                sale_info = {
                    'date': event_date,
                    'date_str': event_date_str,
                    'quantity': qty_or_amount,
                    'price': sale_price,
                    'grant_price': current_grant['grant_price'],
                    'capital_gain': capital_gain,
                    'capital_gains_tax': capital_gains_tax,
                    'tax_type': tax_type,
                    'tax_rate': tax_rate,
                    'exchange_rate': exchange_rate
                }
                current_grant['sales'].append(sale_info)
                current_grant['total_sold_qty'] += qty_or_amount
                current_grant['sale_dates'].append(event_date_str)

        # Handle Vest Schedule records
        elif record_type == 'Vest Schedule' and current_grant is not None:
            vest_date_str = str(row['Vest Date']).strip() if pd.notna(row.get('Vest Date')) else ''
            vested_qty = float(row['Vested Qty.']) if pd.notna(row.get('Vested Qty.')) else 0
            released_qty = float(row['Released Qty']) if pd.notna(row.get('Released Qty')) else 0
            vest_period = str(row['Vest Period']).strip() if pd.notna(row.get('Vest Period')) else ''

            vest_date = parse_date(vest_date_str)

            vest_schedule = {
                'vest_date': vest_date,
                'vest_date_str': vest_date_str,
                'vested_qty': vested_qty,
                'released_qty': released_qty,
                'vest_period': vest_period,
                'is_future': vest_date > datetime.now() if vest_date else False
            }

            current_grant['vest_schedules'].append(vest_schedule)

        # Handle Tax Withholding records (only for RSU, not ESPP)
        elif record_type == 'Tax Withholding' and current_grant is not None and grant_type == 'RSU':
            withholding_date_str = str(row['Date']).strip() if pd.notna(row.get('Date')) else ''
            tax_rate = parse_percentage(row['Effective Tax Rate']) if pd.notna(row.get('Effective Tax Rate')) else 0
            withholding_amount = float(row['Withholding Amount']) if pd.notna(row.get('Withholding Amount')) else 0
            tax_description = str(row['Tax Description']).strip() if pd.notna(row.get('Tax Description')) else ''

            # Only include non-zero tax rate withholdings
            if tax_rate > 0:
                withholding_date = parse_date(withholding_date_str)

                # Get exchange rate on withholding date
                exchange_rate = None
                if YFINANCE_AVAILABLE and withholding_date_str:
                    exchange_rate = get_exchange_rate(withholding_date_str)

                tax_info = {
                    'date': withholding_date,
                    'date_str': withholding_date_str,
                    'tax_rate': tax_rate,
                    'withholding_amount': withholding_amount,
                    'tax_description': tax_description,
                    'exchange_rate': exchange_rate
                }

                current_grant['tax_withholdings'].append(tax_info)
                current_grant['total_tax_withheld'] += withholding_amount

    return grants

def process_espp(df, symbol_for_price='PTC'):
    """
    Process ESPP data and return grants dictionary.
    ESPP is immediately sellable and taxes are paid after selling.

    Parameters:
    -----------
    df : DataFrame
        Input data from ESPP sheet
    symbol_for_price : str
        Stock ticker symbol for historical price lookup (default: 'PTC')

    Returns:
    --------
    dict : Dictionary of processed ESPP grants
    """
    # Standardize column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Remove completely empty rows
    df = df.dropna(how='all')

    # Reset index for easier processing
    df = df.reset_index(drop=True)

    # Dictionary to store grant information
    grants = {}
    current_grant = None
    grant_counter = 0

    # Process each row
    for idx, row in df.iterrows():
        record_type = str(row['Record Type']).strip() if pd.notna(row.get('Record Type')) else ''

        # Handle Grant records (for ESPP, this is a Purchase)
        if record_type == 'Grant':
            grant_counter += 1
            symbol = str(row['Symbol']).strip() if pd.notna(row.get('Symbol')) else ''
            purchase_date_str = str(row['Purchase Date']).strip() if pd.notna(row.get('Purchase Date')) else ''

            # Create unique grant ID
            grant_id = f"ESPP_{purchase_date_str}_{grant_counter}"

            # Parse purchase date
            purchase_date = parse_date(purchase_date_str)

            # Get purchase price for capital gains calculation
            purchase_price = float(row['Purchase Price']) if pd.notna(row.get('Purchase Price')) else None

            # Get grant date (for reference)
            grant_date_str = str(row['Grant Date']).strip() if pd.notna(row.get('Grant Date')) else purchase_date_str
            grant_date = parse_date(grant_date_str)

            # Get purchased quantity and tax collection shares
            purchased_qty = float(row['Purchased Qty.']) if pd.notna(row.get('Purchased Qty.')) else 0
            tax_collection_shares = float(row['Tax Collection Shares']) if pd.notna(row.get('Tax Collection Shares')) else 0
            net_shares = float(row['Net Shares']) if pd.notna(row.get('Net Shares')) else 0
            sellable_qty = float(row['Sellable Qty.']) if pd.notna(row.get('Sellable Qty.')) else 0

            # Initialize ESPP grant dictionary
            current_grant = {
                'grant_id': grant_id,
                'grant_type': 'ESPP',
                'symbol': symbol,
                'grant_date': grant_date,
                'grant_date_str': grant_date_str,
                'purchase_date': purchase_date,
                'purchase_date_str': purchase_date_str,
                'grant_price': purchase_price,  # Use purchase price as basis
                'granted_qty': purchased_qty,  # Use purchased qty
                'withheld_qty': tax_collection_shares,
                'vested_qty': net_shares,  # All purchased shares are immediately available
                'sellable_qty': sellable_qty,
                'unvested_qty': 0,  # ESPP is immediately vested/sellable
                'released_qty': 0,
                'est_market_value': float(row['Est. Market Value']) if pd.notna(row.get('Est. Market Value')) else 0,
                'events': [],  # List of events (sell, dividend, etc)
                'vest_schedules': [],  # Not applicable for ESPP
                'tax_withholdings': [],  # Not applicable for ESPP (taxes paid after sale)
                'sales': [],  # List of sales
                'capital_gains_tax': [],  # List of capital gains taxes
                'total_tax_withheld': 0,  # Will be calculated from sales tax
                'total_capital_gains_tax': 0,
                'total_sold_qty': 0,
                'total_sale_proceeds': 0,
                'sale_dates': [],
                'validation_issues': []
            }

            grants[grant_id] = current_grant

        # Handle Event records
        elif record_type == 'Event' and current_grant is not None:
            event_date_str = str(row['Date']).strip() if pd.notna(row.get('Date')) else ''
            event_type = str(row['Event Type']).strip() if pd.notna(row.get('Event Type')) else ''
            qty = float(row['Qty']) if pd.notna(row.get('Qty')) else 0

            event_date = parse_date(event_date_str)

            event_info = {
                'date': event_date,
                'date_str': event_date_str,
                'type': event_type,
                'quantity': qty
            }

            current_grant['events'].append(event_info)

            # Track sales
            if 'sold' in event_type.lower():
                sale_price = float(row['Sale Price']) if pd.notna(row.get('Sale Price')) else None

                # Try to fetch historical stock price if not available
                if sale_price is None and YFINANCE_AVAILABLE:
                    sale_price = get_stock_price(symbol, event_date_str)

                # Get exchange rate on sale date
                exchange_rate = None
                if YFINANCE_AVAILABLE:
                    exchange_rate = get_exchange_rate(event_date_str)

                # Calculate capital gains tax based on holding period
                capital_gain = 0
                capital_gains_tax = 0
                tax_rate = 0
                tax_type = "N/A"

                if sale_price is not None and current_grant['grant_price'] is not None:
                    capital_gain = (sale_price - current_grant['grant_price']) * qty

                    # Determine tax rate based on holding period
                    tax_rate, tax_type = get_capital_gains_tax_rate(current_grant['purchase_date'], event_date)

                    if tax_rate is not None:
                        capital_gains_tax = capital_gain * tax_rate
                        current_grant['total_capital_gains_tax'] += capital_gains_tax

                        # Track capital gain tax separately
                        current_grant['capital_gains_tax'].append({
                            'date': event_date,
                            'date_str': event_date_str,
                            'grant_price': current_grant['grant_price'],
                            'sale_price': sale_price,
                            'quantity': qty,
                            'capital_gain': capital_gain,
                            'holding_days': (event_date - current_grant['purchase_date']).days,
                            'tax_type': tax_type,
                            'tax_rate': tax_rate,
                            'tax_amount': capital_gains_tax
                        })

                sale_info = {
                    'date': event_date,
                    'date_str': event_date_str,
                    'quantity': qty,
                    'price': sale_price,
                    'grant_price': current_grant['grant_price'],
                    'capital_gain': capital_gain,
                    'capital_gains_tax': capital_gains_tax,
                    'tax_type': tax_type,
                    'tax_rate': tax_rate,
                    'exchange_rate': exchange_rate
                }
                current_grant['sales'].append(sale_info)
                current_grant['total_sold_qty'] += qty
                current_grant['sale_dates'].append(event_date_str)

    return grants

def process_benefit_history(input_file, output_file, symbol_for_price='PTC'):
    """
    Process BenefitHistory Excel file with multiple sheets (ESPP and Restricted Stock).
    Generates combined summary.

    Parameters:
    -----------
    input_file : str
        Path to input Excel file
    output_file : str
        Path for output Excel file
    symbol_for_price : str
        Stock ticker symbol for historical price lookup (default: 'PTC')
    """

    print(f"Reading file: {input_file}")

    # Try to read both sheets from BenefitHistory.xlsx
    try:
        espp_df = pd.read_excel(input_file, sheet_name='ESPP')
        print("Found ESPP sheet")
    except:
        espp_df = None
        print("ESPP sheet not found")

    try:
        rs_df = pd.read_excel(input_file, sheet_name='Restricted Stock')
        print("Found Restricted Stock sheet")
    except:
        rs_df = None
        print("Restricted Stock sheet not found")

    # If neither sheet found, try old format (single sheet)
    if espp_df is None and rs_df is None:
        print("BenefitHistory format not found, trying old single-sheet format")
        rs_df = pd.read_excel(input_file)

    # Process both sheets if available
    all_grants = {}

    if rs_df is not None:
        print("\nProcessing Restricted Stock sheet...")
        rs_grants = process_restricted_stock(rs_df, symbol_for_price, grant_type='RSU')
        all_grants.update(rs_grants)
        print(f"Found {len(rs_grants)} Restricted Stock grants")

    if espp_df is not None:
        print("\nProcessing ESPP sheet...")
        espp_grants = process_espp(espp_df, symbol_for_price)
        all_grants.update(espp_grants)
        print(f"Found {len(espp_grants)} ESPP grants")

    grants = all_grants

    print(f"Found {len(grants)} total grants")

    # Process and validate each grant
    summary_data = []

    for grant_id, grant in grants.items():
        # Calculate derived values
        total_released = sum(event['quantity'] for event in grant['events']
                            if 'released' in event['type'].lower())

        total_vested = sum(event['quantity'] for event in grant['events']
                          if 'vested' in event['type'].lower())

        # Calculate from vest schedules (alternative method)
        vested_from_schedules = sum(schedule['vested_qty'] for schedule in grant['vest_schedules'])
        released_from_schedules = sum(schedule['released_qty'] for schedule in grant['vest_schedules'])

        # Calculate future vesting from schedules
        future_vesting_qty = sum(schedule['vested_qty'] for schedule in grant['vest_schedules']
                                if schedule['is_future'])

        # Calculate next vest date
        future_vest_dates = [schedule['vest_date'] for schedule in grant['vest_schedules']
                            if schedule['is_future'] and schedule['vest_date']]
        next_vest_date = min(future_vest_dates) if future_vest_dates else None

        # Calculate sellable quantity (alternative calculation)
        calculated_sellable = (grant['granted_qty'] - grant['vested_qty'] -
                              grant['withheld_qty'] - grant['total_sold_qty'] +
                              total_released)

        # Calculate unvested quantity (alternative calculation)
        calculated_unvested = grant['granted_qty'] - grant['vested_qty']

        # Validation checks
        validation_issues = []

        # Check 1: Granted = Vested + Unvested
        if abs(grant['granted_qty'] - (grant['vested_qty'] + grant['unvested_qty'])) > 0.01:
            validation_issues.append(f"Granted ({grant['granted_qty']}) ≠ Vested ({grant['vested_qty']}) + Unvested ({grant['unvested_qty']})")

        # Check 2: Sellable Qty matches calculation
        if abs(grant['sellable_qty'] - calculated_sellable) > 0.01:
            validation_issues.append(f"Sellable Qty mismatch: Stored={grant['sellable_qty']}, Calculated={calculated_sellable}")

        # Check 3: Unvested Qty matches calculation
        if abs(grant['unvested_qty'] - calculated_unvested) > 0.01:
            validation_issues.append(f"Unvested Qty mismatch: Stored={grant['unvested_qty']}, Calculated={calculated_unvested}")

        # Format sale dates
        sale_dates_str = '; '.join(sorted(set(grant['sale_dates']))) if grant['sale_dates'] else 'None'

        # Format next vest date
        next_vest_str = next_vest_date.strftime('%Y-%m-%d') if next_vest_date else 'N/A'

        # Format validation issues
        validation_str = ' | '.join(validation_issues) if validation_issues else 'OK'

        # Prepare summary row with Grant Type
        summary_row = {
            'Grant Type': grant.get('grant_type', 'RSU'),  # RSU or ESPP
            'Grant ID': grant['grant_id'],
            'Symbol': grant['symbol'],
            'Grant Date': grant['grant_date_str'],
            'Units': grant['granted_qty'],
            'Vested to Date': grant['vested_qty'],
            'Withheld for Taxes': grant['withheld_qty'],
            'Released to Account': total_released,
            'Tax Withheld ($)': grant['total_tax_withheld'],
            'Sold': grant['total_sold_qty'],
            'Sale Dates': sale_dates_str,
            'Sellable': grant['sellable_qty'],
            'Calc Sellable': calculated_sellable,
            'Unvested': grant['unvested_qty'],
            'Calc Unvested': calculated_unvested,
            'Future Vesting (from schedules)': future_vesting_qty,
            'Next Vest Date': next_vest_str,
            'Estimated Market Value ($)': grant['est_market_value'],
            'Validation Status': validation_str,
            '# of Sales': len(grant['sales']),
            '# of Vest Schedules': len(grant['vest_schedules']),
            '# of Tax Withholdings': len(grant['tax_withholdings'])
        }

        summary_data.append(summary_row)

    # Create summary DataFrame
    summary_df = pd.DataFrame(summary_data)

    # Sort by Grant Date
    summary_df['Grant Date Parsed'] = pd.to_datetime(summary_df['Grant Date'], errors='coerce')
    summary_df = summary_df.sort_values('Grant Date Parsed', ascending=False)
    summary_df = summary_df.drop('Grant Date Parsed', axis=1)

    # Create additional sheets for detailed views
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Main summary sheet
        summary_df.to_excel(writer, sheet_name='Grant Summary', index=False)

        # Auto-adjust column widths
        worksheet = writer.sheets['Grant Summary']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # Create year-wise tax summary sheet
        year_tax_data = []

        # Add withholding taxes
        for grant_id, grant in grants.items():
            for tax in grant['tax_withholdings']:
                fy = get_financial_year(tax['date']) if tax['date'] else get_financial_year(grant['grant_date'])
                tax_row = {
                    'FY': fy,
                    'Grant Type': grant.get('grant_type', 'RSU'),
                    'Grant ID': grant['grant_id'],
                    'Symbol': grant['symbol'],
                    'Tax Type': 'Withholding',
                    'Amount ($)': tax['withholding_amount']
                }
                year_tax_data.append(tax_row)

        # Add capital gains taxes from sales
        for grant_id, grant in grants.items():
            for cg_tax in grant['capital_gains_tax']:
                fy = get_financial_year(cg_tax['date']) if cg_tax['date'] else get_financial_year(grant['grant_date'])
                tax_row = {
                    'FY': fy,
                    'Grant Type': grant.get('grant_type', 'RSU'),
                    'Grant ID': grant['grant_id'],
                    'Symbol': grant['symbol'],
                    'Tax Type': f"Capital Gains ({cg_tax['tax_type']})",
                    'Amount ($)': cg_tax['tax_amount']
                }
                year_tax_data.append(tax_row)

        if year_tax_data:
            year_tax_df = pd.DataFrame(year_tax_data)
            year_tax_df = year_tax_df.sort_values(['FY', 'Grant Type'], ascending=[False, True])
            year_tax_df.to_excel(writer, sheet_name='Year-wise Tax Summary', index=False)

            if OPENPYXL_AVAILABLE:
                worksheet = writer.sheets['Year-wise Tax Summary']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        # Create detailed vesting schedule sheet
        vesting_data = []
        for grant_id, grant in grants.items():
            for schedule in grant['vest_schedules']:
                vesting_row = {
                    'Grant Type': grant.get('grant_type', 'RSU'),
                    'Grant ID': grant['grant_id'],
                    'Symbol': grant['symbol'],
                    'Grant Date': grant['grant_date_str'],
                    'Vest Date': schedule['vest_date_str'],
                    'Vest Period': schedule['vest_period'],
                    'Vested Qty.': schedule['vested_qty'],
                    'Released Qty': schedule['released_qty'],
                    'Is Future': 'Yes' if schedule['is_future'] else 'No'
                }
                vesting_data.append(vesting_row)

        if vesting_data:
            vesting_df = pd.DataFrame(vesting_data)
            vesting_df.to_excel(writer, sheet_name='Vesting Schedule', index=False)

        # Create sales history sheet
        sales_data = []
        for grant_id, grant in grants.items():
            for sale in grant['sales']:
                sales_row = {
                    'Grant Type': grant.get('grant_type', 'RSU'),
                    'Grant ID': grant['grant_id'],
                    'Symbol': grant['symbol'],
                    'Grant Date': grant['grant_date_str'],
                    'Sale Date': sale['date_str'],
                    'Qty. Sold': sale['quantity'],
                    'Sale Price ($)': sale['price'],
                    'Grant Price ($)': sale['grant_price'],
                    'Capital Gain ($)': sale['capital_gain'],
                    'Holding Days': sale.get('holding_days', (sale['date'] - grant['grant_date']).days) if sale['date'] and grant['grant_date'] else 0,
                    'Tax Type': sale['tax_type'],
                    'Tax Rate (%)': sale['tax_rate'] * 100 if sale['tax_rate'] else 0,
                    'Capital Gains Tax ($)': sale['capital_gains_tax'],
                    'Exchange Rate (USD-INR)': sale['exchange_rate']
                }
                sales_data.append(sales_row)

        if sales_data:
            sales_df = pd.DataFrame(sales_data)
            sales_df.to_excel(writer, sheet_name='Sales History', index=False)

            if OPENPYXL_AVAILABLE:
                worksheet = writer.sheets['Sales History']
                # Add formulas for calculated columns
                for row_idx, (idx, row) in enumerate(sales_df.iterrows(), start=2):
                    # Get column indices dynamically
                    col_indices = {col: idx for idx, col in enumerate(sales_df.columns, 1)}

                    cap_gain_col = get_column_letter(col_indices.get('Capital Gain ($)', 1))
                    sale_price_col = get_column_letter(col_indices.get('Sale Price ($)', 1))
                    grant_price_col = get_column_letter(col_indices.get('Grant Price ($)', 1))
                    qty_col = get_column_letter(col_indices.get('Qty. Sold', 1))
                    tax_rate_col = get_column_letter(col_indices.get('Tax Rate (%)', 1))
                    cap_gain_tax_col = get_column_letter(col_indices.get('Capital Gains Tax ($)', 1))
                    proceeds_col = get_column_letter(col_indices.get('Estimated Proceeds ($)', 1))
                    exchange_col = get_column_letter(col_indices.get('Exchange Rate (USD-INR)', 1))
                    proceeds_inr_col = get_column_letter(col_indices.get('Estimated Proceeds (INR)', 1))

                    # Capital Gain ($) = (Sale Price - Grant Price) * Quantity
                    worksheet[f'{cap_gain_col}{row_idx}'] = f'=IF(AND(ISNUMBER({sale_price_col}{row_idx}), ISNUMBER({grant_price_col}{row_idx}), ISNUMBER({qty_col}{row_idx})), ({sale_price_col}{row_idx} - {grant_price_col}{row_idx}) * {qty_col}{row_idx}, "")'

                    # Capital Gains Tax ($) = Capital Gain * Tax Rate / 100
                    worksheet[f'{cap_gain_tax_col}{row_idx}'] = f'=IF(AND(ISNUMBER({cap_gain_col}{row_idx}), ISNUMBER({tax_rate_col}{row_idx})), {cap_gain_col}{row_idx} * {tax_rate_col}{row_idx} / 100, "")'

                    # Apply currency formatting
                    for col in [cap_gain_col, sale_price_col, grant_price_col, cap_gain_tax_col]:
                        worksheet[f'{col}{row_idx}'].number_format = '$#,##0.00'
                    worksheet[f'{tax_rate_col}{row_idx}'].number_format = '0.00'

                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        # Create tax withholding sheet
        tax_data = []
        for grant_id, grant in grants.items():
            for tax in grant['tax_withholdings']:
                tax_row = {
                    'Grant ID': grant['grant_id'],
                    'Grant Type': grant.get('grant_type', 'RSU'),
                    'Symbol': grant['symbol'],
                    'Grant Date': grant['grant_date_str'],
                    'Tax Description': tax['tax_description'],
                    'Tax Rate (%)': tax['tax_rate'],
                    'Withholding Amount ($)': tax['withholding_amount']
                }
                tax_data.append(tax_row)

        if tax_data:
            tax_df = pd.DataFrame(tax_data)
            # Sort by Grant Date
            tax_df['Grant Date Parsed'] = pd.to_datetime(tax_df['Grant Date'], errors='coerce')
            tax_df = tax_df.sort_values('Grant Date Parsed', ascending=False)
            tax_df = tax_df.drop('Grant Date Parsed', axis=1)
            tax_df.to_excel(writer, sheet_name='Tax Withholdings', index=False)

    print(f"Summary created successfully: {output_file}")

    # Check for validation issues
    issues_df = summary_df[summary_df['Validation Status'] != 'OK']
    if not issues_df.empty:
        print(f"\n⚠️  Validation issues found in {len(issues_df)} grants:")
        for idx, row in issues_df.iterrows():
            print(f"  - {row['Grant ID']}: {row['Validation Status']}")
    else:
        print("\n✅ All grants passed validation checks!")

    return summary_df

def process_rsu_tracker(input_file, output_file, symbol_for_price='PTC'):
    """
    Process RSU tracker Excel file and generate structured summary.
    [DEPRECATED: Use process_benefit_history() for new multi-sheet format]

    Parameters:
    -----------
    input_file : str
        Path to input Excel file
    output_file : str
        Path for output Excel file
    symbol_for_price : str
        Stock ticker symbol for historical price lookup (default: 'PTC')
    """

    print(f"Reading file: {input_file}")

    # Read the Excel file - try new format first
    try:
        espp_df = pd.read_excel(input_file, sheet_name='ESPP')
    except:
        espp_df = None

    try:
        rs_df = pd.read_excel(input_file, sheet_name='Restricted Stock')
    except:
        rs_df = None

    # If new format not found, try single sheet
    if rs_df is None and espp_df is None:
        df = pd.read_excel(input_file)
    else:
        # Redirect to new function
        return process_benefit_history(input_file, output_file, symbol_for_price)

    # Standardize column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Remove completely empty rows
    df = df.dropna(how='all')

    # Reset index for easier processing
    df = df.reset_index(drop=True)

    # Dictionary to store grant information
    grants = {}
    current_grant = None
    grant_counter = 0

    print("Processing data...")

    # Process each row
    for idx, row in df.iterrows():
        record_type = str(row['Record Type']).strip() if pd.notna(row.get('Record Type')) else ''

        # Handle Grant records
        if record_type == 'Grant':
            grant_counter += 1
            symbol = str(row['Symbol']).strip() if pd.notna(row.get('Symbol')) else ''
            grant_date_str = str(row['Grant Date']).strip() if pd.notna(row.get('Grant Date')) else ''

            # Create unique grant ID (date + counter for duplicates)
            grant_id = f"{grant_date_str}_{grant_counter}"

            # Parse grant date
            grant_date = parse_date(grant_date_str)

            # Get grant date stock price for capital gains calculation
            grant_price = None
            if YFINANCE_AVAILABLE and symbol:
                grant_price = get_stock_price(symbol, grant_date_str)

            # Initialize grant dictionary
            current_grant = {
                'grant_id': grant_id,
                'symbol': symbol,
                'grant_date': grant_date,
                'grant_date_str': grant_date_str,
                'grant_price': grant_price,
                'granted_qty': float(row['Granted Qty.']) if pd.notna(row.get('Granted Qty.')) else 0,
                'withheld_qty': float(row['Withheld Qty.']) if pd.notna(row.get('Withheld Qty.')) else 0,
                'vested_qty': float(row['Vested Qty.']) if pd.notna(row.get('Vested Qty.')) else 0,
                'sellable_qty': float(row['Sellable Qty.']) if pd.notna(row.get('Sellable Qty.')) else 0,
                'unvested_qty': float(row['Unvested Qty.']) if pd.notna(row.get('Unvested Qty.')) else 0,
                'released_qty': float(row['Released Qty']) if pd.notna(row.get('Released Qty')) else 0,
                'est_market_value': float(row['Est. Market Value']) if pd.notna(row.get('Est. Market Value')) else 0,
                'events': [],  # List of events (vest, release, sell)
                'vest_schedules': [],  # List of vest schedules
                'tax_withholdings': [],  # List of tax withholdings
                'sales': [],  # List of sales
                'capital_gains_tax': [],  # List of capital gains taxes
                'total_tax_withheld': 0,
                'total_capital_gains_tax': 0,
                'total_sold_qty': 0,
                'total_sale_proceeds': 0,
                'sale_dates': [],
                'validation_issues': []
            }

            grants[grant_id] = current_grant

        # Handle Event records (grant, vest, release, sell)
        elif record_type == 'Event' and current_grant is not None:
            event_date_str = str(row['Date']).strip() if pd.notna(row.get('Date')) else ''
            event_type = str(row['Event Type']).strip() if pd.notna(row.get('Event Type')) else ''
            qty_or_amount = float(row['Qty. or Amount']) if pd.notna(row.get('Qty. or Amount')) else 0

            event_date = parse_date(event_date_str)

            event_info = {
                'date': event_date,
                'date_str': event_date_str,
                'type': event_type,
                'quantity': qty_or_amount
            }

            current_grant['events'].append(event_info)

            # Track sales separately
            if 'sold' in event_type.lower():
                sale_price = float(row['Sale Price']) if pd.notna(row.get('Sale Price')) else None

                # Try to fetch historical stock price if not available
                if sale_price is None and YFINANCE_AVAILABLE:
                    sale_price = get_stock_price(symbol_for_price, event_date_str)

                # Get exchange rate on sale date
                exchange_rate = None
                if YFINANCE_AVAILABLE:
                    exchange_rate = get_exchange_rate(event_date_str)

                # Calculate capital gains tax based on holding period
                capital_gain = 0
                capital_gains_tax = 0
                tax_rate = 0
                tax_type = "N/A"

                if sale_price is not None and current_grant['grant_price'] is not None:
                    capital_gain = (sale_price - current_grant['grant_price']) * qty_or_amount

                    # Determine tax rate based on holding period
                    tax_rate, tax_type = get_capital_gains_tax_rate(current_grant['grant_date'], event_date)

                    if tax_rate is not None:
                        capital_gains_tax = capital_gain * tax_rate
                        current_grant['total_capital_gains_tax'] += capital_gains_tax

                        # Track capital gain tax separately
                        current_grant['capital_gains_tax'].append({
                            'date': event_date,
                            'date_str': event_date_str,
                            'grant_price': current_grant['grant_price'],
                            'sale_price': sale_price,
                            'quantity': qty_or_amount,
                            'capital_gain': capital_gain,
                            'holding_days': (event_date - current_grant['grant_date']).days,
                            'tax_type': tax_type,
                            'tax_rate': tax_rate,
                            'tax_amount': capital_gains_tax
                        })

                sale_info = {
                    'date': event_date,
                    'date_str': event_date_str,
                    'quantity': qty_or_amount,
                    'price': sale_price,
                    'grant_price': current_grant['grant_price'],
                    'capital_gain': capital_gain,
                    'capital_gains_tax': capital_gains_tax,
                    'tax_type': tax_type,
                    'tax_rate': tax_rate,
                    'exchange_rate': exchange_rate
                }
                current_grant['sales'].append(sale_info)
                current_grant['total_sold_qty'] += qty_or_amount
                current_grant['sale_dates'].append(event_date_str)

        # Handle Vest Schedule records
        elif record_type == 'Vest Schedule' and current_grant is not None:
            vest_date_str = str(row['Vest Date']).strip() if pd.notna(row.get('Vest Date')) else ''
            vested_qty = float(row['Vested Qty.']) if pd.notna(row.get('Vested Qty.')) else 0
            released_qty = float(row['Released Qty']) if pd.notna(row.get('Released Qty')) else 0
            vest_period = str(row['Vest Period']).strip() if pd.notna(row.get('Vest Period')) else ''

            vest_date = parse_date(vest_date_str)

            vest_schedule = {
                'vest_date': vest_date,
                'vest_date_str': vest_date_str,
                'vested_qty': vested_qty,
                'released_qty': released_qty,
                'vest_period': vest_period,
                'is_future': vest_date > datetime.now() if vest_date else False
            }

            current_grant['vest_schedules'].append(vest_schedule)

        # Handle Tax Withholding records
        elif record_type == 'Tax Withholding' and current_grant is not None:
            withholding_date_str = str(row['Date']).strip() if pd.notna(row.get('Date')) else ''
            tax_rate = parse_percentage(row['Effective Tax Rate']) if pd.notna(row.get('Effective Tax Rate')) else 0
            withholding_amount = float(row['Withholding Amount']) if pd.notna(row.get('Withholding Amount')) else 0
            tax_description = str(row['Tax Description']).strip() if pd.notna(row.get('Tax Description')) else ''

            # Only include non-zero tax rate withholdings
            if tax_rate > 0:
                withholding_date = parse_date(withholding_date_str)

                # Get exchange rate on withholding date
                exchange_rate = None
                if YFINANCE_AVAILABLE and withholding_date_str:
                    exchange_rate = get_exchange_rate(withholding_date_str)

                tax_info = {
                    'date': withholding_date,
                    'date_str': withholding_date_str,
                    'tax_rate': tax_rate,
                    'withholding_amount': withholding_amount,
                    'tax_description': tax_description,
                    'exchange_rate': exchange_rate
                }

                current_grant['tax_withholdings'].append(tax_info)
                current_grant['total_tax_withheld'] += withholding_amount

    print(f"Found {len(grants)} grants")

    # Process and validate each grant
    summary_data = []

    for grant_id, grant in grants.items():
        # Calculate derived values
        total_released = sum(event['quantity'] for event in grant['events']
                            if 'released' in event['type'].lower())

        total_vested = sum(event['quantity'] for event in grant['events']
                          if 'vested' in event['type'].lower())

        # Calculate from vest schedules (alternative method)
        vested_from_schedules = sum(schedule['vested_qty'] for schedule in grant['vest_schedules'])
        released_from_schedules = sum(schedule['released_qty'] for schedule in grant['vest_schedules'])

        # Calculate future vesting from schedules
        future_vesting_qty = sum(schedule['vested_qty'] for schedule in grant['vest_schedules']
                                if schedule['is_future'])

        # Calculate next vest date
        future_vest_dates = [schedule['vest_date'] for schedule in grant['vest_schedules']
                            if schedule['is_future'] and schedule['vest_date']]
        next_vest_date = min(future_vest_dates) if future_vest_dates else None

        # Calculate sellable quantity (alternative calculation)
        calculated_sellable = (grant['granted_qty'] - grant['vested_qty'] -
                              grant['withheld_qty'] - grant['total_sold_qty'] +
                              total_released)

        # Calculate unvested quantity (alternative calculation)
        calculated_unvested = grant['granted_qty'] - grant['vested_qty']

        # Validation checks
        validation_issues = []

        # Check 1: Granted = Vested + Unvested
        if abs(grant['granted_qty'] - (grant['vested_qty'] + grant['unvested_qty'])) > 0.01:
            validation_issues.append(f"Granted ({grant['granted_qty']}) ≠ Vested ({grant['vested_qty']}) + Unvested ({grant['unvested_qty']})")

        # Check 2: Sellable Qty matches calculation
        if abs(grant['sellable_qty'] - calculated_sellable) > 0.01:
            validation_issues.append(f"Sellable Qty mismatch: Stored={grant['sellable_qty']}, Calculated={calculated_sellable}")

        # Check 3: Unvested Qty matches calculation
        if abs(grant['unvested_qty'] - calculated_unvested) > 0.01:
            validation_issues.append(f"Unvested Qty mismatch: Stored={grant['unvested_qty']}, Calculated={calculated_unvested}")

        # Note: Removed checks 4 & 5 comparing Events vs Schedules
        # These can legitimately differ due to:
        # - Tax withholding reducing actual vesting
        # - Events and Schedules coming from different data sources
        # - Planned schedules vs actual outcomes

        # Format sale dates
        sale_dates_str = '; '.join(sorted(set(grant['sale_dates']))) if grant['sale_dates'] else 'None'

        # Format next vest date
        next_vest_str = next_vest_date.strftime('%Y-%m-%d') if next_vest_date else 'N/A'

        # Format validation issues
        validation_str = ' | '.join(validation_issues) if validation_issues else 'OK'

        # Prepare summary row
        summary_row = {
            'Grant Type': grant.get('grant_type', 'RSU'),  # RSU or ESPP
            'Grant ID': grant['grant_id'],
            'Symbol': grant['symbol'],
            'Grant Date': grant['grant_date_str'],
            'Units': grant['granted_qty'],
            'Vested to Date': grant['vested_qty'],
            'Withheld for Taxes': grant['withheld_qty'],
            'Released to Account': total_released,
            'Tax Withheld ($)': grant['total_tax_withheld'],
            'Sold': grant['total_sold_qty'],
            'Sale Dates': sale_dates_str,
            'Sellable': grant['sellable_qty'],
            'Calc Sellable': calculated_sellable,
            'Unvested': grant['unvested_qty'],
            'Calc Unvested': calculated_unvested,
            'Future Vesting (from schedules)': future_vesting_qty,
            'Next Vest Date': next_vest_str,
            'Estimated Market Value ($)': grant['est_market_value'],
            'Validation Status': validation_str,
            '# of Sales': len(grant['sales']),
            '# of Vest Schedules': len(grant['vest_schedules']),
            '# of Tax Withholdings': len(grant['tax_withholdings'])
        }

        summary_data.append(summary_row)

    # Create summary DataFrame
    summary_df = pd.DataFrame(summary_data)

    # Sort by Grant Date
    summary_df['Grant Date Parsed'] = pd.to_datetime(summary_df['Grant Date'], errors='coerce')
    summary_df = summary_df.sort_values('Grant Date Parsed', ascending=False)
    summary_df = summary_df.drop('Grant Date Parsed', axis=1)

    # Create additional sheets for detailed views
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Main summary sheet
        summary_df.to_excel(writer, sheet_name='Grant Summary', index=False)

        # Auto-adjust column widths
        worksheet = writer.sheets['Grant Summary']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # Create year-wise tax summary sheet
        year_tax_data = []

        # Add withholding taxes
        for grant_id, grant in grants.items():
            for tax in grant['tax_withholdings']:
                fy = get_financial_year(tax['date']) if tax['date'] else get_financial_year(grant['grant_date'])
                if fy:
                    # Use exchange rate that was fetched during data processing, default to None if not available
                    exchange_rate = tax.get('exchange_rate', None)

                    year_tax_row = {
                        'Financial Year': fy,
                        'Tax Date': tax['date_str'],
                        'Grant ID': grant['grant_id'],
                        'Symbol': grant['symbol'],
                        'Tax Type': 'Withholding Tax',
                        'Tax Description': tax['tax_description'],
                        'Rate (%)': tax['tax_rate'],
                        'Amount ($)': tax['withholding_amount'],
                        'Exchange Rate (USD-INR)': exchange_rate
                    }
                    year_tax_data.append(year_tax_row)

        # Add capital gains taxes
        for grant_id, grant in grants.items():
            for cg_tax in grant['capital_gains_tax']:
                fy = get_financial_year(cg_tax['date'])
                if fy:
                    holding_days_display = f"{cg_tax['holding_days']} days" if cg_tax['holding_days'] <= 365 else f"{cg_tax['holding_days'] // 365} years {cg_tax['holding_days'] % 365} days"
                    # Get exchange rate from the sale that was already fetched
                    exchange_rate = None
                    for sale in grant['sales']:
                        if sale['date_str'] == cg_tax['date_str']:
                            exchange_rate = sale['exchange_rate']
                            break

                    year_tax_row = {
                        'Financial Year': fy,
                        'Tax Date': cg_tax['date_str'],
                        'Grant ID': grant['grant_id'],
                        'Symbol': grant['symbol'],
                        'Tax Type': f'{cg_tax["tax_type"]} (Holding: {holding_days_display})',
                        'Tax Description': f"Sale @ ${cg_tax['sale_price']:.2f} (Grant: ${cg_tax['grant_price']:.2f})",
                        'Rate (%)': cg_tax['tax_rate'] * 100,
                        'Amount ($)': cg_tax['tax_amount'],
                        'Exchange Rate (USD-INR)': exchange_rate
                    }
                    year_tax_data.append(year_tax_row)

        if year_tax_data:
            year_tax_df = pd.DataFrame(year_tax_data)
            # Sort by Financial Year (descending), then by tax amount
            year_tax_df = year_tax_df.sort_values(['Financial Year', 'Amount ($)'],
                                                  ascending=[False, False])
            year_tax_df.to_excel(writer, sheet_name='Year-wise Tax Summary', index=False)

            # Auto-adjust column widths
            if OPENPYXL_AVAILABLE:
                worksheet = writer.sheets['Year-wise Tax Summary']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        # Create detailed vesting schedule sheet
        vesting_data = []
        for grant_id, grant in grants.items():
            for schedule in grant['vest_schedules']:
                vesting_row = {
                    'Grant ID': grant['grant_id'],
                    'Symbol': grant['symbol'],
                    'Grant Date': grant['grant_date_str'],
                    'Vest Date': schedule['vest_date_str'],
                    'Vest Period': schedule['vest_period'],
                    'Vested Quantity': schedule['vested_qty'],
                    'Released Quantity': schedule['released_qty'],
                    'Is Future Vesting': 'Yes' if schedule['is_future'] else 'No'
                }
                vesting_data.append(vesting_row)

        if vesting_data:
            vesting_df = pd.DataFrame(vesting_data)
            vesting_df.to_excel(writer, sheet_name='Vesting Schedule', index=False)

        # Create sales history sheet
        sales_data = []
        for grant_id, grant in grants.items():
            for sale in grant['sales']:
                # Calculate holding period for display
                holding_days = (sale['date'] - grant['grant_date']).days if sale['date'] and grant['grant_date'] else 0
                holding_display = f"{holding_days} days"
                if holding_days > 365:
                    years = holding_days // 365
                    days = holding_days % 365
                    holding_display = f"{years}y {days}d" if days > 0 else f"{years}y"

                # Store numeric values for formulas
                tax_rate_pct = (sale['tax_rate'] * 100) if sale['tax_rate'] else None

                # Exchange rate display
                exchange_rate_display = sale['exchange_rate'] if sale['exchange_rate'] is not None else 0

                sales_row = {
                    'Grant ID': grant['grant_id'],
                    'Symbol': grant['symbol'],
                    'Grant Date': grant['grant_date_str'],
                    'Sale Date': sale['date_str'],
                    'Holding Period': holding_display,
                    'Quantity Sold': sale['quantity'],
                    'Grant Price ($)': sale['grant_price'],
                    'Sale Price ($)': sale['price'],
                    'Capital Gain ($)': None,  # Will be calculated by formula
                    'Tax Rate (%)': tax_rate_pct,
                    'Tax Type': sale['tax_type'],
                    'Capital Gains Tax ($)': None,  # Will be calculated by formula
                    'Estimated Proceeds ($)': None,  # Will be calculated by formula
                    'Exchange Rate (USD-INR)': exchange_rate_display,
                    'Estimated Proceeds (INR)': None  # Will be replaced with formula
                }
                sales_data.append(sales_row)

        if sales_data:
            sales_df = pd.DataFrame(sales_data)
            # Sort by Sale Date (earliest first)
            sales_df['Sale Date Parsed'] = pd.to_datetime(sales_df['Sale Date'], errors='coerce')
            sales_df = sales_df.sort_values('Sale Date Parsed', ascending=True)
            sales_df = sales_df.drop('Sale Date Parsed', axis=1)

            sales_df.to_excel(writer, sheet_name='Sales History', index=False)

            # Add formulas and formatting for calculated columns
            if OPENPYXL_AVAILABLE:
                from openpyxl.styles import numbers
                worksheet = writer.sheets['Sales History']

                # Find the column indices
                col_indices = {}
                for col_idx, col_cell in enumerate(worksheet[1], 1):
                    col_indices[col_cell.value] = col_idx

                # Add formulas for calculated columns (starting from row 2)
                for row_idx in range(2, len(sales_data) + 2):
                    # Get column letters for reference
                    qty_col = get_column_letter(col_indices.get('Quantity Sold', 1))
                    grant_price_col = get_column_letter(col_indices.get('Grant Price ($)', 1))
                    sale_price_col = get_column_letter(col_indices.get('Sale Price ($)', 1))
                    tax_rate_col = get_column_letter(col_indices.get('Tax Rate (%)', 1))
                    cap_gain_col = get_column_letter(col_indices.get('Capital Gain ($)', 1))
                    cap_gain_tax_col = get_column_letter(col_indices.get('Capital Gains Tax ($)', 1))
                    proceeds_col = get_column_letter(col_indices.get('Estimated Proceeds ($)', 1))
                    exchange_col = get_column_letter(col_indices.get('Exchange Rate (USD-INR)', 1))
                    proceeds_inr_col = get_column_letter(col_indices.get('Estimated Proceeds (INR)', 1))

                    # Capital Gain ($) = (Sale Price - Grant Price) * Quantity
                    worksheet[f'{cap_gain_col}{row_idx}'] = f'=IF(AND(ISNUMBER({sale_price_col}{row_idx}), ISNUMBER({grant_price_col}{row_idx}), ISNUMBER({qty_col}{row_idx})), ({sale_price_col}{row_idx} - {grant_price_col}{row_idx}) * {qty_col}{row_idx}, "")'

                    # Capital Gains Tax ($) = Capital Gain * Tax Rate / 100
                    worksheet[f'{cap_gain_tax_col}{row_idx}'] = f'=IF(AND(ISNUMBER({cap_gain_col}{row_idx}), ISNUMBER({tax_rate_col}{row_idx})), {cap_gain_col}{row_idx} * {tax_rate_col}{row_idx} / 100, "")'

                    # Estimated Proceeds ($) = Quantity * Sale Price
                    worksheet[f'{proceeds_col}{row_idx}'] = f'=IF(AND(ISNUMBER({qty_col}{row_idx}), ISNUMBER({sale_price_col}{row_idx})), {qty_col}{row_idx} * {sale_price_col}{row_idx}, "")'

                    # Estimated Proceeds (INR) = Estimated Proceeds ($) * Exchange Rate
                    worksheet[f'{proceeds_inr_col}{row_idx}'] = f'=IF(AND(ISNUMBER({proceeds_col}{row_idx}), ISNUMBER({exchange_col}{row_idx}), {exchange_col}{row_idx} <> 0), {proceeds_col}{row_idx} * {exchange_col}{row_idx}, "")'

                    # Apply currency formatting
                    for col in [cap_gain_col, sale_price_col, grant_price_col, cap_gain_tax_col, proceeds_col, proceeds_inr_col]:
                        worksheet[f'{col}{row_idx}'].number_format = '$#,##0.00'
                    worksheet[f'{tax_rate_col}{row_idx}'].number_format = '0.00'

                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width

        # Create tax withholding sheet
        tax_data = []
        for grant_id, grant in grants.items():
            for tax in grant['tax_withholdings']:
                tax_row = {
                    'Grant ID': grant['grant_id'],
                    'Symbol': grant['symbol'],
                    'Grant Date': grant['grant_date_str'],
                    'Tax Description': tax['tax_description'],
                    'Tax Rate (%)': tax['tax_rate'],
                    'Withholding Amount ($)': tax['withholding_amount']
                }
                tax_data.append(tax_row)

        if tax_data:
            tax_df = pd.DataFrame(tax_data)
            # Sort by Grant Date
            tax_df['Grant Date Parsed'] = pd.to_datetime(tax_df['Grant Date'], errors='coerce')
            tax_df = tax_df.sort_values('Grant Date Parsed', ascending=False)
            tax_df = tax_df.drop('Grant Date Parsed', axis=1)
            tax_df.to_excel(writer, sheet_name='Tax Withholdings', index=False)

    print(f"Summary created successfully: {output_file}")
    # print(f"\nSummary Statistics:")
    # print(f"- Total grants processed: {len(summary_df)}")
    # print(f"- Total granted quantity: {summary_df['Granted Quantity'].sum():.0f}")
    # print(f"- Total vested quantity: {summary_df['Total Vested to Date'].sum():.0f}")
    # print(f"- Total sellable quantity: {summary_df['Currently Sellable Quantity'].sum():.0f}")
    # print(f"- Total unvested quantity: {summary_df['Currently Locked/Unvested Quantity'].sum():.0f}")

    # Check for validation issues
    issues_df = summary_df[summary_df['Validation Status'] != 'OK']
    if not issues_df.empty:
        print(f"\n⚠️  Validation issues found in {len(issues_df)} grants:")
        for idx, row in issues_df.iterrows():
            print(f"  - {row['Grant ID']}: {row['Validation Status']}")
    else:
        print("\n✅ All grants passed validation checks!")

    return summary_df

def main():
    # File paths
    input_file = "BenefitHistory.xlsx"  # Input file with multiple sheets (ESPP and Restricted Stock)
    output_file = "rsu_summary.xlsx"  # Output file name

    # Process the benefit history (RSU and ESPP)
    summary_df = process_rsu_tracker(input_file, output_file)

    # Display sample of the summary
    # print("\nSample of the summary (first 5 grants):")
    # print(summary_df.head().to_string())

if __name__ == "__main__":
    main()