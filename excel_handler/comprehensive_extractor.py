"""
COMPREHENSIVE Excel Data Extractor
Extracts REAL numbers from EVERY sheet, ALL products, ALL monthly values.
Sends FULL numeric tables to frontend - NO placeholders, NO hardcoded values.
"""
import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Optional, Any, Tuple
from collections import OrderedDict
import openpyxl
from openpyxl import load_workbook

logger = logging.getLogger(__name__)

# Month name normalization
MONTH_VARIANTS = {
    'jan': 'January', 'january': 'January', '01': 'January', '1': 'January',
    'feb': 'February', 'february': 'February', '02': 'February', '2': 'February',
    'mar': 'March', 'march': 'March', '03': 'March', '3': 'March',
    'apr': 'April', 'april': 'April', '04': 'April', '4': 'April',
    'may': 'May', '05': 'May', '5': 'May',
    'jun': 'June', 'june': 'June', '06': 'June', '6': 'June',
    'jul': 'July', 'july': 'July', '07': 'July', '7': 'July',
    'aug': 'August', 'august': 'August', '08': 'August', '8': 'August',
    'sep': 'September', 'sept': 'September', 'september': 'September', '09': 'September', '9': 'September',
    'oct': 'October', 'october': 'October', '10': 'October',
    'nov': 'November', 'november': 'November', '11': 'November',
    'dec': 'December', 'december': 'December', '12': 'December',
    # Japanese months
    '1月': 'January', '2月': 'February', '3月': 'March', '4月': 'April',
    '5月': 'May', '6月': 'June', '7月': 'July', '8月': 'August',
    '9月': 'September', '10月': 'October', '11月': 'November', '12月': 'December',
}

FISCAL_MONTHS = ['April', 'May', 'June', 'July', 'August', 'September',
                 'October', 'November', 'December', 'January', 'February', 'March']


def normalize_month_name(value: Any) -> Optional[str]:
    """Normalize month name to canonical form."""
    if pd.isna(value) or value is None:
        return None
    
    text = str(value).strip()
    
    # Handle Japanese month format
    if '月' in text:
        for jp_month, en_month in MONTH_VARIANTS.items():
            if jp_month in text:
                return en_month
    
    text_lower = text.lower().replace('.', '').strip()
    
    if text_lower in MONTH_VARIANTS:
        return MONTH_VARIANTS[text_lower]
    
    # Check if it's a number
    if text_lower.isdigit():
        num = int(text_lower)
        if 1 <= num <= 12:
            fiscal_idx = (num - 4) % 12 if num >= 4 else num + 8
            return FISCAL_MONTHS[fiscal_idx]
    
    return None


def normalize_numeric_value(value: Any) -> Optional[float]:
    """Convert value to float. Returns None if not a valid number."""
    if pd.isna(value) or value is None:
        return None
    
    if isinstance(value, (int, float, np.number)):
        if np.isnan(value) or np.isinf(value):
            return None
        return float(value)
    
    if isinstance(value, str):
        text = value.replace('$', '').replace('€', '').replace('£', '').replace('¥', '').replace(',', '').strip()
        is_negative = text.startswith('(') and text.endswith(')')
        if is_negative:
            text = text[1:-1]
        text = text.strip()
        
        if not text or text == '-':
            return None
        
        try:
            num = float(text)
            return -num if is_negative else num
        except (ValueError, TypeError):
            return None
    
    return None


def detect_header_row(df: pd.DataFrame, max_rows: int = 20) -> Optional[int]:
    """Detect header row by looking for month names."""
    for row_idx in range(min(max_rows, len(df))):
        month_count = 0
        for col_idx in range(min(30, len(df.columns))):
            cell_value = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
            if normalize_month_name(cell_value) is not None:
                month_count += 1
        if month_count >= 3:  # Found at least 3 months
            return row_idx
    return None


def detect_month_columns(df: pd.DataFrame, header_row: Optional[int] = None) -> Dict[str, int]:
    """Detect which columns contain month data."""
    month_columns = {}
    
    # If header row is provided, check that row
    if header_row is not None and header_row < len(df):
        for col_idx in range(len(df.columns)):
            cell_value = df.iloc[header_row, col_idx]
            normalized_month = normalize_month_name(cell_value)
            if normalized_month:
                month_columns[normalized_month] = col_idx
    
    # Also check standard column positions (D-O for fiscal months)
    if not month_columns:
        for i, month_name in enumerate(FISCAL_MONTHS):
            col_idx = 3 + i  # D=3, E=4, ..., O=14
            if col_idx < len(df.columns):
                month_columns[month_name] = col_idx
    
    return month_columns


def detect_all_products(df: pd.DataFrame, start_row: int = 0) -> List[Dict[str, Any]]:
    """
    Detect ALL products in the sheet.
    Looks for product codes in first few columns.
    """
    products = []
    seen_codes = set()
    
    # Look in first 3 columns for product identifiers
    for row_idx in range(start_row, len(df)):
        for col_idx in range(min(3, len(df.columns))):
            cell_value = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
            if pd.isna(cell_value):
                continue
            
            cell_str = str(cell_value).strip()
            if not cell_str or len(cell_str) < 2:
                continue
            
            # Check if it looks like a product code
            # Product codes typically: contain letters, are 3+ chars, not pure numbers
            has_letters = any(c.isalpha() for c in cell_str)
            is_long_enough = len(cell_str) >= 2
            not_pure_number = not cell_str.replace('.', '').replace('-', '').isdigit()
            
            if has_letters and is_long_enough and not_pure_number:
                code_upper = cell_str.upper()
                if code_upper not in seen_codes:
                    seen_codes.add(code_upper)
                    products.append({
                        'code': code_upper,
                        'row_index': row_idx,
                        'name': cell_str,
                        'column_index': col_idx
                    })
    
    return products


def extract_product_monthly_data(df: pd.DataFrame, product_row: int, 
                                 month_columns: Dict[str, int],
                                 num_rows_to_check: int = 20) -> Dict[str, Optional[float]]:
    """
    Extract ALL monthly values for a product.
    Checks multiple rows below the product row to find data.
    """
    monthly_data = {}
    
    # Check rows from product_row to product_row + num_rows_to_check
    end_row = min(product_row + num_rows_to_check, len(df))
    
    for month_name, col_idx in month_columns.items():
        month_values = []
        
        for row_idx in range(product_row, end_row):
            if col_idx >= len(df.columns):
                continue
            
            cell_value = df.iloc[row_idx, col_idx]
            num_value = normalize_numeric_value(cell_value)
            
            if num_value is not None:
                month_values.append(num_value)
        
        # Sum all values for this month (in case data spans multiple rows)
        if month_values:
            monthly_data[month_name] = sum(month_values)
        else:
            monthly_data[month_name] = None
    
    return monthly_data


def calculate_forecast_from_historical(historical_values: List[float], num_months: int = 12) -> List[float]:
    """Calculate forecast using 3-month moving average."""
    if not historical_values:
        return []
    
    valid_values = [v for v in historical_values if v is not None and not np.isnan(v)]
    
    if not valid_values:
        return []
    
    if len(valid_values) >= 3:
        forecast_value = np.mean(valid_values[-3:])
    elif len(valid_values) >= 2:
        forecast_value = np.mean(valid_values)
    else:
        forecast_value = valid_values[-1]
    
    return [float(forecast_value)] * num_months


def extract_all_sheets_data(file_path: str) -> Dict[str, Any]:
    """
    MAIN FUNCTION: Extract REAL numbers from EVERY sheet, ALL products.
    Returns complete dataset with ALL numeric monthly values.
    """
    result = {
        'error': False,
        'raw_sheets': {},
        'products': [],
        'overall': {
            'months': [],
            'historical': [],
            'predicted': []
        },
        'summary': {
            'total_sheets': 0,
            'total_products': 0,
            'total_forecast': 0.0,
            'total_raw_material': 0.0
        }
    }
    
    try:
        excel_file = pd.ExcelFile(file_path)
        result['summary']['total_sheets'] = len(excel_file.sheet_names)
        
        all_products_data = []
        overall_monthly_totals = {}
        
        # Process EVERY sheet
        for sheet_name in excel_file.sheet_names:
            try:
                logger.info(f"Processing sheet: {sheet_name}")
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Detect header row
                header_row = detect_header_row(df)
                
                # Detect month columns
                month_columns = detect_month_columns(df, header_row)
                
                if not month_columns:
                    logger.warning(f"No month columns found in sheet: {sheet_name}")
                    continue
                
                logger.info(f"Found {len(month_columns)} month columns in {sheet_name}")
                
                # Detect ALL products in this sheet
                start_search_row = header_row + 1 if header_row is not None else 0
                products = detect_all_products(df, start_row=start_search_row)
                
                logger.info(f"Found {len(products)} products in sheet {sheet_name}: {[p['code'] for p in products]}")
                
                # Extract data for each product
                sheet_products_data = {}
                
                for product in products:
                    product_code = product['code']
                    product_row = product['row_index']
                    
                    # Extract ALL monthly values for this product
                    monthly_data = extract_product_monthly_data(
                        df, product_row, month_columns, num_rows_to_check=30
                    )
                    
                    # Build historical data (only months with actual values)
                    historical = []
                    months = []
                    for month_name in FISCAL_MONTHS:
                        if month_name in monthly_data and monthly_data[month_name] is not None:
                            months.append(month_name)
                            historical.append(float(monthly_data[month_name]))
                    
                    if historical:
                        # Calculate forecast
                        predicted_values = calculate_forecast_from_historical(historical, num_months=12)
                        
                        # Format historical data
                        historical_formatted = [
                            {'month': month, 'value': val}
                            for month, val in zip(months, historical)
                        ]
                        
                        # Format predicted data
                        predicted_formatted = [
                            {'month': f"{month_name}_next", 'value': val}
                            for month_name, val in zip(FISCAL_MONTHS[:len(predicted_values)], predicted_values)
                        ]
                        
                        # Store in sheet products
                        sheet_products_data[product_code] = {
                            'historical': historical_formatted,
                            'predicted': predicted_formatted,
                            'forecast_target_months': [f"{m}_next" for m in FISCAL_MONTHS[:12]]
                        }
                        
                        # Add to overall products list
                        all_products_data.append({
                            'product_code': product_code,
                            'sheet_name': sheet_name,
                            'historical': historical_formatted,
                            'predicted': predicted_formatted
                        })
                        
                        # Add to overall totals
                        for month_name, value in monthly_data.items():
                            if value is not None:
                                if month_name not in overall_monthly_totals:
                                    overall_monthly_totals[month_name] = 0.0
                                overall_monthly_totals[month_name] += value
                
                # Store sheet data
                if sheet_products_data:
                    result['raw_sheets'][sheet_name] = sheet_products_data
                
            except Exception as e:
                logger.error(f"Error processing sheet {sheet_name}: {e}", exc_info=True)
                continue
        
        # Build overall aggregated data
        if overall_monthly_totals:
            overall_months = []
            overall_historical = []
            
            for month_name in FISCAL_MONTHS:
                if month_name in overall_monthly_totals:
                    overall_months.append(month_name)
                    overall_historical.append(overall_monthly_totals[month_name])
            
            if overall_historical:
                overall_predicted_values = calculate_forecast_from_historical(overall_historical, num_months=12)
                overall_predicted = [
                    {'month': f"{month_name}_next", 'value': val}
                    for month_name, val in zip(FISCAL_MONTHS[:len(overall_predicted_values)], overall_predicted_values)
                ]
                
                result['overall'] = {
                    'months': overall_months,
                    'historical': overall_historical,
                    'predicted': overall_predicted
                }
        
        # Set products and summary
        result['products'] = all_products_data
        result['summary']['total_products'] = len(all_products_data)
        
        if result['overall']['predicted']:
            result['summary']['total_forecast'] = sum(
                p['value'] for p in result['overall']['predicted'] if p['value'] is not None
            )
        
        if result['overall']['historical']:
            result['summary']['total_raw_material'] = sum(result['overall']['historical'])
        
        logger.info(f"Extraction complete: {result['summary']['total_products']} products from {result['summary']['total_sheets']} sheets")
        
    except Exception as e:
        logger.error(f"Error extracting Excel data: {e}", exc_info=True)
        return {
            'error': True,
            'message': f'Error processing Excel file: {str(e)}',
            'raw_sheets': {},
            'products': [],
            'overall': {'months': [], 'historical': [], 'predicted': []},
            'summary': {'total_sheets': 0, 'total_products': 0, 'total_forecast': 0.0, 'total_raw_material': 0.0}
        }
    
    return result

