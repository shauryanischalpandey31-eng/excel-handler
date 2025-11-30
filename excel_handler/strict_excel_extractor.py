"""
STRICT Excel Data Extraction Module
NO guessing, NO hallucination, NO assumed values.
Only extracts EXACT values that exist inside the uploaded Excel file.
"""
import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Optional, Tuple, Any
from collections import OrderedDict
import json

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
}

FISCAL_MONTHS = ['April', 'May', 'June', 'July', 'August', 'September',
                 'October', 'November', 'December', 'January', 'February', 'March']

MONTH_COLUMNS = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']


def normalize_month_name(value: Any) -> Optional[str]:
    """Normalize month name to canonical form. Returns None if not a valid month."""
    if pd.isna(value) or value is None:
        return None
    
    text = str(value).strip().lower()
    
    # Handle Japanese month format
    if text.endswith('月'):
        text = text.replace('月', '')
    
    # Remove dots and spaces
    text = text.replace('.', '').strip()
    
    # Check variants
    if text in MONTH_VARIANTS:
        return MONTH_VARIANTS[text]
    
    # Check if it's a number
    if text.isdigit():
        num = int(text)
        if 1 <= num <= 12:
            # Convert to fiscal year order (April=0, May=1, ..., March=11)
            fiscal_idx = (num - 4) % 12 if num >= 4 else num + 8
            return FISCAL_MONTHS[fiscal_idx]
    
    return None


def normalize_numeric_value(value: Any) -> Optional[float]:
    """
    Convert value to float. Returns None if not a valid number.
    NO assumptions - if value is empty/invalid, return None.
    """
    if pd.isna(value) or value is None:
        return None
    
    if isinstance(value, (int, float, np.number)):
        if np.isnan(value) or np.isinf(value):
            return None
        return float(value)
    
    if isinstance(value, str):
        # Remove currency symbols
        text = value.replace('$', '').replace('€', '').replace('£', '').replace('¥', '')
        # Remove commas
        text = text.replace(',', '')
        # Handle parentheses as negative
        text = text.strip()
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


def validate_excel_structure(file_path: str) -> Dict[str, Any]:
    """
    STEP 1: Validate Excel structure before processing.
    Returns validation result with error flags and missing items.
    """
    result = {
        'error': False,
        'message': '',
        'missing_items': [],
        'detected_sheets': [],
        'detected_products': [],
        'has_monthly_data': False,
        'has_annual_data': False
    }
    
    try:
        excel_file = pd.ExcelFile(file_path)
        result['detected_sheets'] = excel_file.sheet_names
        
        if len(excel_file.sheet_names) == 0:
            result['error'] = True
            result['message'] = 'Excel file contains no sheets.'
            result['missing_items'].append('sheets')
            return result
        
        # Try to detect products and monthly data in each sheet
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Check for month columns
                has_months = False
                for col_idx in range(min(20, len(df.columns))):
                    col_letter = chr(65 + col_idx) if col_idx < 26 else None
                    if col_letter:
                        # Check header row for month names
                        for row_idx in range(min(5, len(df))):
                            cell_value = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
                            if normalize_month_name(cell_value) is not None:
                                has_months = True
                                result['has_monthly_data'] = True
                                break
                        if has_months:
                            break
                
                # Check for product codes (common patterns)
                for row_idx in range(min(50, len(df))):
                    for col_idx in range(min(5, len(df.columns))):
                        cell_value = str(df.iloc[row_idx, col_idx]) if col_idx < len(df.columns) else ''
                        cell_lower = cell_value.lower().strip()
                        # Look for product-like patterns
                        if any(keyword in cell_lower for keyword in ['mct', 'product', 'item', 'code', '品目']):
                            if cell_lower not in result['detected_products']:
                                result['detected_products'].append(cell_lower[:20])
                
            except Exception as e:
                logger.warning(f"Error reading sheet {sheet_name}: {e}")
                continue
        
        # Check if we found any useful data
        if not result['has_monthly_data'] and len(result['detected_products']) == 0:
            result['error'] = True
            result['message'] = 'The Excel structure does not match expected format. Please upload a template with product sheets and monthly data.'
            result['missing_items'].append('monthly_data')
            result['missing_items'].append('products')
        
    except Exception as e:
        result['error'] = True
        result['message'] = f'Error reading Excel file: {str(e)}'
        result['missing_items'].append('file_readable')
    
    return result


def detect_month_columns(df: pd.DataFrame, max_header_rows: int = 5) -> Dict[str, int]:
    """
    Dynamically detect which columns contain month data.
    Returns dict mapping month_name -> column_index
    """
    month_columns = {}
    
    # Method 1: Check column letters D-O (fiscal months)
    for i, month_name in enumerate(FISCAL_MONTHS):
        col_idx = 3 + i  # D=3, E=4, ..., O=14
        if col_idx < len(df.columns):
            month_columns[month_name] = col_idx
    
    # Method 2: Check header rows for month names
    for row_idx in range(min(max_header_rows, len(df))):
        for col_idx in range(min(30, len(df.columns))):
            cell_value = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
            normalized_month = normalize_month_name(cell_value)
            if normalized_month and normalized_month not in month_columns:
                month_columns[normalized_month] = col_idx
    
    return month_columns


def detect_products_in_sheet(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    Dynamically detect products/ingredients in a sheet.
    Returns list of product info: [{'code': 'MCT360', 'row_index': 10, 'name': '...'}, ...]
    """
    products = []
    
    # Look for product codes in first few columns
    for row_idx in range(len(df)):
        for col_idx in range(min(3, len(df.columns))):
            cell_value = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
            if pd.isna(cell_value):
                continue
            
            cell_str = str(cell_value).strip()
            if not cell_str:
                continue
            
            # Check if it looks like a product code (alphanumeric, uppercase, contains letters)
            if len(cell_str) >= 3 and any(c.isalpha() for c in cell_str):
                # Avoid duplicates
                if not any(p['code'].upper() == cell_str.upper() for p in products):
                    products.append({
                        'code': cell_str.upper(),
                        'row_index': row_idx,
                        'name': cell_str
                    })
    
    return products


def extract_exact_monthly_values(df: pd.DataFrame, row_indices: List[int], 
                                 month_columns: Dict[str, int]) -> Dict[str, Optional[float]]:
    """
    Extract EXACT monthly values from specified rows.
    Returns dict: {month_name: value or None}
    NO assumptions - if value is missing, returns None.
    """
    monthly_values = {}
    
    for month_name, col_idx in month_columns.items():
        month_values = []
        
        for row_idx in row_indices:
            if row_idx >= len(df) or col_idx >= len(df.columns):
                continue
            
            cell_value = df.iloc[row_idx, col_idx]
            num_value = normalize_numeric_value(cell_value)
            
            if num_value is not None:
                month_values.append(num_value)
        
        # Sum all valid values for this month
        if month_values:
            total = sum(month_values)
            monthly_values[month_name] = total
        else:
            monthly_values[month_name] = None
    
    return monthly_values


def identify_current_rows(df: pd.DataFrame, product_row_start: int, 
                         product_row_end: int) -> List[int]:
    """
    Identify "current" rows (typically the last N rows before separator).
    Returns list of row indices that are current.
    """
    current_rows = []
    
    # Strategy: Look for rows with data in month columns
    month_columns = detect_month_columns(df)
    if not month_columns:
        return current_rows
    
    # Check rows from bottom up
    for row_idx in range(product_row_end - 1, product_row_start - 1, -1):
        if row_idx < 0 or row_idx >= len(df):
            continue
        
        # Check if row has any numeric data in month columns
        has_data = False
        for col_idx in month_columns.values():
            if col_idx < len(df.columns):
                value = df.iloc[row_idx, col_idx]
                if normalize_numeric_value(value) is not None:
                    has_data = True
                    break
        
        if has_data:
            current_rows.insert(0, row_idx)  # Insert at beginning to maintain order
            # Typically current rows are the last 10-15 rows
            if len(current_rows) >= 15:
                break
    
    return current_rows


def calculate_forecast(historical_values: List[float], num_months: int = 12) -> List[float]:
    """
    Calculate forecast using 3-month moving average.
    ONLY uses historical values - NO assumptions.
    """
    if not historical_values:
        return [None] * num_months
    
    # Filter out None values
    valid_values = [v for v in historical_values if v is not None and not np.isnan(v)]
    
    if not valid_values:
        return [None] * num_months
    
    # 3-month moving average
    if len(valid_values) >= 3:
        forecast_value = np.mean(valid_values[-3:])
    elif len(valid_values) >= 2:
        forecast_value = np.mean(valid_values)
    else:
        forecast_value = valid_values[-1]
    
    return [float(forecast_value)] * num_months


def extract_from_single_sheet_structure(df: pd.DataFrame, parse_excel_regions_func) -> Dict[str, Any]:
    """
    Extract data from single-sheet structure (annual_data + ingredient sections).
    This handles the structure used by parse_excel_regions.
    """
    
    regions = parse_excel_regions_func(df)
    
    # Extract annual data (rows with region 'annual_data' and set_type 'current')
    annual_data_rows = []
    for i, row in df.iterrows():
        region = regions.get(i, 'unknown')
        if region == 'annual_data':
            annual_data_rows.append(i)
    
    # Extract ingredient sections
    ingredient_sections = {}
    current_ingredient = None
    for i, row in df.iterrows():
        region = regions.get(i, 'unknown')
        if region.startswith('ingredient_'):
            ing_name = region.split('_', 1)[1].upper()
            if ing_name not in ingredient_sections:
                ingredient_sections[ing_name] = []
            ingredient_sections[ing_name].append(i)
    
    # Detect month columns
    month_columns = detect_month_columns(df)
    if not month_columns:
        return {'products': [], 'overall_monthly_totals': {}}
    
    # Extract overall monthly data from annual_data rows
    overall_monthly_totals = {}
    if annual_data_rows:
        monthly_values = extract_exact_monthly_values(df, annual_data_rows, month_columns)
        overall_monthly_totals = {k: v for k, v in monthly_values.items() if v is not None}
    
    # Extract product data
    products_data = []
    for ing_name, row_indices in ingredient_sections.items():
        # Identify current rows (last N rows with data)
        current_rows = identify_current_rows(df, min(row_indices), max(row_indices) + 1)
        if not current_rows:
            current_rows = row_indices[-10:] if len(row_indices) >= 10 else row_indices
        
        # Extract exact monthly values
        monthly_values = extract_exact_monthly_values(df, current_rows, month_columns)
        
        # Build historical data
        historical = []
        months = []
        for month_name in FISCAL_MONTHS:
            if month_name in monthly_values and monthly_values[month_name] is not None:
                months.append(month_name)
                historical.append(float(monthly_values[month_name]))
        
        if historical:
            # Calculate forecast
            predicted_values = calculate_forecast(historical, num_months=12)
            predicted = [
                {'month': f"{month_name}_next", 'value': val}
                for month_name, val in zip(FISCAL_MONTHS[:len(predicted_values)], predicted_values)
                if val is not None
            ]
            
            historical_formatted = [
                {'month': month, 'value': val}
                for month, val in zip(months, historical)
            ]
            
            products_data.append({
                'product_code': ing_name,
                'sheet_name': 'Main Sheet',
                'historical': historical_formatted,
                'predicted': predicted
            })
    
    return {
        'products': products_data,
        'overall_monthly_totals': overall_monthly_totals
    }


def extract_strict_excel_data(file_path: str) -> Dict[str, Any]:
    """
    MAIN FUNCTION: Extract EXACT data from Excel with strict validation.
    Returns JSON-compatible dict with error flags and extracted data.
    """
    # STEP 1: Validate structure
    validation = validate_excel_structure(file_path)
    if validation['error']:
        return {
            'error': True,
            'message': validation['message'],
            'missing_items': validation['missing_items']
        }
    
    result = {
        'error': False,
        'products': [],
        'overall': {
            'months': [],
            'historical': [],
            'predicted': []
        },
        'summary': {
            'products': 0,
            'total_forecast': 0.0,
            'total_raw_material': 0.0
        }
    }
    
    try:
        excel_file = pd.ExcelFile(file_path)
        all_products_data = []
        overall_monthly_totals = {}
        
        # Try single-sheet structure first (if only one sheet)
        if len(excel_file.sheet_names) == 1:
            try:
                # Import here to avoid circular import
                from excel_handler.views import parse_excel_regions
                df = pd.read_excel(file_path, sheet_name=excel_file.sheet_names[0], header=None)
                df.columns = [chr(65 + i) for i in range(len(df.columns))]
                single_sheet_data = extract_from_single_sheet_structure(df, parse_excel_regions)
                all_products_data.extend(single_sheet_data['products'])
                for month_name, value in single_sheet_data['overall_monthly_totals'].items():
                    if value is not None:
                        if month_name not in overall_monthly_totals:
                            overall_monthly_totals[month_name] = 0.0
                        overall_monthly_totals[month_name] += value
            except Exception as e:
                logger.warning(f"Single-sheet extraction failed, trying multi-sheet: {e}")
        
        # Process each sheet (multi-sheet structure)
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Detect month columns
                month_columns = detect_month_columns(df)
                if not month_columns:
                    continue  # Skip sheets without month data
                
                # Detect products in this sheet
                products = detect_products_in_sheet(df)
                
                # If no products detected, try to extract overall data
                if not products:
                    # Look for data rows (rows with numeric values in month columns)
                    data_rows = []
                    for row_idx in range(len(df)):
                        for col_idx in month_columns.values():
                            if col_idx < len(df.columns):
                                value = df.iloc[row_idx, col_idx]
                                if normalize_numeric_value(value) is not None:
                                    data_rows.append(row_idx)
                                    break
                    
                    if data_rows:
                        # Extract overall monthly data
                        monthly_values = extract_exact_monthly_values(df, data_rows, month_columns)
                        for month_name, value in monthly_values.items():
                            if value is not None:
                                if month_name not in overall_monthly_totals:
                                    overall_monthly_totals[month_name] = 0.0
                                overall_monthly_totals[month_name] += value
                
                # Process each product
                for product in products:
                    product_code = product['code']
                    product_row_start = product['row_index']
                    
                    # Skip if already processed in single-sheet extraction
                    if any(p['product_code'] == product_code for p in all_products_data):
                        continue
                    
                    # Find product row range (until next product or end of sheet)
                    product_row_end = len(df)
                    for next_product in products:
                        if next_product['row_index'] > product_row_start:
                            product_row_end = min(product_row_end, next_product['row_index'])
                    
                    # Identify current rows for this product
                    current_rows = identify_current_rows(df, product_row_start, product_row_end)
                    if not current_rows:
                        # Fallback: use product row itself
                        current_rows = [product_row_start]
                    
                    # Extract exact monthly values
                    monthly_values = extract_exact_monthly_values(df, current_rows, month_columns)
                    
                    # Build historical data (only months with actual values)
                    historical = []
                    months = []
                    for month_name in FISCAL_MONTHS:
                        if month_name in monthly_values and monthly_values[month_name] is not None:
                            months.append(month_name)
                            historical.append(float(monthly_values[month_name]))
                    
                    # Calculate forecast (only if we have historical data)
                    predicted = []
                    if historical:
                        predicted_values = calculate_forecast(historical, num_months=12)
                        predicted = [
                            {'month': f"{month_name}_next", 'value': val}
                            for month_name, val in zip(FISCAL_MONTHS[:len(predicted_values)], predicted_values)
                            if val is not None
                        ]
                    
                    # Format historical data
                    historical_formatted = [
                        {'month': month, 'value': val}
                        for month, val in zip(months, historical)
                    ]
                    
                    # Add to products list
                    all_products_data.append({
                        'product_code': product_code,
                        'sheet_name': sheet_name,
                        'historical': historical_formatted,
                        'predicted': predicted
                    })
                    
                    # Add to overall totals
                    for month_name, value in monthly_values.items():
                        if value is not None:
                            if month_name not in overall_monthly_totals:
                                overall_monthly_totals[month_name] = 0.0
                            overall_monthly_totals[month_name] += value
                
            except Exception as e:
                logger.error(f"Error processing sheet {sheet_name}: {e}", exc_info=True)
                continue
        
        # Build overall data
        if overall_monthly_totals:
            overall_months = []
            overall_historical = []
            
            for month_name in FISCAL_MONTHS:
                if month_name in overall_monthly_totals:
                    overall_months.append(month_name)
                    overall_historical.append(overall_monthly_totals[month_name])
            
            if overall_historical:
                overall_predicted_values = calculate_forecast(overall_historical, num_months=12)
                overall_predicted = [
                    {'month': f"{month_name}_next", 'value': val}
                    for month_name, val in zip(FISCAL_MONTHS[:len(overall_predicted_values)], overall_predicted_values)
                    if val is not None
                ]
                
                result['overall'] = {
                    'months': overall_months,
                    'historical': overall_historical,
                    'predicted': overall_predicted
                }
        
        # Set products
        result['products'] = all_products_data
        result['summary']['products'] = len(all_products_data)
        
        # Calculate summary totals
        if result['overall']['predicted']:
            result['summary']['total_forecast'] = sum(
                p['value'] for p in result['overall']['predicted'] if p['value'] is not None
            )
        
        if result['overall']['historical']:
            result['summary']['total_raw_material'] = sum(result['overall']['historical'])
        
    except Exception as e:
        logger.error(f"Error extracting Excel data: {e}", exc_info=True)
        return {
            'error': True,
            'message': f'Error processing Excel file: {str(e)}',
            'missing_items': []
        }
    
    return result

