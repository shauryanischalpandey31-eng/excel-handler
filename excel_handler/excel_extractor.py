"""
Robust Excel extraction module for monthly data series.
Handles various Excel formats with proper header matching and data normalization.
"""
import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Optional, Tuple, OrderedDict
from datetime import datetime
from collections import OrderedDict as OD

logger = logging.getLogger(__name__)

# Month name mappings
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

# Fiscal year order (April to March)
FISCAL_MONTHS = ['April', 'May', 'June', 'July', 'August', 'September',
                 'October', 'November', 'December', 'January', 'February', 'March']

# Column letters for fiscal months (D-O)
MONTH_COLUMNS = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']


def normalize_month_name(value: str) -> Optional[str]:
    """
    Normalize month name to canonical form.
    Handles variations like 'Apr ', 'APRIL', '04', etc.
    """
    if pd.isna(value) or value is None:
        return None
    
    text = str(value).strip().lower()
    
    # Remove trailing spaces
    text = text.strip()
    
    # Handle Japanese month format
    if text.endswith('月'):
        text = text.replace('月', '')
    
    # Remove dots
    text = text.replace('.', '')
    
    # Check variants
    if text in MONTH_VARIANTS:
        return MONTH_VARIANTS[text]
    
    # Check if it's a number
    if text.isdigit():
        num = int(text)
        if 1 <= num <= 12:
            return FISCAL_MONTHS[(num - 4) % 12] if num >= 4 else FISCAL_MONTHS[num + 8]
    
    return None


def normalize_numeric_value(value) -> Optional[float]:
    """
    Convert string/numeric value to float.
    Handles commas, currency symbols, parentheses (negative), whitespace.
    """
    if pd.isna(value) or value is None:
        return None
    
    if isinstance(value, (int, float, np.number)):
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
        # Remove whitespace
        text = text.strip()
        
        if not text or text == '-':
            return None
        
        try:
            num = float(text)
            return -num if is_negative else num
        except (ValueError, TypeError):
            return None
    
    return None


def extract_monthly_series(sheet: pd.DataFrame, product_code: str, 
                          sheet_name: str = None) -> OrderedDict:
    """
    Extract monthly data series for a specific product from Excel sheet.
    
    Args:
        sheet: DataFrame from Excel sheet
        product_code: Product identifier (e.g., 'MCT360', 'MCT165')
        sheet_name: Name of the sheet (for logging)
    
    Returns:
        OrderedDict with keys as 'YYYY-MM' format and values as floats.
        Months are ordered chronologically (April to March fiscal year).
    """
    logger.debug("Extracting series for product '%s' from sheet '%s'", 
                 product_code, sheet_name or 'unknown')
    
    # Normalize product code for matching
    product_lower = product_code.lower().strip()
    
    # Try to find product row by matching in first column
    product_row_idx = None
    for idx, row in sheet.iterrows():
        first_col_value = str(row.iloc[0]).lower().strip() if len(row) > 0 else ""
        if product_lower in first_col_value or first_col_value in product_lower:
            product_row_idx = idx
            logger.debug("Found product '%s' at row %d: %s", 
                        product_code, idx, str(row.iloc[0])[:50])
            break
    
    if product_row_idx is None:
        logger.warning("Product '%s' not found in sheet '%s'", product_code, sheet_name)
        return OD()
    
    # Extract the row
    product_row = sheet.iloc[product_row_idx]
    
    # Try to detect month columns
    # Method 1: Use column letters D-O (fiscal months)
    monthly_series = OD()
    
    # Check if we have enough columns
    if len(product_row) >= 16:  # At least up to column O (index 14)
        for i, month_name in enumerate(FISCAL_MONTHS):
            col_idx = 3 + i  # D=3, E=4, ..., O=14
            if col_idx < len(product_row):
                value = product_row.iloc[col_idx]
                num_value = normalize_numeric_value(value)
                if num_value is not None:
                    # Create YYYY-MM key (using fiscal year)
                    # For now, use month name as key, will convert later
                    monthly_series[month_name] = num_value
                    logger.debug("  %s (col %s): %s -> %f", 
                               month_name, chr(65 + col_idx), str(value)[:20], num_value)
    
    # Method 2: Try header-based matching if column letters didn't work
    if not monthly_series and len(sheet.columns) > 0:
        # Look for month names in header row
        header_row = sheet.iloc[0] if len(sheet) > 0 else None
        if header_row is not None:
            for col_idx, col_name in enumerate(sheet.columns):
                normalized_month = normalize_month_name(str(col_name))
                if normalized_month:
                    value = product_row.iloc[col_idx] if col_idx < len(product_row) else None
                    num_value = normalize_numeric_value(value)
                    if num_value is not None:
                        monthly_series[normalized_month] = num_value
    
    # Log extraction results
    if monthly_series:
        logger.debug("Extracted %d months for '%s': first=%s (%.2f), last=%s (%.2f)",
                    len(monthly_series), product_code,
                    list(monthly_series.keys())[0] if monthly_series else 'N/A',
                    list(monthly_series.values())[0] if monthly_series else 0,
                    list(monthly_series.keys())[-1] if monthly_series else 'N/A',
                    list(monthly_series.values())[-1] if monthly_series else 0)
    else:
        logger.warning("No monthly data extracted for '%s' from sheet '%s'", 
                      product_code, sheet_name)
    
    return monthly_series


def extract_from_workflow4_sheet(excel_path: str, product_code: str) -> OrderedDict:
    """
    Extract monthly series from Workflow 4 sheet in processed Excel.
    
    Args:
        excel_path: Path to Excel file
        product_code: Product identifier
    
    Returns:
        OrderedDict with monthly data
    """
    try:
        excel_file = pd.ExcelFile(excel_path)
        
        # Try to find Workflow 4 sheet
        workflow4_sheet = None
        for sheet_name in excel_file.sheet_names:
            if 'workflow' in sheet_name.lower() and '4' in sheet_name:
                workflow4_sheet = pd.read_excel(excel_path, sheet_name=sheet_name)
                logger.debug("Found Workflow 4 sheet: %s", sheet_name)
                break
        
        if workflow4_sheet is None:
            logger.warning("Workflow 4 sheet not found in %s", excel_path)
            return OD()
        
        # Extract from the sheet
        return extract_monthly_series(workflow4_sheet, product_code, "Workflow 4")
    
    except Exception as e:
        logger.error("Error extracting from Workflow 4 sheet: %s", str(e))
        return OD()


def extract_from_ingredient_section(ingredient_data: List[Dict], 
                                   product_code: str) -> OrderedDict:
    """
    Extract monthly series from ingredient section data structure.
    
    Args:
        ingredient_data: List of row dictionaries from ingredient section
        product_code: Product identifier
    
    Returns:
        OrderedDict with monthly data
    """
    monthly_series = OD()
    
    # Filter to current rows
    current_rows = [row for row in ingredient_data if row.get('set_type') == 'current']
    
    if not current_rows:
        logger.warning("No current rows found for product '%s'", product_code)
        return monthly_series
    
    # Extract from row 7 (index 6) which typically contains monthly values
    # Or sum across all current rows
    for i, month_name in enumerate(FISCAL_MONTHS):
        col = MONTH_COLUMNS[i]
        month_values = []
        
        for row in current_rows:
            value = row.get(col, "")
            num_value = normalize_numeric_value(value)
            if num_value is not None and num_value != 0:
                month_values.append(num_value)
        
        if month_values:
            # Use sum or average - typically we want sum for totals
            total = sum(month_values)
            # Only add if total is valid (not NaN)
            if total is not None and not np.isnan(total):
                monthly_series[month_name] = total
                logger.debug("  %s: %d values, sum=%.2f", month_name, len(month_values), total)
    
    return monthly_series

