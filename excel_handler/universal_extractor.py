"""
Universal Excel Data Extractor
Extracts clean numeric data from ANY Excel file structure.
Returns pure float values - no nested objects, no numpy types.
"""
import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Optional, Any, Tuple
from collections import OrderedDict
import re

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

# Known product codes (will also auto-detect)
KNOWN_PRODUCTS = ['MCT360', 'MCT165', 'MCTSTICK10', 'MCTSTICK30', 'MCTSTICK16', 'MCTITTO_C']


def to_float(value: Any) -> Optional[float]:
    """
    Convert value to pure Python float.
    Returns None if conversion fails.
    """
    if value is None or pd.isna(value):
        return None
    
    if isinstance(value, (int, float)):
        if np.isnan(value) or np.isinf(value):
            return None
        return float(value)
    
    if isinstance(value, (np.integer, np.floating)):
        if np.isnan(value) or np.isinf(value):
            return None
        return float(value)
    
    if isinstance(value, str):
        # Remove currency symbols, commas, whitespace
        text = value.replace('$', '').replace('€', '').replace('£', '').replace('¥', '').replace(',', '').strip()
        # Handle parentheses as negative
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


def detect_month_columns(df: pd.DataFrame, max_header_rows: int = 10) -> Dict[str, int]:
    """Detect which columns contain month data."""
    month_columns = {}
    
    # Check header rows
    for row_idx in range(min(max_header_rows, len(df))):
        for col_idx in range(min(30, len(df.columns))):
            cell_value = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
            normalized_month = normalize_month_name(cell_value)
            if normalized_month and normalized_month not in month_columns:
                month_columns[normalized_month] = col_idx
    
    # Also check standard positions (D-O for fiscal months)
    if not month_columns:
        for i, month_name in enumerate(FISCAL_MONTHS):
            col_idx = 3 + i  # D=3, E=4, ..., O=14
            if col_idx < len(df.columns):
                month_columns[month_name] = col_idx
    
    return month_columns


def detect_product_blocks(df: pd.DataFrame, start_row: int = 0) -> List[Dict[str, Any]]:
    """
    Detect product blocks by looking for product codes.
    Specifically detects: MCT360, MCT165, MCTSTICK10, MCTSTICK30, MCTSTICK16, MCTITTO_C
    Returns list of product info with row ranges.
    """
    products = []
    seen_codes = set()
    
    # Look in first 3 columns for product identifiers
    for row_idx in range(start_row, len(df)):
        for col_idx in range(min(3, len(df.columns))):
            cell_value = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else None
            if pd.isna(cell_value):
                continue
            
            cell_str = str(cell_value).strip().upper()
            if not cell_str or len(cell_str) < 2:
                continue
            
            # Check for exact or partial match with known products
            matched_product = None
            for known_product in KNOWN_PRODUCTS:
                # Check if cell contains product code or vice versa
                if known_product in cell_str or cell_str in known_product:
                    matched_product = known_product
                    break
            
            # If matched a known product and not already seen
            if matched_product and matched_product not in seen_codes:
                seen_codes.add(matched_product)
                products.append({
                    'code': matched_product,
                    'row_index': row_idx,
                    'name': cell_str,
                    'column_index': col_idx
                })
            # Also check for product-like codes (alphanumeric, contains letters)
            elif matched_product is None:
                has_letters = any(c.isalpha() for c in cell_str)
                is_long_enough = len(cell_str) >= 2
                not_pure_number = not cell_str.replace('.', '').replace('-', '').isdigit()
                
                if has_letters and is_long_enough and not_pure_number and cell_str not in seen_codes:
                    seen_codes.add(cell_str)
                    products.append({
                        'code': cell_str,
                        'row_index': row_idx,
                        'name': cell_str,
                        'column_index': col_idx
                    })
    
    return products


def extract_monthly_values_for_product(df: pd.DataFrame, product_row: int,
                                       month_columns: Dict[str, int],
                                       num_rows_to_check: int = 30) -> Dict[str, float]:
    """
    Extract monthly values for a product.
    Returns dict: {month_name: float_value}
    All values are pure Python floats.
    """
    monthly_data = {}
    
    # Check rows from product_row to product_row + num_rows_to_check
    end_row = min(product_row + num_rows_to_check, len(df))
    
    for month_name, col_idx in month_columns.items():
        month_values = []
        
        for row_idx in range(product_row, end_row):
            if col_idx >= len(df.columns):
                continue
            
            try:
                cell_value = df.iloc[row_idx, col_idx]
                num_value = to_float(cell_value)
                
                if num_value is not None:
                    month_values.append(num_value)
            except Exception as e:
                logger.debug(f"Error extracting value at row {row_idx}, col {col_idx}: {e}")
                continue
        
        # Sum all values for this month
        if month_values:
            total = sum(month_values)
            monthly_data[month_name] = float(total)  # Ensure pure float
        else:
            monthly_data[month_name] = None
    
    return monthly_data


class UniversalDataExtractor:
    """
    Universal Excel Data Extractor
    Extracts clean numeric data from ANY Excel file structure.
    """
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.excel_file = None
        self.all_products_data = {}
        
    def extract(self) -> Dict[str, Any]:
        """
        Main extraction method.
        Returns clean dictionary structure with pure float values.
        """
        try:
            self.excel_file = pd.ExcelFile(self.file_path)
            all_products_data = {}
            overall_monthly_totals = {}
            
            # Process EVERY sheet
            for sheet_name in self.excel_file.sheet_names:
                try:
                    logger.info(f"Processing sheet: {sheet_name}")
                    df = pd.read_excel(self.file_path, sheet_name=sheet_name, header=None)
                    
                    # Detect month columns
                    month_columns = detect_month_columns(df)
                    if not month_columns:
                        logger.warning(f"No month columns found in sheet: {sheet_name}")
                        continue
                    
                    # Detect product blocks
                    products = detect_product_blocks(df, start_row=0)
                    logger.info(f"Found {len(products)} products in sheet {sheet_name}")
                    
                    # Extract data for each product
                    for product in products:
                        product_code = product['code']
                        product_row = product['row_index']
                        
                        # Extract monthly values
                        monthly_data = extract_monthly_values_for_product(
                            df, product_row, month_columns, num_rows_to_check=30
                        )
                        
                        # Filter out None values and convert to pure floats
                        historical_data = {}
                        for month_name, value in monthly_data.items():
                            if value is not None:
                                historical_data[month_name] = float(value)
                        
                        if historical_data:
                            # Calculate predictions (3-month moving average)
                            predicted_data = self._calculate_predictions(historical_data)
                            
                            # Store product data
                            if product_code not in all_products_data:
                                all_products_data[product_code] = {
                                    'historical': {},
                                    'predicted': {}
                                }
                            
                            # Merge data (in case product appears in multiple sheets)
                            all_products_data[product_code]['historical'].update(historical_data)
                            all_products_data[product_code]['predicted'].update(predicted_data)
                            
                            # Add to overall totals
                            for month_name, value in historical_data.items():
                                if month_name not in overall_monthly_totals:
                                    overall_monthly_totals[month_name] = 0.0
                                overall_monthly_totals[month_name] += float(value)
                
                except Exception as e:
                    logger.error(f"Error processing sheet {sheet_name}: {e}", exc_info=True)
                    continue
            
            # Build result structure
            result = {
                'products': {},
                'overall': {
                    'historical': {},
                    'predicted': {}
                }
            }
            
            # Convert product data to clean format
            for product_code, data in all_products_data.items():
                result['products'][product_code] = {
                    'historical': data['historical'],  # {month: float_value}
                    'predicted': data['predicted']     # {month: float_value}
                }
            
            # Calculate overall predictions
            if overall_monthly_totals:
                overall_historical = {k: float(v) for k, v in overall_monthly_totals.items()}
                overall_predicted = self._calculate_predictions(overall_historical)
                result['overall'] = {
                    'historical': overall_historical,
                    'predicted': overall_predicted
                }
            
            self.all_products_data = result
            return result
            
        except Exception as e:
            logger.error(f"Error in UniversalDataExtractor: {e}", exc_info=True)
            return {
                'products': {},
                'overall': {'historical': {}, 'predicted': {}}
            }
    
    def _calculate_predictions(self, historical_data: Dict[str, float], num_months: int = 6) -> Dict[str, float]:
        """
        Calculate predictions using 3-month moving average.
        Returns dict: {month_name: float_value} for next 6 months.
        All values are pure Python floats.
        """
        if not historical_data:
            return {}
        
        # Get values in fiscal month order - ensure pure floats
        historical_values = []
        for month_name in FISCAL_MONTHS:
            if month_name in historical_data:
                val = historical_data[month_name]
                # Convert to pure float, handle all types
                if val is None or pd.isna(val):
                    continue
                try:
                    float_val = float(val)
                    if not np.isnan(float_val) and not np.isinf(float_val):
                        historical_values.append(float_val)
                except (ValueError, TypeError):
                    continue
        
        if not historical_values:
            return {}
        
        # Calculate forecast value (3-month moving average)
        if len(historical_values) >= 3:
            forecast_value = float(np.mean(historical_values[-3:]))
        elif len(historical_values) >= 2:
            forecast_value = float(np.mean(historical_values))
        else:
            forecast_value = float(historical_values[-1])
        
        # Ensure forecast_value is a pure float (not numpy type)
        forecast_value = float(forecast_value)
        
        # Generate predicted months (next 6 months only)
        predicted = {}
        last_month_idx = -1
        for i, month_name in enumerate(FISCAL_MONTHS):
            if month_name in historical_data:
                last_month_idx = i
        
        if last_month_idx >= 0:
            for i in range(num_months):  # 6 months
                next_idx = (last_month_idx + i + 1) % 12
                next_month = FISCAL_MONTHS[next_idx]
                predicted[next_month] = forecast_value  # Pure float
        
        return predicted

