"""
Utility functions for extracting monthly data and generating predictions from Excel files.
"""
import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional
from datetime import datetime

# Month mapping: Fiscal year months (April to March) mapped to column letters
FISCAL_MONTHS = ['April', 'May', 'June', 'July', 'August', 'September', 
                 'October', 'November', 'December', 'January', 'February', 'March']
MONTH_COLUMNS = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']


def extract_monthly_data_from_annual(annual_data: List[Dict]) -> Dict[str, List[float]]:
    """
    Extract monthly values from annual data structure.
    Returns a dictionary with month names as keys and lists of values as values.
    """
    monthly_data = {month: [] for month in FISCAL_MONTHS}
    
    for row in annual_data:
        if row.get('set_type') in ['previous', 'current']:
            for i, month in enumerate(FISCAL_MONTHS):
                col = MONTH_COLUMNS[i]
                value = row.get(col, "")
                try:
                    # Try to convert to float, handling currency symbols and commas
                    if isinstance(value, str):
                        value = value.replace('$', '').replace(',', '').strip()
                    num_value = float(value) if value else 0.0
                    monthly_data[month].append(num_value)
                except (ValueError, TypeError):
                    monthly_data[month].append(0.0)
    
    return monthly_data


def extract_monthly_data_from_ingredients(ingredient_list: List[Tuple[str, List[Dict]]]) -> Dict[str, Dict[str, List[float]]]:
    """
    Extract monthly data for each ingredient.
    Returns a dictionary: {ingredient_name: {month: [values]}}
    """
    ingredient_monthly_data = {}
    
    for ing_name, rows in ingredient_list:
        if not rows:
            continue
            
        monthly_data = {month: [] for month in FISCAL_MONTHS}
        
        # Filter to current rows (set_type == 'current')
        current_rows = [row for row in rows if row.get('set_type') == 'current']
        
        for row in current_rows:
            for i, month in enumerate(FISCAL_MONTHS):
                col = MONTH_COLUMNS[i]
                value = row.get(col, "")
                try:
                    if isinstance(value, str):
                        value = value.replace('$', '').replace(',', '').strip()
                    num_value = float(value) if value else 0.0
                    monthly_data[month].append(num_value)
                except (ValueError, TypeError):
                    monthly_data[month].append(0.0)
        
        ingredient_monthly_data[ing_name.upper()] = monthly_data
    
    return ingredient_monthly_data


def calculate_monthly_totals(monthly_data: Dict[str, List[float]]) -> Dict[str, float]:
    """Calculate total for each month."""
    return {month: sum(values) if values else 0.0 for month, values in monthly_data.items()}


def predict_next_months(values: List[float], num_months: int = 3, method: str = 'moving_average') -> List[float]:
    """
    Predict next N months based on historical data.
    
    Args:
        values: Historical monthly values
        num_months: Number of months to predict
        method: 'moving_average' or 'linear_trend'
    
    Returns:
        List of predicted values
    """
    if not values or len(values) == 0:
        return [0.0] * num_months
    
    # Remove zeros and get meaningful data
    clean_values = [v for v in values if v > 0]
    if not clean_values:
        return [0.0] * num_months
    
    predictions = []
    
    if method == 'moving_average':
        # Use 3-month moving average
        window = min(3, len(clean_values))
        if window > 0:
            avg = np.mean(clean_values[-window:])
            predictions = [max(0.0, avg)] * num_months
        else:
            predictions = [0.0] * num_months
    
    elif method == 'linear_trend':
        # Simple linear trend
        if len(clean_values) >= 2:
            x = np.arange(len(clean_values))
            y = np.array(clean_values)
            # Fit linear trend
            coeffs = np.polyfit(x, y, 1)
            # Predict next months
            for i in range(1, num_months + 1):
                pred = coeffs[0] * (len(clean_values) + i - 1) + coeffs[1]
                predictions.append(max(0.0, pred))
        else:
            predictions = [clean_values[-1]] * num_months
    
    return predictions


def generate_forecast_data(monthly_totals: Dict[str, float], num_future_months: int = 6) -> Dict:
    """
    Generate forecast data including historical and predicted values.
    
    Returns:
        Dictionary with 'months', 'historical', 'predicted', 'all_months', 'all_values'
    """
    # Get historical months with data
    historical_months = [month for month, value in monthly_totals.items() if value > 0]
    historical_values = [monthly_totals[month] for month in historical_months]
    
    if not historical_values:
        return {
            'months': [],
            'historical': [],
            'predicted': [],
            'all_months': [],
            'all_values': []
        }
    
    # Predict future months
    predictions = predict_next_months(historical_values, num_future_months, 'moving_average')
    
    # Find the last month with data and generate future month names
    if historical_months:
        last_month_idx = FISCAL_MONTHS.index(historical_months[-1])
        future_months = []
        for i in range(1, num_future_months + 1):
            next_idx = (last_month_idx + i) % 12
            future_months.append(FISCAL_MONTHS[next_idx])
    else:
        future_months = FISCAL_MONTHS[:num_future_months]
    
    all_months = historical_months + future_months
    all_values = historical_values + predictions
    
    return {
        'months': historical_months,
        'historical': historical_values,
        'predicted': predictions,
        'future_months': future_months,
        'all_months': all_months,
        'all_values': all_values,
        'prediction_point': len(historical_months) - 1  # Index where prediction starts
    }


def prepare_chart_data(monthly_totals: Dict[str, float], forecast_data: Optional[Dict] = None) -> Dict:
    """
    Prepare data structure for Chart.js.
    
    Returns:
        Dictionary with labels, datasets for Chart.js
    """
    if forecast_data is None:
        forecast_data = generate_forecast_data(monthly_totals)
    
    labels = forecast_data['all_months']
    historical = forecast_data['historical']
    predicted = forecast_data['predicted']
    
    # Combine historical and predicted values
    all_values = historical + predicted
    
    # Create datasets
    datasets = [
        {
            'label': 'Historical Data',
            'data': historical + [None] * len(predicted),  # Pad with None for predicted
            'borderColor': 'rgb(75, 192, 192)',
            'backgroundColor': 'rgba(75, 192, 192, 0.2)',
            'tension': 0.1
        },
        {
            'label': 'Predicted',
            'data': [None] * len(historical) + predicted,  # Pad with None for historical
            'borderColor': 'rgb(255, 99, 132)',
            'backgroundColor': 'rgba(255, 99, 132, 0.2)',
            'borderDash': [5, 5],
            'tension': 0.1
        }
    ]
    
    return {
        'labels': labels,
        'datasets': datasets,
        'all_values': all_values,
        'prediction_index': len(historical) - 1
    }

