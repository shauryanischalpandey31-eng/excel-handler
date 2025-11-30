"""
Builds chart data from Excel files and workflow4 results.
Extracts real data and generates proper forecasts for charts.
"""
import pandas as pd
import numpy as np
import logging
from typing import Dict, List, Tuple, Optional
from datetime import datetime
from collections import OrderedDict

from .workflow4 import MONTH_NAMES
from .excel_extractor import (
    extract_from_ingredient_section,
    FISCAL_MONTHS,
    MONTH_COLUMNS,
    normalize_numeric_value,
)

logger = logging.getLogger(__name__)


def extract_real_data_from_excel(
    ingredient_list: List[Tuple[str, List[Dict]]],
    annual_data: List[Dict]
) -> Dict:
    """
    Extract real monthly data from Excel structure.
    Returns data for overall chart and ingredient charts.
    """
    # Extract overall monthly totals from annual data
    overall_monthly_values = {}
    overall_months = []
    overall_historical = []
    
    # Extract from annual data (columns D-O)
    for month_idx, month_name in enumerate(FISCAL_MONTHS):
        col = MONTH_COLUMNS[month_idx]
        month_values = []
        
        for row in annual_data:
            if row.get('set_type') in ['previous', 'current']:
                value = row.get(col, "")
                num_value = normalize_numeric_value(value)
                if num_value is not None and num_value != 0:
                    month_values.append(num_value)
        
        if month_values:
            total = sum(month_values)
            # Only add if total is valid and > 0
            if total is not None and not np.isnan(total) and total > 0:
                overall_monthly_values[month_name] = total
                overall_months.append(month_name)
                overall_historical.append(total)
    
    # Generate overall forecast using 3-month moving average
    overall_predicted = []
    if len(overall_historical) >= 3:
        # Use last 3 months
        forecast_value = np.mean(overall_historical[-3:])
        overall_predicted = [float(forecast_value)] * 6  # 6 months ahead
    elif len(overall_historical) >= 2:
        # Use available months
        forecast_value = np.mean(overall_historical)
        overall_predicted = [float(forecast_value)] * 6
    elif len(overall_historical) >= 1:
        # Use last value
        forecast_value = overall_historical[-1]
        overall_predicted = [float(forecast_value)] * 6
    else:
        overall_predicted = [0.0] * 6
    
    # Extract ingredient data
    ingredient_data = {}
    
    for ing_name, rows in ingredient_list:
        if not rows:
            continue
        
        # Extract monthly series using the extractor function
        monthly_series = extract_from_ingredient_section(rows, ing_name.upper())
        
        if not monthly_series:
            continue
        
        # Convert to lists in fiscal month order
        ing_months = []
        ing_historical = []
        
        for month_name in FISCAL_MONTHS:
            if month_name in monthly_series:
                value = monthly_series[month_name]
                # Check if value is valid (not NaN, not None, and > 0)
                if value is not None and not np.isnan(value) and value > 0:
                    ing_months.append(month_name)
                    ing_historical.append(float(value))
        
        # Generate forecast for this ingredient using 3-month moving average
        ing_predicted = []
        if len(ing_historical) >= 3:
            forecast_value = np.mean(ing_historical[-3:])
            ing_predicted = [float(forecast_value)] * 6
        elif len(ing_historical) >= 2:
            forecast_value = np.mean(ing_historical)
            ing_predicted = [float(forecast_value)] * 6
        elif len(ing_historical) >= 1:
            forecast_value = ing_historical[-1]
            ing_predicted = [float(forecast_value)] * 6
        else:
            ing_predicted = [0.0] * 6
        
        ingredient_data[ing_name.upper()] = {
            'months': ing_months,
            'historical': ing_historical,
            'predicted': ing_predicted
        }
    
    return {
        'overall': {
            'months': overall_months,
            'historical': overall_historical,
            'predicted': overall_predicted
        },
        'ingredients': ingredient_data
    }


def build_chart_data_from_workflow4(
    result,
    ingredient_list: List[Tuple[str, List[Dict]]],
    annual_data: List[Dict]
) -> Dict:
    """
    Build chart data from workflow4 results.
    Uses real data from monthly_trend and forecast_table.
    """
    # Extract overall data from monthly_trend
    overall_months = []
    overall_historical = []
    overall_predicted = []
    
    if not result.monthly_trend.empty:
        # Aggregate all products for overall chart
        trend_copy = result.monthly_trend.copy()
        trend_copy['month_index'] = trend_copy['month'].apply(
            lambda m: MONTH_NAMES.index(m) + 1 if m in MONTH_NAMES else None
        )
        
        # Group by month and sum demand
        monthly_aggregated = (
            trend_copy.dropna(subset=['month_index'])
            .groupby(['month', 'month_index'], as_index=False)['demand']
            .sum()
            .sort_values('month_index')
        )
        
        for _, row in monthly_aggregated.iterrows():
            month_name = str(row['month'])
            demand_value = row.get('demand', 0)
            # Check if value is valid
            if demand_value is not None and not np.isnan(demand_value):
                overall_months.append(month_name)
                overall_historical.append(float(demand_value))
        
        # Calculate overall forecast (average of last 3 months)
        if len(overall_historical) >= 3:
            forecast_value = np.mean(overall_historical[-3:])
        elif len(overall_historical) >= 2:
            forecast_value = np.mean(overall_historical)
        elif len(overall_historical) >= 1:
            forecast_value = overall_historical[-1]
        else:
            forecast_value = 0.0
        
        overall_predicted = [float(forecast_value)] * 6
    
    # Extract ingredient data
    ingredient_data = {}
    
    if not result.monthly_trend.empty:
        # Group by product to get ingredient-specific data
        for product, group in result.monthly_trend.groupby('product'):
            product_str = str(product).upper()
            
            # Check if this is one of our ingredients
            ingredient_names = ['MCT360', 'MCT165', 'MCTSTICK10', 'MCTSTICK30', 'MCTSTICK16', 'MCTITTO_C']
            if any(ing in product_str for ing in ingredient_names):
                # Find matching ingredient name
                matching_ing = None
                for ing in ingredient_names:
                    if ing in product_str:
                        matching_ing = ing
                        break
                
                if matching_ing:
                    # Sort by month
                    group_sorted = group.sort_values('month')
                    
                    ing_months = []
                    ing_historical = []
                    
                    for _, row in group_sorted.iterrows():
                        month_name = str(row['month'])
                        demand_value = row.get('demand', 0)
                        # Check if value is valid
                        if month_name in MONTH_NAMES and demand_value is not None and not np.isnan(demand_value):
                            ing_months.append(month_name)
                            ing_historical.append(float(demand_value))
                    
                    # Get forecast for this product
                    product_forecast = result.forecast_table[
                        result.forecast_table['Product'] == product
                    ]
                    
                    if not product_forecast.empty:
                        forecast_value = float(product_forecast.iloc[0]['Forecast Demand'])
                    else:
                        # Calculate using 3-month moving average
                        if len(ing_historical) >= 3:
                            forecast_value = np.mean(ing_historical[-3:])
                        elif len(ing_historical) >= 2:
                            forecast_value = np.mean(ing_historical)
                        elif len(ing_historical) >= 1:
                            forecast_value = ing_historical[-1]
                        else:
                            forecast_value = 0.0
                    
                    ingredient_data[matching_ing] = {
                        'months': ing_months,
                        'historical': ing_historical,
                        'predicted': [float(forecast_value)] * 6
                    }
    
    # If no workflow4 data, fall back to Excel extraction
    if not overall_months and annual_data:
        fallback_data = extract_real_data_from_excel(ingredient_list, annual_data)
        return fallback_data
    
    return {
        'overall': {
            'months': overall_months,
            'historical': overall_historical,
            'predicted': overall_predicted
        },
        'ingredients': ingredient_data
    }

