"""
Chart Data Builder
Converts extracted data to chart-friendly arrays.
Ensures all values are pure floats, no nested objects.
"""
import logging
from typing import Dict, List, Any, Tuple, Optional
import numpy as np

logger = logging.getLogger(__name__)

FISCAL_MONTHS = ['April', 'May', 'June', 'July', 'August', 'September',
                 'October', 'November', 'December', 'January', 'February', 'March']


class ChartDataBuilder:
    """
    Converts extracted data structure to chart-friendly format.
    Returns clean arrays: months[], historical[], predicted[]
    """
    
    @staticmethod
    def build_chart_data(extracted_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Convert extracted data to chart format.
        
        Input format:
        {
            "products": {
                "MCT360": {
                    "historical": {"April": 17625.22, "May": 18200.44, ...},
                    "predicted": {"July": 18000.0, "August": 18000.0, ...}
                },
                ...
            },
            "overall": {
                "historical": {"April": 50000.0, ...},
                "predicted": {"July": 52000.0, ...}
            }
        }
        
        Output format:
        {
            "products": {
                "MCT360": {
                    "months": ["April", "May", ...],
                    "historical": [17625.22, 18200.44, ...],
                    "predicted": [18000.0, 18000.0, ...]
                },
                ...
            },
            "overall": {
                "months": ["April", "May", ...],
                "historical": [50000.0, ...],
                "predicted": [52000.0, ...]
            }
        }
        """
        result = {
            'products': {},
            'overall': {
                'months': [],
                'historical': [],
                'predicted': []
            }
        }
        
        # Process each product
        products_data = extracted_data.get('products', {})
        for product_code, product_data in products_data.items():
            historical_dict = product_data.get('historical', {})
            predicted_dict = product_data.get('predicted', {})
            
            # Convert to arrays in fiscal month order
            months = []
            historical = []
            predicted = []
            
            # Add historical data
            for month_name in FISCAL_MONTHS:
                if month_name in historical_dict:
                    value = historical_dict[month_name]
                    if value is not None:
                        months.append(month_name)
                        historical.append(float(value))  # Ensure pure float
            
            # Add predicted data (only months not in historical)
            predicted_months = []
            predicted_values = []
            for month_name in FISCAL_MONTHS:
                if month_name in predicted_dict and month_name not in months:
                    value = predicted_dict[month_name]
                    if value is not None:
                        predicted_months.append(month_name)
                        predicted_values.append(float(value))  # Ensure pure float
            
            if months or predicted_months:
                result['products'][product_code] = {
                    'months': months,
                    'historical': historical,
                    'predicted': predicted_values,
                    'predicted_months': predicted_months
                }
        
        # Process overall data
        overall_data = extracted_data.get('overall', {})
        overall_historical_dict = overall_data.get('historical', {})
        overall_predicted_dict = overall_data.get('predicted', {})
        
        overall_months = []
        overall_historical = []
        overall_predicted = []
        
        # Add overall historical
        for month_name in FISCAL_MONTHS:
            if month_name in overall_historical_dict:
                value = overall_historical_dict[month_name]
                if value is not None:
                    overall_months.append(month_name)
                    overall_historical.append(float(value))
        
        # Add overall predicted
        for month_name in FISCAL_MONTHS:
            if month_name in overall_predicted_dict and month_name not in overall_months:
                value = overall_predicted_dict[month_name]
                if value is not None:
                    overall_predicted.append(float(value))
        
        result['overall'] = {
            'months': overall_months,
            'historical': overall_historical,
            'predicted': overall_predicted
        }
        
        return result
    
    @staticmethod
    def build_template_context(chart_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Build Django template context from chart data.
        Returns context dict ready for template rendering.
        """
        context = {
            'chart_data_json': '',
            'overall_months': [],
            'overall_historical': [],
            'overall_predicted': [],
            'ingredient_chart_data': {},
            'ingredients_list': []
        }
        
        # Extract overall data
        overall = chart_data.get('overall', {})
        context['overall_months'] = overall.get('months', [])
        context['overall_historical'] = overall.get('historical', [])
        context['overall_predicted'] = overall.get('predicted', [])
        
        # Extract product data
        products = chart_data.get('products', {})
        ingredient_chart_data = {}
        ingredients_list = []
        
        for product_code, product_data in products.items():
            ingredient_chart_data[product_code] = {
                'months': product_data.get('months', []),
                'historical': product_data.get('historical', []),
                'predicted': product_data.get('predicted', [])
            }
            ingredients_list.append(product_code)
        
        context['ingredient_chart_data'] = ingredient_chart_data
        context['ingredients_list'] = ingredients_list
        
        # Build JSON for JavaScript
        import json
        chart_data_for_js = {
            'overall': {
                'months': context['overall_months'],
                'historical': context['overall_historical'],
                'predicted': context['overall_predicted']
            },
            'ingredients': ingredient_chart_data,
            'ingredients_list': ingredients_list
        }
        context['chart_data_json'] = json.dumps(chart_data_for_js)
        
        return context


# Legacy functions for backward compatibility
def extract_real_data_from_excel(
    ingredient_list: List[Tuple[str, List[Dict]]],
    annual_data: List[Dict]
) -> Dict:
    """
    Legacy function for backward compatibility.
    Extracts data from ingredient_list and annual_data.
    """
    from .excel_extractor import (
        extract_from_ingredient_section,
        FISCAL_MONTHS,
        normalize_numeric_value,
    )
    
    result = {
        'overall': {'months': [], 'historical': [], 'predicted': []},
        'ingredients': {}
    }
    
    # Extract overall monthly totals from annual data
    overall_monthly_values = {}
    
    # Extract from annual data (columns D-O)
    for month_idx, month_name in enumerate(FISCAL_MONTHS):
        col = chr(65 + 3 + month_idx)  # D=3, E=4, ..., O=14
        month_values = []
        
        for row in annual_data:
            if row.get('set_type') in ['previous', 'current']:
                value = row.get(col, "")
                num_value = normalize_numeric_value(value)
                if num_value is not None and num_value != 0:
                    month_values.append(num_value)
        
        if month_values:
            total = sum(month_values)
            if total is not None and not np.isnan(total) and total > 0:
                overall_monthly_values[month_name] = float(total)
    
    # Generate overall forecast
    if overall_monthly_values:
        overall_months = []
        overall_historical = []
        for month_name in FISCAL_MONTHS:
            if month_name in overall_monthly_values:
                overall_months.append(month_name)
                overall_historical.append(overall_monthly_values[month_name])
        
        if len(overall_historical) >= 3:
            forecast_value = float(np.mean(overall_historical[-3:]))
        elif len(overall_historical) >= 2:
            forecast_value = float(np.mean(overall_historical))
        elif len(overall_historical) >= 1:
            forecast_value = float(overall_historical[-1])
        else:
            forecast_value = 0.0
        
        overall_predicted = [forecast_value] * 6
        
        result['overall'] = {
            'months': overall_months,
            'historical': overall_historical,
            'predicted': overall_predicted
        }
    
    # Extract ingredient data
    for ing_name, rows in ingredient_list:
        if not rows:
            continue
        
        monthly_series = extract_from_ingredient_section(rows, ing_name.upper())
        
        if not monthly_series:
            continue
        
        ing_months = []
        ing_historical = []
        
        for month_name in FISCAL_MONTHS:
            if month_name in monthly_series:
                value = monthly_series[month_name]
                if value is not None and not np.isnan(value) and value > 0:
                    ing_months.append(month_name)
                    ing_historical.append(float(value))
        
        # Generate forecast
        if len(ing_historical) >= 3:
            forecast_value = float(np.mean(ing_historical[-3:]))
        elif len(ing_historical) >= 2:
            forecast_value = float(np.mean(ing_historical))
        elif len(ing_historical) >= 1:
            forecast_value = float(ing_historical[-1])
        else:
            forecast_value = 0.0
        
        ing_predicted = [forecast_value] * 6
        
        result['ingredients'][ing_name.upper()] = {
            'months': ing_months,
            'historical': ing_historical,
            'predicted': ing_predicted
        }
    
    return result


def build_chart_data_from_workflow4(
    result,
    ingredient_list: List[Tuple[str, List[Dict]]],
    annual_data: List[Dict]
) -> Dict:
    """
    Legacy function for backward compatibility.
    Builds chart data from workflow4 results.
    """
    chart_data = {
        'overall': {'months': [], 'historical': [], 'predicted': []},
        'ingredients': {}
    }
    
    # Use extract_real_data_from_excel for consistency
    extracted = extract_real_data_from_excel(ingredient_list, annual_data)
    
    chart_data['overall'] = extracted.get('overall', {'months': [], 'historical': [], 'predicted': []})
    chart_data['ingredients'] = extracted.get('ingredients', {})
    
    return chart_data
