"""
Chart Data Builder
Converts extracted data to chart-friendly arrays.
Ensures all values are pure floats, no nested objects.
"""
import logging
from typing import Dict, List, Any

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
