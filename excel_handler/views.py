from datetime import datetime
from pathlib import Path

from django.conf import settings
from django.http import HttpResponse, JsonResponse
from django.shortcuts import render
import pandas as pd
import json
import io
import numpy as np
import os

from .models import ProcessedData, UploadedExcelFile
from .workflow4 import (
    MONTH_NAMES, 
    Workflow4Result, 
    run_workflow4_pipeline,
    write_results_to_original_excel,
)
from .prediction_utils import (
    extract_monthly_data_from_annual,
    extract_monthly_data_from_ingredients,
    calculate_monthly_totals,
    generate_forecast_data,
    prepare_chart_data,
)
from .excel_extractor import (
    extract_monthly_series,
    extract_from_workflow4_sheet,
    extract_from_ingredient_section,
    normalize_month_name,
    normalize_numeric_value,
    FISCAL_MONTHS,
)
from .chart_data_builder import (
    extract_real_data_from_excel,
    build_chart_data_from_workflow4,
)
from .strict_excel_extractor import (
    extract_strict_excel_data,
    validate_excel_structure,
)
from .comprehensive_extractor import extract_all_sheets_data
from .universal_extractor import UniversalDataExtractor
from .chart_data_builder import ChartDataBuilder
import logging

logger = logging.getLogger(__name__)

def compare_files(df, sample_df):
    # Skeletal comparison function - always pass for now
    # TODO: Implement validation rules here
    return True

def index(request):
    data = None
    annual_data = []
    ingredient_list = []
    uploaded_file = None
    error_message = None
    warning_message = None
    forecast_data = None
    chart_data_json = None
    ingredient_charts_json = {}
    ingredients_json = "[]"  # Default empty array for JavaScript
    
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        uploaded_file = UploadedExcelFile.objects.create(file=excel_file)
        try:
            df = pd.read_excel(uploaded_file.file.path, header=None)
            df.columns = [chr(65 + i) for i in range(len(df.columns))]
            # Parse the structure and assign regions
            regions = parse_excel_regions(df)
            data = []
            for i, row in df.iterrows():
                row_dict = row.to_dict()
                row_dict = {k: v if pd.notna(v) else "" for k, v in row_dict.items()}
                region = regions.get(i, 'unknown')
                row_dict['region'] = region
                row_dict['original_row'] = i
                data.append(row_dict)
            
            annual_data = [{k: v for k, v in row.items() if k != 'region' and k <= 'P'} for row in data if row['region'] in ['annual_data', 'annual_separator']][1:]
            # Assign set types for coloring
            data_row_count = 0
            for row in annual_data:
                if all(v == "" for v in row.values()):
                    row['set_type'] = 'separator'
                else:
                    if data_row_count < 5:
                        row['set_type'] = 'previous'
                    else:
                        row['set_type'] = 'current'
                    data_row_count += 1
            request.session['annual_data'] = annual_data
            
            # Extract ingredient data
            ingredients = ['mct360', 'mct165', 'mctstick10', 'mctstick30', 'mctstick16', 'mctitto_c']
            ingredient_list = [(ing, []) for ing in ingredients]
            for row in data:
                region = row.get('region', 'unknown')
                if region.startswith('ingredient_'):
                    ing = region.split('_', 1)[1]
                    for item in ingredient_list:
                        if item[0] == ing:
                             filtered_row = {k: v for k, v in row.items() if k != 'region' and k <= 'Q'}
                             item[1].append(filtered_row)
                             break
            
            # Remove trailing empty rows from each ingredient block
            for item in ingredient_list:
                rows = item[1]
                while rows and all(rows[-1].get(col, "") == "" for col in ['C','D','E','F']):
                    rows.pop()
                # Assign set types for coloring
                for i, row in enumerate(rows):
                    if i == 0:
                        row['set_type'] = 'header'
                    elif 1 <= i <= 10:
                        row['set_type'] = 'previous'
                    elif i >= len(rows) - 10:
                        row['set_type'] = 'current'
                    else:
                        row['set_type'] = 'separator'
            
            # Check if we found any ingredient data
            has_ingredient_data = any(len(item[1]) > 0 for item in ingredient_list)
            if not has_ingredient_data:
                warning_message = "Warning: The uploaded file does not match the expected structure. The file should contain ingredient sections (MCT360, MCT165, MCTSTICK10, MCTSTICK30, MCTSTICK16, MCTITTO_C) and annual data sections. Please download the sample file to see the expected format."
            
            # Extract monthly data and generate predictions
            monthly_data = None
            ingredient_monthly_data = None
            
            if annual_data:
                try:
                    monthly_data = extract_monthly_data_from_annual(annual_data)
                    monthly_totals = calculate_monthly_totals(monthly_data)
                    forecast_data = generate_forecast_data(monthly_totals, num_future_months=6)
                    chart_data = prepare_chart_data(monthly_totals, forecast_data)
                    # Convert to JSON for template
                    chart_data_json = json.dumps(chart_data)
                except Exception as e:
                    print(f"Error generating forecast: {str(e)}")
            
            if has_ingredient_data:
                try:
                    ingredient_monthly_data = extract_monthly_data_from_ingredients(ingredient_list)
                    for ing_name, ing_monthly in ingredient_monthly_data.items():
                        ing_totals = calculate_monthly_totals(ing_monthly)
                        ing_forecast = generate_forecast_data(ing_totals, num_future_months=6)
                        ing_chart_data = prepare_chart_data(ing_totals, ing_forecast)
                        ingredient_charts_json[ing_name] = json.dumps(ing_chart_data)
                except Exception as e:
                    print(f"Error generating ingredient charts: {str(e)}")
            
            ProcessedData.objects.create(original_file=uploaded_file, data=json.loads(json.dumps(data, default=str)))
        except Exception as e:
            error_message = f"Error processing file: {str(e)}"
            logger.error("Error processing file: %s", str(e), exc_info=True)
    
    # Extract chart data using Universal Extractor
    overall_months = []
    overall_historical = []
    overall_predicted = []
    ingredient_chart_data = {}
    ingredients_json = "[]"  # Default empty array
    all_products_list = []
    
    # Use Universal Extractor to get clean data from ALL sheets
    chart_data_json = json.dumps({
        'overall': {'months': [], 'historical': [], 'predicted': []},
        'ingredients': {},
        'ingredients_list': []
    })
    
    if uploaded_file:
        try:
            file_path = uploaded_file.file.path
            if file_path and os.path.exists(file_path):
                logger.info(f"Extracting data using Universal Extractor from: {file_path}")
                
                # Extract data
                extractor = UniversalDataExtractor(file_path)
                extracted_data = extractor.extract()
                
                # Build chart data
                chart_builder = ChartDataBuilder()
                chart_data = chart_builder.build_chart_data(extracted_data)
                
                # Build template context
                template_context = chart_builder.build_template_context(chart_data)
                
                # Extract values for template
                overall_months = template_context.get('overall_months', [])
                overall_historical = template_context.get('overall_historical', [])
                overall_predicted = template_context.get('overall_predicted', [])
                ingredient_chart_data = template_context.get('ingredient_chart_data', {})
                all_products_list = template_context.get('ingredients_list', [])
                ingredients_json = json.dumps(all_products_list)
                chart_data_json = template_context.get('chart_data_json', chart_data_json)
                
                logger.info(f"Extracted {len(all_products_list)} products: {all_products_list}")
                logger.info(f"Overall data: {len(overall_months)} months, {len(overall_historical)} historical, {len(overall_predicted)} predicted")
            else:
                logger.warning(f"File path does not exist: {file_path if uploaded_file else 'None'}")
                
        except Exception as e:
            logger.error("Error extracting data with Universal Extractor: %s", str(e), exc_info=True)
            # Fallback to empty data - don't crash
            chart_data_json = json.dumps({
                'overall': {'months': [], 'historical': [], 'predicted': []},
                'ingredients': {},
                'ingredients_list': []
            })
    
    return render(request, 'excel_handler/index.html', {
        'data': data, 
        'annual_data': annual_data, 
        'ingredient_list': ingredient_list, 
        'uploaded_file': uploaded_file,
        'error_message': error_message,
        'warning_message': warning_message,
        'chart_data_json': chart_data_json,
    })

def parse_excel_regions(df):
    """
    Parse Excel file structure to identify regions.
    Handles both the expected structure (with annual data and ingredients) 
    and simpler files that don't match the expected format.
    """
    regions = {}
    total_rows = len(df)
    
    if total_rows == 0:
        return regions
    
    # Row 0: title
    regions[0] = 'title'
    
    # Row 1: empty (if exists)
    if 1 < total_rows:
        regions[1] = 'empty'
    
    # Find the next completely empty row (separator)
    separator_row = -1
    for i in range(2, min(total_rows, 100)):  # Limit search to first 100 rows
        if df.iloc[i].isna().all():
            separator_row = i
            break
    
    # If no separator found, try to detect structure by looking for ingredient codes
    if separator_row == -1:
        # Check if we can find ingredient codes anywhere in column A
        ingredients = ['MCT360', 'MCT165', 'MCTSTICK10', 'MCTSTICK30', 'MCTSTICK16', 'MCTITTO_C']
        found_ingredients = {}
        
        # Search column A for ingredient codes
        if 'A' in df.columns:
            for i in range(2, min(total_rows, 200)):  # Search first 200 rows
                try:
                    cell_value = str(df.iloc[i]['A']).lower()
                    for ing in ingredients:
                        if ing.lower() in cell_value and ing not in found_ingredients:
                            found_ingredients[ing] = i
                except (KeyError, IndexError):
                    continue
        
        # If we found ingredients, mark them
        if found_ingredients:
            sorted_ingredients = sorted(found_ingredients.items(), key=lambda x: x[1])
            for idx, (ing, start) in enumerate(sorted_ingredients):
                end = sorted_ingredients[idx + 1][1] if idx + 1 < len(sorted_ingredients) else total_rows
                for j in range(start, end):
                    regions[j] = 'ingredient_' + ing.lower()
            # Mark rows before first ingredient as annual_data or unknown
            first_ingredient_row = sorted_ingredients[0][1]
            for i in range(2, first_ingredient_row):
                regions[i] = 'annual_data' if i > 2 else 'annual_header'
        else:
            # No structure detected, mark all as unknown
            for i in range(2, total_rows):
                regions[i] = 'unknown'
        return regions
    
    # Annual data set: rows 2 to separator_row - 1
    # First row: annual_header
    if 2 < separator_row:
        regions[2] = 'annual_header'
    
    # Then, sets of 5 rows, separated by empty in C-P
    row_idx = 3
    while row_idx < separator_row:
        # Check if columns C-P exist before checking
        max_col_idx = min(16, len(df.columns))
        if max_col_idx > 2:
            if df.iloc[row_idx].iloc[2:max_col_idx].isna().all():  # empty in C to P
                regions[row_idx] = 'annual_separator'
                row_idx += 1
            else:
                # Take next 5 rows as annual_data set
                for i in range(5):
                    if row_idx + i < separator_row:
                        regions[row_idx + i] = 'annual_data'
                row_idx += 5
        else:
            # Not enough columns, mark as annual_data
            regions[row_idx] = 'annual_data'
            row_idx += 1
    
    # Ingredient data set: from separator_row to end
    start_ingredient = separator_row
    if start_ingredient < total_rows and df.iloc[start_ingredient].isna().all():
        regions[start_ingredient] = 'empty'
        start_ingredient += 1
    
    # Strip first three rows as ingredient headers
    for i in range(3):
        if start_ingredient + i < total_rows:
            regions[start_ingredient + i] = 'ingredient_header'
    
    # Detect ingredient regions starting from after headers
    ingredients = ['MCT360', 'MCT165', 'MCTSTICK10', 'MCTSTICK30', 'MCTSTICK16', 'MCTITTO_C']
    current_start = start_ingredient + 3
    starts = {}
    temp_start = current_start
    
    # Check if column A exists
    if 'A' in df.columns:
        for ing in ingredients:
            for i in range(temp_start, min(total_rows, temp_start + 500)):  # Limit search
                try:
                    cell_value = str(df.iloc[i]['A']).lower()
                    if ing.lower() in cell_value:
                        starts[ing] = i
                        temp_start = i + 1
                        break
                except (KeyError, IndexError):
                    continue
    
    # Sort by start index
    if starts:
        sorted_starts = sorted(starts.items(), key=lambda x: x[1])
        for idx, (ing, start) in enumerate(sorted_starts):
            end = sorted_starts[idx + 1][1] if idx + 1 < len(sorted_starts) else total_rows
            for j in range(start, end):
                regions[j] = 'ingredient_' + ing.lower()
    
    return regions


def workflow4_view(request):
    context = {}
    if request.method == 'POST':
        excel_file = request.FILES.get('excel_file')
        if not excel_file:
            context['error'] = 'Please upload an Excel file to run workflow 4.'
        else:
            try:
                # Save uploaded file temporarily to get path
                uploaded_file_obj = UploadedExcelFile.objects.create(file=excel_file)
                original_path = Path(uploaded_file_obj.file.path)
                
                result = run_workflow4_pipeline(
                    str(original_path),
                    processed_dir=Path(settings.MEDIA_ROOT) / 'uploads' / 'processed',
                    charts_dir=Path(settings.BASE_DIR) / 'static' / 'charts',
                    original_file_path=original_path,
                )
                context.update(_build_workflow4_context(result))
                context['success'] = 'Workflow 4 completed successfully.'
            except ValueError as exc:
                context['error'] = str(exc)
            except Exception as exc:
                context['error'] = f'An error occurred: {str(exc)}'
    return render(request, 'excel_handler/workflow4.html', context)


def _build_workflow4_context(result: Workflow4Result) -> dict:
    forecast_rows = [
        {
            'product': row['Product'],
            'forecast_demand': row['Forecast Demand'],
            'per_unit_consumption': row['Per Unit Consumption'],
            'raw_material_needed': row['Raw Material Needed'],
        }
        for row in result.forecast_table.to_dict('records')
    ]
    raw_material_rows = [
        {
            'product': row['product'],
            'forecast_demand': row['forecast_demand'],
            'raw_material_needed': row['raw_material_needed'],
        }
        for row in forecast_rows
    ]
    trend_rows = []
    if not result.monthly_trend.empty:
        demand_copy = result.monthly_trend.copy()
        demand_copy['month_index'] = demand_copy['month'].apply(
            lambda label: MONTH_NAMES.index(label) + 1 if label in MONTH_NAMES else None
        )
        aggregated = (
            demand_copy.dropna(subset=['month_index'])
            .groupby(['month', 'month_index'], as_index=False)['demand']
            .sum()
            .sort_values('month_index')
            .drop(columns=['month_index'])
        )
        trend_rows = aggregated.to_dict('records')
    chart_version = int(datetime.utcnow().timestamp())
    demand_chart_url = f"{settings.STATIC_URL}{result.charts['demand']}?v={chart_version}"
    raw_chart_url = f"{settings.STATIC_URL}{result.charts['raw_material']}?v={chart_version}"
    try:
        relative_excel = result.final_excel_path.relative_to(settings.MEDIA_ROOT)
        relative_str = str(relative_excel)
    except ValueError:
        relative_str = result.final_excel_path.name
    download_url = f"{settings.MEDIA_URL}{relative_str.replace(os.sep, '/')}"
    return {
        'forecast_rows': forecast_rows,
        'raw_material_rows': raw_material_rows,
        'demand_trend': trend_rows,
        'summary': result.summary,
        'demand_chart_url': demand_chart_url,
        'raw_chart_url': raw_chart_url,
        'final_excel_url': download_url,
        'workflow_outputs': {name: df.head(10).to_dict('records') for name, df in result.workflow_outputs.items()},
    }

def upload_excel(request):
    if request.method == 'POST' and request.FILES.get('excel_file'):
        excel_file = request.FILES['excel_file']
        uploaded_file = UploadedExcelFile.objects.create(file=excel_file)
        # Process the file
        df = pd.read_excel(uploaded_file.file.path, header=None)
        df.columns = [chr(65 + i) for i in range(len(df.columns))]
        # Example manipulation: filter rows where first column > 10 (assuming numeric)
        if not df.empty and len(df.columns) > 0:
            col = df.columns[0]
            df_filtered = df[df[col] > 10] if pd.api.types.is_numeric_dtype(df[col]) else df
        else:
            df_filtered = df
        df_filtered = df_filtered.replace({np.nan: None})
        # Save processed data
        data_json = df_filtered.to_dict(orient='records')
        ProcessedData.objects.create(original_file=uploaded_file, data=data_json)
        return JsonResponse({'message': 'File uploaded and processed', 'id': uploaded_file.id})
    return JsonResponse({'error': 'Invalid request'}, status=400)

def download_processed(request, file_id):
    try:
        processed = ProcessedData.objects.get(original_file_id=file_id)
        df = pd.DataFrame(processed.data)
        buffer = io.BytesIO()
        df.to_excel(buffer, index=False, engine='openpyxl')
        buffer.seek(0)
        response = HttpResponse(buffer.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename="processed_{file_id}.xlsx"'
        return response
    except ProcessedData.DoesNotExist:
        return JsonResponse({'error': 'Processed data not found'}, status=404)


def process_all_workflows(request):
    """
    Main processing endpoint that runs all workflows including Workflow-4,
    writes results back to the original Excel file, and returns download link.
    """
    if request.method != 'POST':
        return JsonResponse({'error': 'Invalid request method'}, status=400)
    
    file_id = request.POST.get('file_id')
    if not file_id:
        return JsonResponse({'error': 'File ID is required'}, status=400)
    
    try:
        uploaded_file_obj = UploadedExcelFile.objects.get(id=file_id)
        original_path = Path(uploaded_file_obj.file.path)
        
        # Run Workflow-4 pipeline
        result = run_workflow4_pipeline(
            str(original_path),
            processed_dir=Path(settings.MEDIA_ROOT) / 'uploads' / 'processed',
            charts_dir=Path(settings.BASE_DIR) / 'static' / 'charts',
            original_file_path=original_path,
        )
        
        # Write results back to original Excel file structure
        processed_dir = Path(settings.MEDIA_ROOT) / 'uploads' / 'processed'
        processed_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y_%m_%d_%H%M%S")
        final_output_path = processed_dir / f"processed_{file_id}_{timestamp}.xlsx"
        
        # Copy original file first
        import shutil
        shutil.copy2(original_path, final_output_path)
        
        # Write Workflow-4 results to the copied file
        write_results_to_original_excel(
            final_output_path,
            result.forecast_table,
            final_output_path,
        )
        
        # Store the final file path in session or return it
        try:
            relative_path = final_output_path.relative_to(settings.MEDIA_ROOT)
            download_url = f"{settings.MEDIA_URL}{str(relative_path).replace(os.sep, '/')}"
        except ValueError:
            download_url = f"{settings.MEDIA_URL}uploads/processed/{final_output_path.name}"
        
        # Get ingredient_list and annual_data from processed data
        processed_data = ProcessedData.objects.filter(original_file=uploaded_file_obj).first()
        ingredient_list = []
        annual_data = []
        
        if processed_data:
            data = processed_data.data
            annual_data = [{k: v for k, v in row.items() if k != 'region' and k <= 'P'} 
                          for row in data if row.get('region') in ['annual_data', 'annual_separator']][1:]
            
            # Extract ingredient data
            ingredients = ['mct360', 'mct165', 'mctstick10', 'mctstick30', 'mctstick16', 'mctitto_c']
            ingredient_list = [(ing, []) for ing in ingredients]
            for row in data:
                region = row.get('region', 'unknown')
                if region.startswith('ingredient_'):
                    ing = region.split('_', 1)[1]
                    for item in ingredient_list:
                        if item[0] == ing:
                            filtered_row = {k: v for k, v in row.items() if k != 'region' and k <= 'Q'}
                            item[1].append(filtered_row)
                            break
        
        # Build chart data from workflow4 results
        try:
            chart_data_dict = build_chart_data_from_workflow4(result, ingredient_list, annual_data)
        except Exception as e:
            logger.error("Error building chart data from workflow4: %s", str(e), exc_info=True)
            # Fallback: use empty structure
            chart_data_dict = {
                'overall': {'months': [], 'historical': [], 'predicted': []},
                'ingredients': {}
            }
        
        # Convert to JSON format for frontend
        products_chart_data = []
        
        # Helper function to generate predicted months from last historical month
        def generate_predicted_months(last_month_name, num_months=6):
            """Generate next N months following fiscal month order"""
            predicted_months = []
            if last_month_name in FISCAL_MONTHS:
                last_month_idx = FISCAL_MONTHS.index(last_month_name)
                for i in range(1, num_months + 1):
                    next_idx = (last_month_idx + i) % 12
                    predicted_months.append(FISCAL_MONTHS[next_idx])
            elif last_month_name in MONTH_NAMES:
                # Convert calendar month to fiscal month order
                # Find the fiscal month index
                last_month_idx = MONTH_NAMES.index(last_month_name)
                # Map to fiscal order: April=0, May=1, ..., March=11
                fiscal_idx = (last_month_idx - 3) % 12  # April (index 3) -> 0
                for i in range(1, num_months + 1):
                    next_fiscal_idx = (fiscal_idx + i) % 12
                    # Map back to calendar month
                    calendar_idx = (next_fiscal_idx + 3) % 12
                    predicted_months.append(MONTH_NAMES[calendar_idx])
            else:
                # Default: start from current month
                current_month = datetime.now().month - 1  # 0-indexed
                for i in range(1, num_months + 1):
                    next_idx = (current_month + i) % 12
                    predicted_months.append(MONTH_NAMES[next_idx])
            return predicted_months
        
        # Add overall data as first product
        overall_data = chart_data_dict.get('overall', {})
        if overall_data and overall_data.get('months') and overall_data.get('historical'):
            overall_historical = overall_data.get('historical', [])
            overall_predicted = overall_data.get('predicted', [])
            overall_months = overall_data.get('months', [])
            
            # Generate predicted month names
            if overall_months:
                predicted_month_names = generate_predicted_months(overall_months[-1], len(overall_predicted))
            else:
                predicted_month_names = [MONTH_NAMES[(datetime.now().month - 1 + i) % 12] for i in range(1, len(overall_predicted) + 1)]
            
            products_chart_data.append({
                'product_code': 'OVERALL',
                'sheet_name': 'Workflow 4',
                'historical': [
                    {'month': month, 'value': val} 
                    for month, val in zip(overall_months, overall_historical)
                ],
                'predicted': [
                    {'month': month, 'value': val} 
                    for month, val in zip(predicted_month_names, overall_predicted)
                ]
            })
        
        # Add ingredient data
        ingredients_data = chart_data_dict.get('ingredients', {})
        for ing_name, ing_data in ingredients_data.items():
            ing_months = ing_data.get('months', [])
            ing_historical = ing_data.get('historical', [])
            ing_predicted = ing_data.get('predicted', [])
            
            if ing_months and ing_historical:
                # Generate predicted month names
                predicted_month_names = generate_predicted_months(ing_months[-1], len(ing_predicted))
                
                products_chart_data.append({
                    'product_code': ing_name,
                    'sheet_name': 'Workflow 4',
                    'historical': [
                        {'month': month, 'value': val} 
                        for month, val in zip(ing_months, ing_historical)
                    ],
                    'predicted': [
                        {'month': month, 'value': val} 
                        for month, val in zip(predicted_month_names, ing_predicted)
                    ]
                })
        
        return JsonResponse({
            'success': True,
            'message': 'All workflows completed successfully.',
            'download_url': download_url,
            'file_id': file_id,
            'summary': {
                'products': result.summary['products'],
                'total_forecast': result.summary['total_forecast'],
                'total_raw_material': result.summary['total_raw_material'],
            },
            'chart_data': {
                'file_id': str(file_id),
                'products': products_chart_data,
                'processed_at': datetime.utcnow().isoformat() + 'Z'
            }
        })
    except UploadedExcelFile.DoesNotExist:
        return JsonResponse({'error': 'File not found'}, status=404)
    except ValueError as exc:
        return JsonResponse({'error': str(exc)}, status=400)
    except Exception as exc:
        return JsonResponse({'error': f'An error occurred: {str(exc)}'}, status=500)


def download_final_file(request, file_id):
    """
    Download the final processed Excel file with all workflow results.
    """
    try:
        uploaded_file_obj = UploadedExcelFile.objects.get(id=file_id)
        
        # Find the most recent processed file for this upload
        processed_dir = Path(settings.MEDIA_ROOT) / 'uploads' / 'processed'
        pattern = f"processed_{file_id}_*.xlsx"
        matching_files = list(processed_dir.glob(pattern))
        
        if not matching_files:
            # Fallback: try to find any processed file
            return JsonResponse({'error': 'Processed file not found'}, status=404)
        
        # Get the most recent file
        latest_file = max(matching_files, key=lambda p: p.stat().st_mtime)
        
        with open(latest_file, 'rb') as f:
            response = HttpResponse(
                f.read(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = f'attachment; filename="processed_{file_id}.xlsx"'
            return response
    except UploadedExcelFile.DoesNotExist:
        return JsonResponse({'error': 'File not found'}, status=404)
    except Exception as exc:
        return JsonResponse({'error': f'An error occurred: {str(exc)}'}, status=500)


def get_chart_data(request, file_id):
    """
    Standardized API endpoint to get chart data for a specific file.
    Returns data in the format expected by frontend charts.
    """
    try:
        uploaded_file_obj = UploadedExcelFile.objects.get(id=file_id)
        
        # Try to get processed Excel file from workflow4
        processed_dir = Path(settings.MEDIA_ROOT) / 'uploads' / 'processed'
        pattern = f"processed_{file_id}_*.xlsx"
        matching_files = list(processed_dir.glob(pattern))
        
        products_data = []
        
        if matching_files:
            # Use the most recent processed file
            latest_file = max(matching_files, key=lambda p: p.stat().st_mtime)
            logger.info("Extracting chart data from processed file: %s", latest_file)
            
            # Read the Excel file
            excel_file = pd.ExcelFile(latest_file)
            
            # Try to get data from Workflow 4 sheet
            workflow4_sheet = None
            for sheet_name in excel_file.sheet_names:
                if 'workflow' in sheet_name.lower() and '4' in sheet_name:
                    workflow4_sheet = pd.read_excel(latest_file, sheet_name=sheet_name)
                    break
            
            if workflow4_sheet is not None and not workflow4_sheet.empty:
                # Extract product codes from Workflow 4 sheet
                if 'Product' in workflow4_sheet.columns:
                    product_codes = workflow4_sheet['Product'].dropna().unique()
                    
                    for product_code in product_codes:
                        product_str = str(product_code).strip()
                        if not product_str:
                            continue
                        
                        # Get historical data from monthly_trend if available
                        # Otherwise extract from original Excel
                        historical = []
                        predicted = []
                        
                        # Try to get from workflow4 monthly_trend
                        # For now, extract from the original file structure
                        try:
                            # Read original file to get historical data
                            original_df = pd.read_excel(uploaded_file_obj.file.path, header=None)
                            original_df.columns = [chr(65 + i) for i in range(len(original_df.columns))]
                            
                            # Extract monthly series
                            monthly_series = extract_from_workflow4_sheet(
                                str(latest_file), product_str
                            )
                            
                            # Convert to historical format
                            for month_name, value in monthly_series.items():
                                # Create YYYY-MM format (using current year as base)
                                current_year = datetime.now().year
                                month_num = FISCAL_MONTHS.index(month_name) + 4  # April = 4
                                if month_num > 12:
                                    month_num -= 12
                                    year = current_year + 1
                                else:
                                    year = current_year
                                
                                historical.append({
                                    'month': f"{year}-{month_num:02d}",
                                    'value': float(value)
                                })
                            
                            # Get predicted values from Workflow 4 forecast table
                            product_row = workflow4_sheet[workflow4_sheet['Product'] == product_code]
                            if not product_row.empty:
                                forecast_demand = float(product_row.iloc[0]['Forecast Demand'])
                                
                                # Generate predicted months (next 6 months)
                                if historical:
                                    last_month = historical[-1]['month']
                                    year, month = map(int, last_month.split('-'))
                                    
                                    for i in range(1, 7):
                                        month += 1
                                        if month > 12:
                                            month = 1
                                            year += 1
                                        
                                        predicted.append({
                                            'month': f"{year}-{month:02d}",
                                            'value': forecast_demand  # Use forecast demand for all predicted months
                                        })
                            
                        except Exception as e:
                            logger.error("Error extracting data for product %s: %s", product_str, str(e))
                            continue
                        
                        products_data.append({
                            'product_code': product_str,
                            'sheet_name': 'Workflow 4',
                            'historical': historical,
                            'predicted': predicted
                        })
        
        # If no workflow4 data, try to extract from ingredient sections
        if not products_data:
            processed_data = ProcessedData.objects.filter(original_file=uploaded_file_obj).first()
            if processed_data:
                data = processed_data.data
                ingredient_list = []
                ingredients = ['mct360', 'mct165', 'mctstick10', 'mctstick30', 'mctstick16', 'mctitto_c']
                
                for ing_name in ingredients:
                    ing_rows = [row for row in data 
                               if row.get('region', '').startswith(f'ingredient_{ing_name}')]
                    if ing_rows:
                        ingredient_list.append((ing_name, ing_rows))
                
                for ing_name, rows in ingredient_list:
                    monthly_series = extract_from_ingredient_section(rows, ing_name.upper())
                    
                    historical = []
                    predicted = []
                    
                    current_year = datetime.now().year
                    for month_name, value in monthly_series.items():
                        month_num = FISCAL_MONTHS.index(month_name) + 4
                        if month_num > 12:
                            month_num -= 12
                            year = current_year + 1
                        else:
                            year = current_year
                        
                        historical.append({
                            'month': f"{year}-{month_num:02d}",
                            'value': float(value)
                        })
                    
                    # Generate predictions
                    if historical:
                        historical_values = [h['value'] for h in historical]
                        from .prediction_utils import predict_next_months
                        predictions = predict_next_months(historical_values, 6, 'moving_average')
                        
                        last_month = historical[-1]['month']
                        year, month = map(int, last_month.split('-'))
                        
                        for pred_value in predictions:
                            month += 1
                            if month > 12:
                                month = 1
                                year += 1
                            
                            predicted.append({
                                'month': f"{year}-{month:02d}",
                                'value': float(pred_value)
                            })
                    
                    products_data.append({
                        'product_code': ing_name.upper(),
                        'sheet_name': 'Ingredient Section',
                        'historical': historical,
                        'predicted': predicted
                    })
        
        return JsonResponse({
            'file_id': str(file_id),
            'products': products_data,
            'processed_at': datetime.utcnow().isoformat() + 'Z'
        })
        
    except UploadedExcelFile.DoesNotExist:
        return JsonResponse({'error': 'File not found'}, status=404)
    except Exception as exc:
        logger.error("Error in get_chart_data: %s", str(exc), exc_info=True)
        return JsonResponse({'error': f'An error occurred: {str(exc)}'}, status=500)


def strict_extract_excel(request, file_id):
    """
    STRICT Excel extraction endpoint.
    Returns EXACT values from Excel with NO assumptions.
    Uses Universal Extractor to get clean data with pure float values.
    """
    try:
        uploaded_file_obj = UploadedExcelFile.objects.get(id=file_id)
        file_path = uploaded_file_obj.file.path
        
        # Use Universal Extractor
        extractor = UniversalDataExtractor(file_path)
        extracted_data = extractor.extract()
        
        # Build chart data
        chart_builder = ChartDataBuilder()
        chart_data = chart_builder.build_chart_data(extracted_data)
        
        # Build response with clean structure
        result = {
            'error': False,
            'products': {},
            'overall': {
                'months': chart_data['overall']['months'],
                'historical': chart_data['overall']['historical'],
                'predicted': chart_data['overall']['predicted']
            },
            'summary': {
                'products': len(chart_data['products']),
                'total_forecast': sum(chart_data['overall']['predicted']) if chart_data['overall']['predicted'] else 0.0,
                'total_raw_material': sum(chart_data['overall']['historical']) if chart_data['overall']['historical'] else 0.0
            }
        }
        
        # Convert products to array format for response
        products_array = []
        for product_code, product_data in chart_data['products'].items():
            products_array.append({
                'product_code': product_code,
                'sheet_name': 'Main Sheet',
                'historical': [
                    {'month': month, 'value': float(value)}
                    for month, value in zip(product_data['months'], product_data['historical'])
                ],
                'predicted': [
                    {'month': month, 'value': float(value)}
                    for month, value in zip(product_data.get('predicted_months', []), product_data['predicted'])
                ]
            })
        
        result['products'] = products_array
        
        # Return JSON response with clean dataset
        return JsonResponse(result, json_dumps_params={'ensure_ascii': False})
        
    except UploadedExcelFile.DoesNotExist:
        return JsonResponse({
            'error': True,
            'message': 'File not found',
            'missing_items': ['file_id'],
            'products': [],
            'overall': {'months': [], 'historical': [], 'predicted': []},
            'summary': {'products': 0, 'total_forecast': 0.0, 'total_raw_material': 0.0}
        }, status=404)
    except Exception as exc:
        logger.error("Error in strict_extract_excel: %s", str(exc), exc_info=True)
        return JsonResponse({
            'error': True,
            'message': f'Error processing Excel file: {str(exc)}',
            'missing_items': [],
            'products': [],
            'overall': {'months': [], 'historical': [], 'predicted': []},
            'summary': {'products': 0, 'total_forecast': 0.0, 'total_raw_material': 0.0}
        }, status=500)
