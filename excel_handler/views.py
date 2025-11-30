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
    
    # Convert forecast_data to JSON for template
    forecast_data_json = json.dumps(forecast_data) if forecast_data else None
    
    return render(request, 'excel_handler/index.html', {
        'data': data, 
        'annual_data': annual_data, 
        'ingredient_list': ingredient_list, 
        'uploaded_file': uploaded_file,
        'error_message': error_message,
        'warning_message': warning_message,
        'chart_data_json': chart_data_json,
        'ingredient_charts_json': ingredient_charts_json,
        'forecast_data': forecast_data_json,
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
        
        return JsonResponse({
            'success': True,
            'message': 'All workflows completed successfully.',
            'download_url': download_url,
            'file_id': file_id,
            'summary': {
                'products': result.summary['products'],
                'total_forecast': result.summary['total_forecast'],
                'total_raw_material': result.summary['total_raw_material'],
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
    API endpoint to get chart data for a specific file.
    """
    try:
        uploaded_file_obj = UploadedExcelFile.objects.get(id=file_id)
        processed_data = ProcessedData.objects.filter(original_file=uploaded_file_obj).first()
        
        if not processed_data:
            return JsonResponse({'error': 'No processed data found'}, status=404)
        
        data = processed_data.data
        annual_data = [{k: v for k, v in row.items() if k != 'region' and k <= 'P'} 
                       for row in data if row.get('region') in ['annual_data', 'annual_separator']][1:]
        
        # Extract and prepare chart data
        monthly_data = extract_monthly_data_from_annual(annual_data)
        monthly_totals = calculate_monthly_totals(monthly_data)
        forecast_data = generate_forecast_data(monthly_totals, num_future_months=6)
        chart_data = prepare_chart_data(monthly_totals, forecast_data)
        
        return JsonResponse({
            'success': True,
            'chart_data': chart_data,
            'forecast_data': forecast_data
        })
    except UploadedExcelFile.DoesNotExist:
        return JsonResponse({'error': 'File not found'}, status=404)
    except Exception as exc:
        return JsonResponse({'error': f'An error occurred: {str(exc)}'}, status=500)
