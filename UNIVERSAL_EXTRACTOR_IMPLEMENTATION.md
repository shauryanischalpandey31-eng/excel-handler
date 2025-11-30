# Universal Excel Extractor Implementation

## Overview

Implemented a robust, universal Excel data extraction system that:
- Extracts clean numeric data from ANY Excel file structure
- Returns pure Python float values (no numpy types, no nested objects)
- Handles Japanese headers and complex structures
- Prevents server crashes with robust error handling
- Auto-detects all products (not just hardcoded list)

## Key Components

### 1. UniversalDataExtractor (`universal_extractor.py`)

**Purpose**: Extract clean numeric data from Excel files

**Features**:
- Scans all sheets automatically
- Detects product blocks dynamically (MCT360, MCT165, etc.)
- Identifies month columns (handles Japanese: 4月, English: April, etc.)
- Extracts monthly values as pure Python floats
- Calculates predictions using 3-month moving average
- Handles missing data gracefully

**Output Format**:
```python
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
```

**Key Methods**:
- `extract()`: Main extraction method
- `_calculate_predictions()`: 3-month moving average forecast
- `to_float()`: Converts any value to pure Python float

### 2. ChartDataBuilder (`chart_data_builder.py`)

**Purpose**: Convert extracted data to chart-friendly arrays

**Features**:
- Converts dict format to arrays
- Ensures all values are pure floats
- Orders data by fiscal months
- Builds Django template context

**Output Format**:
```python
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
```

**Key Methods**:
- `build_chart_data()`: Convert to chart format
- `build_template_context()`: Build Django context

### 3. Updated Views (`views.py`)

**Changes**:
- `index()` view now uses `UniversalDataExtractor` and `ChartDataBuilder`
- `strict_extract_excel()` endpoint updated to use new extractor
- All data is converted to pure floats before sending to template
- Robust error handling prevents server crashes

**Data Flow**:
1. Upload Excel file
2. `UniversalDataExtractor.extract()` → Clean dict structure
3. `ChartDataBuilder.build_chart_data()` → Chart arrays
4. `ChartDataBuilder.build_template_context()` → Django context
5. Template receives clean arrays of floats

### 4. Updated Template (`index.html`)

**Changes**:
- Table population handles pure float arrays
- No more "[object Object]" - all values are numbers
- Clean number formatting with `toLocaleString()`
- Handles missing values gracefully (shows "-")

**JavaScript Updates**:
- `renderOverallChart()` receives arrays: `months[]`, `historical[]`, `predicted[]`
- Table code converts all values to floats before display
- Proper error handling for null/undefined values

## Data Structure

### Backend → Frontend

**Before (Problematic)**:
```json
{
  "April": {"value": 17625.22}  // Nested object - causes [object Object]
}
```

**After (Fixed)**:
```json
{
  "months": ["April", "May", ...],
  "historical": [17625.22, 18200.44, ...],  // Pure numbers
  "predicted": [18000.0, 18000.0, ...]     // Pure numbers
}
```

## Error Prevention

1. **Type Conversion**: All values converted to `float()` explicitly
2. **Null Handling**: Checks for `None`, `NaN`, `null` before processing
3. **Try/Except**: Wrapped around all Excel parsing operations
4. **Empty Data**: Returns empty structures instead of crashing
5. **Memory Efficient**: No large debug prints that cause 502 errors

## Product Detection

**Auto-Detection**:
- Scans first 3 columns of each sheet
- Looks for product codes (MCT360, MCT165, etc.)
- Also detects unknown products dynamically
- Handles partial matches (e.g., "MCT360" matches "MCT360_EXTRA")

**Known Products**:
- MCT360
- MCT165
- MCTSTICK10
- MCTSTICK30
- MCTSTICK16
- MCTITTO_C

**Plus**: Any other product codes found in Excel

## Month Detection

**Supported Formats**:
- English: "April", "May", "Apr", "04"
- Japanese: "4月", "5月"
- Numbers: "1", "2", "01", "02"
- Handles trailing spaces, dots, etc.

## Prediction Algorithm

**3-Month Moving Average**:
1. If ≥3 months available: Average of last 3 months
2. If 2 months available: Average of both
3. If 1 month available: Use that value
4. If 0 months: Return empty dict

**Fallback**: Never crashes, always returns valid structure

## Testing

### Test Cases Covered:
1. ✅ Japanese Excel files with 4月 headers
2. ✅ English Excel files with "April" headers
3. ✅ Files with all 6 products
4. ✅ Files with missing products
5. ✅ Files with extra products
6. ✅ Files with missing months
7. ✅ Files with merged cells
8. ✅ Large files (memory efficient)

### Expected Results:
- ✅ All products detected automatically
- ✅ Charts show correct numeric values
- ✅ No "[object Object]" anywhere
- ✅ Server doesn't crash (no 502/500 errors)
- ✅ Tables display formatted numbers correctly

## Deployment

Code has been pushed to GitHub and will auto-deploy on Render.

**Files Changed**:
- `excel_handler/universal_extractor.py` (NEW)
- `excel_handler/chart_data_builder.py` (UPDATED)
- `excel_handler/views.py` (UPDATED)
- `excel_handler/templates/excel_handler/index.html` (UPDATED)

## Usage

The system works automatically:
1. User uploads Excel file
2. Backend extracts data using `UniversalDataExtractor`
3. Data is converted to chart format using `ChartDataBuilder`
4. Template receives clean arrays
5. Charts render with correct numeric values

No configuration needed - works with any Excel file structure!

