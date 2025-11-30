# Strict Excel Data Extraction System

## Overview

This system implements **STRICT Excel data extraction** with:
- ✅ **NO guessing** - Only extracts values that exist in Excel
- ✅ **NO hallucination** - Never invents or assumes values
- ✅ **NO random numbers** - All data comes from actual Excel cells
- ✅ **Universal support** - Works with ANY Excel file structure

## Key Features

### 1. Structure Validation
Before processing, the system validates:
- Excel file is readable
- Sheets exist
- Month columns are detected
- Products/ingredients are found
- Returns clear error messages if structure doesn't match

### 2. Exact Value Extraction
- Reads **EXACT** values from Excel cells
- No modification or "fixing" of values
- If a cell is empty → returns `null`
- If a value is invalid → returns `null` (not a guessed value)

### 3. Universal Detection
- **Dynamically detects:**
  - Number of products (not hardcoded)
  - Sheet names (works with any sheet names)
  - Month names (handles variations: "Apr", "April", "04", etc.)
  - Historical data length (adapts to available data)
  - Missing sheets/columns (reports what's missing)

### 4. Strict Forecasting
- Uses **3-month moving average** algorithm
- **ONLY** uses historical values from Excel
- If < 3 months available: uses available months
- If 1 month available: uses that value
- If 0 months: returns `null` (no assumptions)

## API Endpoint

### GET `/excel/strict-extract/<file_id>/`

Returns JSON response with exact extracted data.

**Success Response:**
```json
{
  "error": false,
  "products": [
    {
      "product_code": "MCT360",
      "sheet_name": "Main Sheet",
      "historical": [
        {"month": "April", "value": 114266},
        {"month": "May", "value": 125000},
        ...
      ],
      "predicted": [
        {"month": "April_next", "value": 120000},
        ...
      ]
    },
    ...
  ],
  "overall": {
    "months": ["April", "May", ...],
    "historical": [114266, 125000, ...],
    "predicted": [120000, ...]
  },
  "summary": {
    "products": 7,
    "total_forecast": 840000,
    "total_raw_material": 1500000
  }
}
```

**Error Response:**
```json
{
  "error": true,
  "message": "Missing or incorrect data format in Excel. Please upload correct template.",
  "missing_items": ["monthly_data", "products"]
}
```

## Usage

### 1. Upload Excel File
```python
# File is uploaded via the main index view
POST /excel/
```

### 2. Get Strict Extraction
```python
# After upload, get file_id from response
GET /excel/strict-extract/<file_id>/
```

### 3. Use the Data
The JSON response contains:
- **products**: Array of all detected products with historical and predicted data
- **overall**: Aggregated monthly data across all products
- **summary**: Statistics about the extraction

## Data Flow

1. **Upload** → Excel file saved to server
2. **Validate** → Check structure, sheets, columns
3. **Detect** → Find products, months, data rows
4. **Extract** → Read exact values from cells
5. **Forecast** → Calculate predictions from historical data only
6. **Return** → JSON with all extracted data

## Supported Excel Structures

### Single-Sheet Structure
- Annual data section (rows with monthly values)
- Ingredient sections (MCT360, MCT165, etc.)
- Uses `parse_excel_regions` logic

### Multi-Sheet Structure
- Each sheet can contain products
- Products detected by product codes in first columns
- Month columns detected automatically

### Universal Structure
- Works with any sheet names
- Works with any product codes
- Works with any month format (April, Apr, 04, etc.)

## Error Handling

### Missing Data
- Empty cells → `null` in JSON
- Missing months → Not included in arrays
- Missing products → Not in products array

### Invalid Structure
- Returns `error: true`
- Lists missing items in `missing_items` array
- Provides clear error message

## Example: Using in Frontend

```javascript
// After file upload
const fileId = uploadedFile.id;

// Get strict extraction
fetch(`/excel/strict-extract/${fileId}/`)
  .then(response => response.json())
  .then(data => {
    if (data.error) {
      console.error('Error:', data.message);
      console.error('Missing:', data.missing_items);
    } else {
      // Use exact data
      console.log('Products:', data.products);
      console.log('Overall:', data.overall);
      
      // Render charts with real data
      renderCharts(data.overall, data.products);
    }
  });
```

## Testing

Test with your Excel file:
```
C.【2025年10月製造計画表】_全部品 (1).xlsx
```

The system will:
1. Detect all products automatically
2. Extract exact monthly values
3. Generate forecasts from real data
4. Return complete JSON structure

## Key Differences from Previous System

| Previous | Strict System |
|----------|---------------|
| Assumed values if missing | Returns `null` |
| Hardcoded product list | Auto-detects all products |
| Fixed sheet names | Works with any sheet names |
| Guessed missing months | Only includes existing months |
| Random/placeholder values | Only real Excel values |

## Files

- `excel_handler/strict_excel_extractor.py` - Main extraction logic
- `excel_handler/views.py` - API endpoint (`strict_extract_excel`)
- `excel_handler/urls.py` - URL routing

## Next Steps

1. Test with your Excel file
2. Verify extracted values match Excel exactly
3. Check forecasts are calculated from real data
4. Integrate with frontend charts

