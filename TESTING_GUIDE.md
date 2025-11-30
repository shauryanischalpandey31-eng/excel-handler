# Testing Guide for Chart Data Fix

## Quick Test Steps

### 1. Upload Excel File
- Go to `/excel/`
- Upload an Excel file with monthly data
- File should have columns D-O representing months April-March

### 2. Process the File
- Click "Process" button
- Wait for "Adjust Annual Data" workflow to complete
- Click "Adjust Annual Data" button
- This triggers `process_all_workflows()` which runs Workflow 4

### 3. Verify Chart Data
- Open browser DevTools → Network tab
- Look for response from `process_all_workflows` endpoint
- Check `chart_data` field in JSON response
- Verify structure matches:
  ```json
  {
    "chart_data": {
      "file_id": "123",
      "products": [
        {
          "product_code": "MCT360",
          "sheet_name": "Workflow 4",
          "historical": [{"month": "2025-04", "value": 720.0}, ...],
          "predicted": [{"month": "2025-10", "value": 646.0}, ...]
        }
      ],
      "processed_at": "2025-11-30T09:00:00Z"
    }
  }
  ```

### 4. Check Charts
- Charts should appear automatically after processing
- Hover over historical points → tooltip shows exact Excel value
- Hover over predicted points → tooltip shows forecast value
- Check table below chart → values match chart points

### 5. Verify Excel Output
- Download processed Excel file
- Open "Workflow 4" sheet
- Verify forecast values match chart predictions

## Test Cases

### Test Case 1: Standard File
- File with 12 months of data (April-March)
- Expected: All 12 months extracted, 6 months predicted

### Test Case 2: File with Trailing Spaces
- Column headers like "Apr " (with space)
- Expected: Still extracts correctly

### Test Case 3: Less than 3 Months
- File with only 2 months of data
- Expected: Forecast still works (uses available data)

### Test Case 4: Missing Months
- File with gaps in monthly data
- Expected: Missing months show as NaN, not zero

## Debugging

### Check Backend Logs
Look for extraction logs:
```
DEBUG: Extracting series for product 'MCT360' from sheet 'Workflow 4'
DEBUG: Found product 'MCT360' at row 1
DEBUG: Extracted 12 months for 'MCT360': first=April (720.00), last=March (850.00)
```

### Check Browser Console
- No JavaScript errors
- Chart data logged: `console.log(chartData)`
- Verify `renderChartsFromData()` is called

### Check Network Tab
- Response status: 200
- `chart_data.products` array has data
- Historical and predicted arrays populated

## Common Issues

### Issue: Charts not showing
- **Fix**: Check if `renderChartsFromData()` is called after processing
- **Fix**: Verify chart container exists: `#monthlyForecastChart`

### Issue: Wrong values in charts
- **Fix**: Check extraction logs for parsed values
- **Fix**: Verify Excel file structure matches expected format
- **Fix**: Check if month names are being normalized correctly

### Issue: Tooltips show "N/A"
- **Fix**: Verify values are numeric, not strings
- **Fix**: Check `context.parsed.y` is not null

## Sample Test Files

### File 1: Standard Format
- Row 1: Headers (A=Product, D=April, E=May, ..., O=March)
- Row 2: MCT360, 720, 760, 800, ... (monthly values)

### File 2: With Trailing Spaces
- Headers: "Apr ", "May ", etc.
- Should still extract correctly

### File 3: Minimal Data
- Only 2 months: April=720, May=760
- Should still generate forecast

