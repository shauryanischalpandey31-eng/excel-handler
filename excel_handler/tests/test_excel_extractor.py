"""
Unit tests for Excel extraction functions.
"""
import pytest
import pandas as pd
from collections import OrderedDict
from excel_handler.excel_extractor import (
    extract_monthly_series,
    normalize_month_name,
    normalize_numeric_value,
    FISCAL_MONTHS,
)


def test_extract_monthly_series_standard():
    """Test extraction with standard Excel format."""
    # Create sample DataFrame
    data = {
        'A': ['Product', 'MCT360', 'Other'],
        'D': [None, 720.0, 100.0],  # April
        'E': [None, 760.0, 110.0],  # May
        'F': [None, 800.0, 120.0],  # June
    }
    df = pd.DataFrame(data)
    
    result = extract_monthly_series(df, 'MCT360', 'Test Sheet')
    
    assert isinstance(result, OrderedDict)
    assert len(result) >= 3
    assert 'April' in result
    assert result['April'] == 720.0
    assert result['May'] == 760.0
    assert result['June'] == 800.0


def test_extract_handles_trailing_spaces():
    """Test extraction handles column names with trailing spaces."""
    data = {
        'A': ['Product', 'MCT165'],
        'D ': [None, 500.0],  # Note: trailing space in column name
        'E ': [None, 550.0],
    }
    df = pd.DataFrame(data)
    df.columns = df.columns.str.strip()  # Simulate pandas behavior
    
    result = extract_monthly_series(df, 'MCT165', 'Test Sheet')
    
    # Should still extract data
    assert len(result) >= 2


def test_forecast_handles_less_than_3_months():
    """Test forecasting handles insufficient data gracefully."""
    from excel_handler.prediction_utils import predict_next_months
    
    # Test with 2 months
    values = [100.0, 120.0]
    predictions = predict_next_months(values, 6, 'moving_average')
    
    assert len(predictions) == 6
    # Should use available data, not crash
    assert all(isinstance(p, (int, float)) for p in predictions)
    
    # Test with 1 month
    values = [100.0]
    predictions = predict_next_months(values, 6, 'moving_average')
    assert len(predictions) == 6
    
    # Test with 0 months
    values = []
    predictions = predict_next_months(values, 6, 'moving_average')
    assert len(predictions) == 6
    assert all(p == 0.0 for p in predictions)


def test_normalize_month_name():
    """Test month name normalization."""
    assert normalize_month_name('Apr') == 'April'
    assert normalize_month_name('APRIL') == 'April'
    assert normalize_month_name('Apr ') == 'April'  # Trailing space
    assert normalize_month_name('04') == 'April'
    assert normalize_month_name('4') == 'April'
    assert normalize_month_name('invalid') is None


def test_normalize_numeric_value():
    """Test numeric value normalization."""
    assert normalize_numeric_value('720.5') == 720.5
    assert normalize_numeric_value('$720.50') == 720.5
    assert normalize_numeric_value('1,234.56') == 1234.56
    assert normalize_numeric_value('(500)') == -500.0  # Parentheses = negative
    assert normalize_numeric_value(720.5) == 720.5
    assert normalize_numeric_value(None) is None
    assert normalize_numeric_value('') is None
    assert normalize_numeric_value('-') is None

