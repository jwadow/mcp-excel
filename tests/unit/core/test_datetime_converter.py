# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Unit tests for DateTimeConverter component.

Tests cover:
- Single value conversion (Excel number to datetime)
- Column conversion (vectorized)
- Epoch detection (Windows vs Mac)
- Edge cases (NaN, negative values, boundary values)
"""

import pytest
import pandas as pd
from datetime import datetime


def test_convert_single_value_windows_epoch(datetime_converter):
    """Test converting single Excel number with Windows epoch.
    
    Verifies:
    - Correct conversion for Windows epoch (1899-12-30)
    - Returns pd.Timestamp object
    - Date and time components are correct
    """
    print(f"\n📂 Testing single value conversion (Windows epoch)")
    
    # Excel number 44927.0 = 2023-01-01 00:00:00 (Windows epoch)
    excel_number = 44927.0
    
    # Act
    result = datetime_converter.convert_excel_number_to_datetime(excel_number, epoch="windows")
    
    # Assert
    print(f"✅ Excel {excel_number} -> {result}")
    
    assert isinstance(result, pd.Timestamp), "Should return pd.Timestamp"
    assert result.year == 2023, "Year should be 2023"
    assert result.month == 1, "Month should be 1"
    assert result.day == 1, "Day should be 1"


def test_convert_single_value_with_time(datetime_converter):
    """Test converting Excel number with time component.
    
    Verifies:
    - Fractional part represents time
    - Time components are correct
    """
    print(f"\n📂 Testing single value with time component")
    
    # Excel number 44927.5 = 2023-01-01 12:00:00
    excel_number = 44927.5
    
    # Act
    result = datetime_converter.convert_excel_number_to_datetime(excel_number, epoch="windows")
    
    # Assert
    print(f"✅ Excel {excel_number} -> {result}")
    
    assert result.year == 2023, "Year should be 2023"
    assert result.month == 1, "Month should be 1"
    assert result.day == 1, "Day should be 1"
    assert result.hour == 12, "Hour should be 12"
    assert result.minute == 0, "Minute should be 0"


def test_convert_single_value_mac_epoch(datetime_converter):
    """Test converting single Excel number with Mac epoch.
    
    Verifies:
    - Correct conversion for Mac epoch (1904-01-01)
    - Different result than Windows epoch
    """
    print(f"\n📂 Testing single value conversion (Mac epoch)")
    
    # Same Excel number, different epoch
    excel_number = 44927.0
    
    # Act
    result_windows = datetime_converter.convert_excel_number_to_datetime(excel_number, epoch="windows")
    result_mac = datetime_converter.convert_excel_number_to_datetime(excel_number, epoch="mac")
    
    # Assert
    print(f"✅ Windows: {result_windows}")
    print(f"   Mac: {result_mac}")
    
    assert result_windows != result_mac, "Windows and Mac epochs should give different results"
    # Mac epoch is 4 years + 1 day later than Windows
    assert result_mac.year == result_windows.year + 4, "Mac result should be ~4 years later"


def test_convert_nan_value(datetime_converter):
    """Test converting NaN value.
    
    Verifies:
    - NaN input returns pd.NaT (Not a Time)
    - Doesn't crash on missing values
    """
    print(f"\n📂 Testing NaN value conversion")
    
    # Act
    result = datetime_converter.convert_excel_number_to_datetime(float('nan'), epoch="windows")
    
    # Assert
    print(f"✅ NaN -> {result}")
    
    assert pd.isna(result), "NaN should convert to NaT"
    assert isinstance(result, type(pd.NaT)), "Should return pd.NaT type"


def test_convert_column_windows_epoch(datetime_converter):
    """Test converting entire column with Windows epoch.
    
    Verifies:
    - Vectorized conversion works
    - All values converted correctly
    - Returns Series with datetime64 dtype
    """
    print(f"\n📂 Testing column conversion (Windows epoch)")
    
    # Create Series with Excel date numbers
    excel_numbers = pd.Series([44927.0, 44928.0, 44929.0, 44930.0, 44931.0])
    
    # Act
    result = datetime_converter.convert_column(excel_numbers, epoch="windows")
    
    # Assert
    print(f"✅ Converted {len(result)} values")
    print(f"   First: {result.iloc[0]}")
    print(f"   Last: {result.iloc[-1]}")
    print(f"   Dtype: {result.dtype}")
    
    assert len(result) == len(excel_numbers), "Should convert all values"
    assert pd.api.types.is_datetime64_any_dtype(result), "Should have datetime64 dtype"
    
    # Check first and last values
    assert result.iloc[0].year == 2023, "First value should be 2023"
    assert result.iloc[-1].year == 2023, "Last value should be 2023"
    assert result.iloc[-1].day == 5, "Last value should be day 5"


def test_convert_column_with_nan(datetime_converter):
    """Test converting column with NaN values.
    
    Verifies:
    - NaN values are preserved as NaT
    - Other values converted correctly
    """
    print(f"\n📂 Testing column conversion with NaN")
    
    # Create Series with NaN
    excel_numbers = pd.Series([44927.0, float('nan'), 44929.0, float('nan'), 44931.0])
    
    # Act
    result = datetime_converter.convert_column(excel_numbers, epoch="windows")
    
    # Assert
    print(f"✅ Converted {len(result)} values")
    print(f"   NaN count: {result.isna().sum()}")
    
    assert result.isna().sum() == 2, "Should have 2 NaT values"
    assert not pd.isna(result.iloc[0]), "First value should not be NaT"
    assert pd.isna(result.iloc[1]), "Second value should be NaT"


def test_convert_column_mac_epoch(datetime_converter):
    """Test converting column with Mac epoch.
    
    Verifies:
    - Mac epoch conversion works for columns
    - Results differ from Windows epoch
    """
    print(f"\n📂 Testing column conversion (Mac epoch)")
    
    excel_numbers = pd.Series([44927.0, 44928.0, 44929.0])
    
    # Act
    result_windows = datetime_converter.convert_column(excel_numbers, epoch="windows")
    result_mac = datetime_converter.convert_column(excel_numbers, epoch="mac")
    
    # Assert
    print(f"✅ Windows first: {result_windows.iloc[0]}")
    print(f"   Mac first: {result_mac.iloc[0]}")
    
    assert not result_windows.equals(result_mac), "Windows and Mac results should differ"
    # Mac dates should be ~4 years later
    assert result_mac.iloc[0].year == result_windows.iloc[0].year + 4, "Mac should be ~4 years later"


def test_detect_epoch_windows(datetime_converter):
    """Test epoch detection for Windows-based dates.
    
    Verifies:
    - Detects Windows epoch for typical date values
    - Returns "windows" string
    """
    print(f"\n📂 Testing epoch detection (Windows)")
    
    # Typical Windows Excel dates (large numbers)
    series = pd.Series([44927.0, 44928.0, 44929.0, 44930.0])
    
    # Act
    detected_epoch = datetime_converter.detect_epoch(series)
    
    # Assert
    print(f"✅ Detected epoch: {detected_epoch}")
    
    assert detected_epoch == "windows", "Should detect Windows epoch for large values"


def test_detect_epoch_mac(datetime_converter):
    """Test epoch detection for Mac-based dates.
    
    Verifies:
    - Detects Mac epoch for small date values
    - Returns "mac" string
    """
    print(f"\n📂 Testing epoch detection (Mac)")
    
    # Small values suggest Mac epoch (< 1462 = 4 years)
    series = pd.Series([100.0, 200.0, 300.0, 400.0])
    
    # Act
    detected_epoch = datetime_converter.detect_epoch(series)
    
    # Assert
    print(f"✅ Detected epoch: {detected_epoch}")
    
    assert detected_epoch == "mac", "Should detect Mac epoch for small values"


def test_detect_epoch_empty_series(datetime_converter):
    """Test epoch detection with empty series.
    
    Verifies:
    - Returns default "windows" for empty series
    - Doesn't crash on edge case
    """
    print(f"\n📂 Testing epoch detection with empty series")
    
    # Empty series
    series = pd.Series([], dtype=float)
    
    # Act
    detected_epoch = datetime_converter.detect_epoch(series)
    
    # Assert
    print(f"✅ Detected epoch: {detected_epoch}")
    
    assert detected_epoch == "windows", "Should default to Windows for empty series"


def test_detect_epoch_all_nan(datetime_converter):
    """Test epoch detection with all-NaN series.
    
    Verifies:
    - Returns default "windows" for all-NaN
    - Handles missing data gracefully
    """
    print(f"\n📂 Testing epoch detection with all-NaN series")
    
    # All NaN series
    series = pd.Series([float('nan'), float('nan'), float('nan')])
    
    # Act
    detected_epoch = datetime_converter.detect_epoch(series)
    
    # Assert
    print(f"✅ Detected epoch: {detected_epoch}")
    
    assert detected_epoch == "windows", "Should default to Windows for all-NaN series"


def test_boundary_value_conversion(datetime_converter):
    """Test conversion of boundary values.
    
    Verifies:
    - Very early dates (close to epoch)
    - Very late dates (far from epoch)
    - Edge cases don't crash
    """
    print(f"\n📂 Testing boundary value conversion")
    
    # Test boundary values
    test_cases = [
        (1.0, "Very early date (1900-01-01)"),
        (60000.0, "Very late date (~2064)"),
        (0.5, "Midnight time only"),
    ]
    
    for excel_number, description in test_cases:
        result = datetime_converter.convert_excel_number_to_datetime(excel_number, epoch="windows")
        print(f"   {description}: {excel_number} -> {result}")
        assert isinstance(result, pd.Timestamp), f"Should convert {description}"
    
    print(f"✅ All boundary values converted successfully")


def test_real_world_with_dates_fixture(with_dates_fixture, file_loader, datetime_converter):
    """Test conversion on real fixture with dates.
    
    Verifies:
    - Works with actual Excel file data
    - Converts multiple columns correctly
    """
    print(f"\n📂 Testing on real fixture: {with_dates_fixture.name}")
    
    # Load without datetime conversion to get raw numbers
    df = file_loader.load(with_dates_fixture.path_str, with_dates_fixture.sheet_name, header_row=0, convert_dates=False)
    
    # Find float columns (potential dates)
    float_cols = [col for col in df.columns if pd.api.types.is_float_dtype(df[col])]
    
    print(f"   Float columns found: {len(float_cols)}")
    
    if float_cols:
        # Try converting first float column
        test_col = float_cols[0]
        print(f"   Testing column: '{test_col}'")
        
        # Act
        converted = datetime_converter.convert_column(df[test_col], epoch="windows")
        
        # Assert
        print(f"✅ Converted {len(converted)} values")
        print(f"   First value: {converted.iloc[0]}")
        print(f"   Dtype: {converted.dtype}")
        
        assert pd.api.types.is_datetime64_any_dtype(converted), "Should have datetime dtype"
        assert len(converted) == len(df[test_col]), "Should convert all values"
    else:
        print(f"   ℹ️ No float columns found (dates already converted)")


def test_epoch_constants(datetime_converter):
    """Test that epoch constants are correctly defined.
    
    Verifies:
    - EXCEL_EPOCH_WINDOWS is 1899-12-30
    - EXCEL_EPOCH_MAC is 1904-01-01
    """
    print(f"\n📂 Testing epoch constants")
    
    # Check Windows epoch
    windows_epoch = datetime_converter.EXCEL_EPOCH_WINDOWS
    print(f"   Windows epoch: {windows_epoch}")
    assert windows_epoch.year == 1899, "Windows epoch year should be 1899"
    assert windows_epoch.month == 12, "Windows epoch month should be 12"
    assert windows_epoch.day == 30, "Windows epoch day should be 30"
    
    # Check Mac epoch
    mac_epoch = datetime_converter.EXCEL_EPOCH_MAC
    print(f"   Mac epoch: {mac_epoch}")
    assert mac_epoch.year == 1904, "Mac epoch year should be 1904"
    assert mac_epoch.month == 1, "Mac epoch month should be 1"
    assert mac_epoch.day == 1, "Mac epoch day should be 1"
    
    print(f"✅ Epoch constants are correct")
