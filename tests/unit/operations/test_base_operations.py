# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Unit tests for BaseOperations class.

Tests cover:
- Response size validation (context overflow protection)
- Column limit application (smart defaults)
- Row limit enforcement (hard caps)
- Numeric column conversion (text-to-number)
"""

import pytest
import pandas as pd
from pydantic import BaseModel

from mcp_excel.operations.base import (
    BaseOperations,
    DEFAULT_COLUMN_LIMIT,
    DEFAULT_ROW_LIMIT,
    MAX_ROW_LIMIT,
    MAX_RESPONSE_CHARS,
)


# ============================================================================
# Test Response Size Validation (Context Overflow Protection)
# ============================================================================

def test_validate_response_size_within_limit(file_loader):
    """Test _validate_response_size with response within safe limit.
    
    Verifies:
    - No error raised for small responses
    - Normal operation continues
    """
    print("\n📏 Testing response size validation - within limit")
    
    ops = BaseOperations(file_loader)
    
    # Create small response (well within limit)
    class SmallResponse(BaseModel):
        data: str = "Small data"
        count: int = 10
    
    response = SmallResponse()
    
    # Should not raise any error
    try:
        ops._validate_response_size(response)
        print("  ✅ Small response passed validation")
    except ValueError:
        pytest.fail("Should not raise error for small response")


def test_validate_response_size_exceeds_limit(file_loader):
    """Test _validate_response_size with response exceeding limit.
    
    Verifies:
    - ValueError raised when response too large
    - Error message contains character count
    - Error message mentions MCP philosophy
    """
    print("\n📏 Testing response size validation - exceeds limit")
    
    ops = BaseOperations(file_loader)
    
    # Create large response (exceeds MAX_RESPONSE_CHARS)
    class LargeResponse(BaseModel):
        data: str = "x" * (MAX_RESPONSE_CHARS + 1000)
    
    response = LargeResponse()
    
    # Should raise ValueError
    with pytest.raises(ValueError) as exc_info:
        ops._validate_response_size(response)
    
    error_msg = str(exc_info.value)
    print(f"  ✅ Caught expected error")
    print(f"     Error message (first 200 chars): {error_msg[:200]}...")
    
    # Verify error message content
    assert "Response too large" in error_msg, "Should mention response size"
    # Number is formatted with comma: "10,000" not "10000"
    assert "10,000" in error_msg or str(MAX_RESPONSE_CHARS) in error_msg, "Should mention limit"
    assert "MCP Philosophy" in error_msg, "Should explain MCP philosophy"


def test_validate_response_size_error_message_with_rows_columns(file_loader):
    """Test error message includes row/column counts when provided.
    
    Verifies:
    - Error message shows current request dimensions
    - Helps user understand what caused overflow
    """
    print("\n📏 Testing response size error message with dimensions")
    
    ops = BaseOperations(file_loader)
    
    # Create large response
    class LargeResponse(BaseModel):
        data: str = "x" * (MAX_RESPONSE_CHARS + 1000)
    
    response = LargeResponse()
    
    # Provide row/column counts
    with pytest.raises(ValueError) as exc_info:
        ops._validate_response_size(
            response,
            rows_count=1000,
            columns_count=50
        )
    
    error_msg = str(exc_info.value)
    print(f"  ✅ Error message includes dimensions")
    print(f"     Message snippet: ...{error_msg[100:300]}...")
    
    # Verify dimensions are mentioned
    assert "1000 rows" in error_msg or "1,000 rows" in error_msg, "Should mention row count"
    assert "50 columns" in error_msg, "Should mention column count"


def test_validate_response_size_error_message_with_request_limit(file_loader):
    """Test error message includes suggestions based on request limit.
    
    Verifies:
    - Suggests reducing limit parameter
    - Provides specific recommended value
    - Mentions atomic operations alternative
    """
    print("\n📏 Testing response size error message with suggestions")
    
    ops = BaseOperations(file_loader)
    
    # Create large response
    class LargeResponse(BaseModel):
        data: str = "x" * (MAX_RESPONSE_CHARS + 1000)
    
    response = LargeResponse()
    
    # Provide request limit > DEFAULT_ROW_LIMIT
    with pytest.raises(ValueError) as exc_info:
        ops._validate_response_size(
            response,
            rows_count=500,
            columns_count=20,
            request_limit=500
        )
    
    error_msg = str(exc_info.value)
    print(f"  ✅ Error message includes actionable suggestions")
    print(f"     Suggestions section: ...{error_msg[300:500]}...")
    
    # Verify suggestions are present
    assert "How to fix" in error_msg, "Should have suggestions section"
    assert "Reduce 'limit' parameter" in error_msg or "limit" in error_msg.lower(), "Should suggest reducing limit"
    assert "atomic operations" in error_msg.lower(), "Should mention atomic operations"


def test_validate_response_size_suggests_fewer_columns(file_loader):
    """Test error message suggests reducing columns when many columns present.
    
    Verifies:
    - Suggests specifying fewer columns
    - Mentions default column limit
    """
    print("\n📏 Testing response size suggests fewer columns")
    
    ops = BaseOperations(file_loader)
    
    # Create large response
    class LargeResponse(BaseModel):
        data: str = "x" * (MAX_RESPONSE_CHARS + 1000)
    
    response = LargeResponse()
    
    # Provide many columns
    with pytest.raises(ValueError) as exc_info:
        ops._validate_response_size(
            response,
            rows_count=100,
            columns_count=50  # More than DEFAULT_COLUMN_LIMIT
        )
    
    error_msg = str(exc_info.value)
    print(f"  ✅ Error message suggests reducing columns")
    
    # Verify column suggestion
    assert "fewer columns" in error_msg.lower() or "columns" in error_msg, "Should suggest fewer columns"
    assert str(DEFAULT_COLUMN_LIMIT) in error_msg, "Should mention default column limit"


# ============================================================================
# Test Column Limit Application (Smart Defaults)
# ============================================================================

def test_apply_column_limit_none_columns(simple_fixture, file_loader):
    """Test _apply_column_limit with no columns specified (apply default).
    
    Verifies:
    - Returns only first DEFAULT_COLUMN_LIMIT columns
    - Returns correct column names
    """
    print("\n📊 Testing column limit - no columns specified")
    
    ops = BaseOperations(file_loader)
    
    # Load fixture with many columns
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    
    # Apply limit with None (should use default)
    limited_df, actual_columns = ops._apply_column_limit(df, None)
    
    print(f"  Original columns: {len(df.columns)}")
    print(f"  Limited columns: {len(actual_columns)}")
    print(f"  Columns returned: {actual_columns}")
    
    assert len(actual_columns) == min(DEFAULT_COLUMN_LIMIT, len(df.columns)), \
        f"Should return max {DEFAULT_COLUMN_LIMIT} columns"
    assert actual_columns == list(df.columns[:DEFAULT_COLUMN_LIMIT]), \
        "Should return first N columns"
    assert len(limited_df.columns) == len(actual_columns), \
        "DataFrame should have same columns as returned list"


def test_apply_column_limit_empty_list(simple_fixture, file_loader):
    """Test _apply_column_limit with empty list (apply default).
    
    Verifies:
    - Empty list treated same as None
    - Returns default column limit
    """
    print("\n📊 Testing column limit - empty list")
    
    ops = BaseOperations(file_loader)
    
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    
    # Apply limit with empty list
    limited_df, actual_columns = ops._apply_column_limit(df, [])
    
    print(f"  Limited columns: {len(actual_columns)}")
    
    assert len(actual_columns) == min(DEFAULT_COLUMN_LIMIT, len(df.columns)), \
        "Empty list should use default limit"


def test_apply_column_limit_requested_columns(simple_fixture, file_loader):
    """Test _apply_column_limit with specific columns requested.
    
    Verifies:
    - Returns exactly requested columns
    - No default limit applied
    """
    print("\n📊 Testing column limit - specific columns requested")
    
    ops = BaseOperations(file_loader)
    
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    
    # Request specific columns
    requested = [simple_fixture.columns[0], simple_fixture.columns[1]]
    limited_df, actual_columns = ops._apply_column_limit(df, requested)
    
    print(f"  Requested: {requested}")
    print(f"  Returned: {actual_columns}")
    
    assert actual_columns == requested, "Should return exactly requested columns"
    assert len(limited_df.columns) == len(requested), "DataFrame should have requested columns"


def test_apply_column_limit_more_than_default(wide_table_fixture, file_loader):
    """Test _apply_column_limit when requesting more than default limit.
    
    Verifies:
    - Can request more than DEFAULT_COLUMN_LIMIT if explicitly specified
    - No artificial cap on explicit requests
    """
    print("\n📊 Testing column limit - request more than default")
    
    ops = BaseOperations(file_loader)
    
    df = file_loader.load(wide_table_fixture.path_str, wide_table_fixture.sheet_name, header_row=0)
    
    # Request 10 columns (more than DEFAULT_COLUMN_LIMIT=5)
    requested = list(df.columns[:10])
    limited_df, actual_columns = ops._apply_column_limit(df, requested)
    
    print(f"  Requested: {len(requested)} columns")
    print(f"  Returned: {len(actual_columns)} columns")
    
    assert len(actual_columns) == 10, "Should allow explicit request > default"
    assert actual_columns == requested, "Should return all requested columns"


# ============================================================================
# Test Row Limit Enforcement (Hard Caps)
# ============================================================================

def test_enforce_row_limit_within_max(file_loader):
    """Test _enforce_row_limit with limit within MAX_ROW_LIMIT.
    
    Verifies:
    - Returns requested limit unchanged
    - No capping applied
    """
    print("\n📏 Testing row limit enforcement - within max")
    
    ops = BaseOperations(file_loader)
    
    # Request limit within max
    requested_limit = 500
    enforced_limit = ops._enforce_row_limit(requested_limit)
    
    print(f"  Requested: {requested_limit}")
    print(f"  Enforced: {enforced_limit}")
    
    assert enforced_limit == requested_limit, "Should return requested limit unchanged"


def test_enforce_row_limit_exceeds_max(file_loader):
    """Test _enforce_row_limit with limit exceeding MAX_ROW_LIMIT.
    
    Verifies:
    - Returns MAX_ROW_LIMIT instead of requested
    - Protects against excessive requests
    """
    print("\n📏 Testing row limit enforcement - exceeds max")
    
    ops = BaseOperations(file_loader)
    
    # Request limit exceeding max
    requested_limit = MAX_ROW_LIMIT + 500
    enforced_limit = ops._enforce_row_limit(requested_limit)
    
    print(f"  Requested: {requested_limit}")
    print(f"  Enforced: {enforced_limit}")
    print(f"  MAX_ROW_LIMIT: {MAX_ROW_LIMIT}")
    
    assert enforced_limit == MAX_ROW_LIMIT, "Should cap at MAX_ROW_LIMIT"
    assert enforced_limit < requested_limit, "Should reduce excessive request"


def test_enforce_row_limit_exactly_max(file_loader):
    """Test _enforce_row_limit with limit exactly at MAX_ROW_LIMIT.
    
    Verifies:
    - Allows exactly MAX_ROW_LIMIT
    - Boundary condition handled correctly
    """
    print("\n📏 Testing row limit enforcement - exactly at max")
    
    ops = BaseOperations(file_loader)
    
    # Request exactly max
    requested_limit = MAX_ROW_LIMIT
    enforced_limit = ops._enforce_row_limit(requested_limit)
    
    print(f"  Requested: {requested_limit}")
    print(f"  Enforced: {enforced_limit}")
    
    assert enforced_limit == MAX_ROW_LIMIT, "Should allow exactly MAX_ROW_LIMIT"


def test_enforce_row_limit_small_value(file_loader):
    """Test _enforce_row_limit with small limit value.
    
    Verifies:
    - Small values pass through unchanged
    - No minimum limit enforced
    """
    print("\n📏 Testing row limit enforcement - small value")
    
    ops = BaseOperations(file_loader)
    
    # Request small limit
    requested_limit = 10
    enforced_limit = ops._enforce_row_limit(requested_limit)
    
    print(f"  Requested: {requested_limit}")
    print(f"  Enforced: {enforced_limit}")
    
    assert enforced_limit == requested_limit, "Should allow small limits"


# ============================================================================
# Test Numeric Column Conversion (Text-to-Number)
# ============================================================================

def test_ensure_numeric_column_already_numeric(file_loader):
    """Test _ensure_numeric_column with already numeric column.
    
    Verifies:
    - Returns column unchanged if already numeric
    - No conversion attempted
    """
    print("\n🔢 Testing numeric column conversion - already numeric")
    
    ops = BaseOperations(file_loader)
    
    # Create numeric series
    col_data = pd.Series([1, 2, 3, 4, 5], dtype='int64')
    
    result = ops._ensure_numeric_column(col_data, "TestColumn")
    
    print(f"  Original dtype: {col_data.dtype}")
    print(f"  Result dtype: {result.dtype}")
    
    assert pd.api.types.is_numeric_dtype(result), "Should remain numeric"
    assert result.equals(col_data), "Should be unchanged"


def test_ensure_numeric_column_convertible_strings(file_loader):
    """Test _ensure_numeric_column with numeric strings.
    
    Verifies:
    - Converts numeric strings to numbers
    - Returns float64 dtype
    """
    print("\n🔢 Testing numeric column conversion - convertible strings")
    
    ops = BaseOperations(file_loader)
    
    # Create string series with numeric values
    col_data = pd.Series(["1", "2", "3", "4", "5"], dtype='object')
    
    result = ops._ensure_numeric_column(col_data, "TestColumn")
    
    print(f"  Original dtype: {col_data.dtype}")
    print(f"  Result dtype: {result.dtype}")
    print(f"  Sample values: {result.head().tolist()}")
    
    assert pd.api.types.is_numeric_dtype(result), "Should convert to numeric"
    assert result.tolist() == [1.0, 2.0, 3.0, 4.0, 5.0], "Should have correct values"


def test_ensure_numeric_column_partial_conversion(file_loader):
    """Test _ensure_numeric_column with partially convertible data.
    
    Verifies:
    - Converts when >= 50% values are numeric
    - Non-numeric values become NaN
    """
    print("\n🔢 Testing numeric column conversion - partial conversion")
    
    ops = BaseOperations(file_loader)
    
    # Create series with 60% numeric strings (above 50% threshold)
    col_data = pd.Series(["1", "2", "3", "4", "5", "abc", "def", "ghi", "6", "7"], dtype='object')
    
    result = ops._ensure_numeric_column(col_data, "TestColumn")
    
    print(f"  Original dtype: {col_data.dtype}")
    print(f"  Result dtype: {result.dtype}")
    print(f"  Non-null count: {result.notna().sum()}/10")
    
    assert pd.api.types.is_numeric_dtype(result), "Should convert to numeric"
    assert result.notna().sum() == 7, "Should have 7 numeric values (70%)"


def test_ensure_numeric_column_non_convertible(file_loader):
    """Test _ensure_numeric_column with non-convertible text.
    
    Verifies:
    - Raises ValueError when < 50% values convertible
    - Error message is informative
    """
    print("\n🔢 Testing numeric column conversion - non-convertible")
    
    ops = BaseOperations(file_loader)
    
    # Create series with mostly text (< 50% numeric)
    col_data = pd.Series(["abc", "def", "ghi", "jkl", "1"], dtype='object')
    
    with pytest.raises(ValueError) as exc_info:
        ops._ensure_numeric_column(col_data, "TestColumn")
    
    error_msg = str(exc_info.value)
    print(f"  ✅ Caught expected error")
    print(f"     Error: {error_msg}")
    
    assert "TestColumn" in error_msg, "Should mention column name"
    assert "not numeric" in error_msg.lower(), "Should explain problem"
    assert "1/5" in error_msg or "1" in error_msg, "Should show conversion stats"


def test_ensure_numeric_column_datetime_type(file_loader):
    """Test _ensure_numeric_column with datetime column.
    
    Verifies:
    - Raises ValueError for datetime columns
    - Error message mentions current type
    """
    print("\n🔢 Testing numeric column conversion - datetime type")
    
    ops = BaseOperations(file_loader)
    
    # Create datetime series
    col_data = pd.Series(pd.date_range('2024-01-01', periods=5))
    
    with pytest.raises(ValueError) as exc_info:
        ops._ensure_numeric_column(col_data, "DateColumn")
    
    error_msg = str(exc_info.value)
    print(f"  ✅ Caught expected error")
    print(f"     Error: {error_msg}")
    
    assert "DateColumn" in error_msg, "Should mention column name"
    assert "must be numeric" in error_msg.lower(), "Should explain requirement"


def test_ensure_numeric_column_custom_conversion_rate(file_loader):
    """Test _ensure_numeric_column with custom min_conversion_rate.
    
    Verifies:
    - Respects custom conversion rate threshold
    - Can require higher conversion rate
    """
    print("\n🔢 Testing numeric column conversion - custom conversion rate")
    
    ops = BaseOperations(file_loader)
    
    # Create series with 60% numeric (would pass 50% threshold)
    col_data = pd.Series(["1", "2", "3", "4", "5", "6", "abc", "def", "ghi", "jkl"], dtype='object')
    
    # Require 70% conversion rate (should fail)
    with pytest.raises(ValueError) as exc_info:
        ops._ensure_numeric_column(col_data, "TestColumn", min_conversion_rate=0.7)
    
    error_msg = str(exc_info.value)
    print(f"  ✅ Custom threshold enforced")
    print(f"     Error: {error_msg}")
    
    assert "6/10" in error_msg or "6" in error_msg, "Should show actual conversion count"


def test_ensure_numeric_column_all_nulls(file_loader):
    """Test _ensure_numeric_column with all null values.
    
    Verifies:
    - Handles all-null column gracefully
    - Returns numeric series with all NaN (0/0 = 100% conversion success)
    """
    print("\n🔢 Testing numeric column conversion - all nulls")
    
    ops = BaseOperations(file_loader)
    
    # Create series with all nulls
    col_data = pd.Series([None, None, None, None, None], dtype='object')
    
    # Should convert successfully (0 non-null original, 0 non-null converted = 0/0 = 100% success)
    result = ops._ensure_numeric_column(col_data, "NullColumn")
    
    print(f"  ✅ All-null column converted")
    print(f"     Result dtype: {result.dtype}")
    print(f"     All NaN: {result.isna().all()}")
    
    assert pd.api.types.is_numeric_dtype(result), "Should be numeric dtype"
    assert result.isna().all(), "All values should be NaN"
