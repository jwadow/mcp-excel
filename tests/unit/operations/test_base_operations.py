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


# ============================================================================
# Test Unicode Normalization (Column Name Matching)
# ============================================================================

def test_normalize_column_name_nfc_nfd(file_loader):
    """Test _normalize_column_name with NFC vs NFD Unicode forms.
    
    Verifies:
    - NFC and NFD forms normalize to same result
    - Result is in NFC form
    - Handles composed vs decomposed characters
    """
    print("\n🔤 Testing Unicode normalization - NFC vs NFD")
    
    import unicodedata
    ops = BaseOperations(file_loader)
    
    cafe_nfc = "café"  # NFC (composed)
    cafe_nfd = unicodedata.normalize('NFD', "café")  # NFD (decomposed)
    
    normalized_nfc = ops._normalize_column_name(cafe_nfc)
    normalized_nfd = ops._normalize_column_name(cafe_nfd)
    
    print(f"  NFC input: {repr(cafe_nfc)} → {repr(normalized_nfc)}")
    print(f"  NFD input: {repr(cafe_nfd)} → {repr(normalized_nfd)}")
    
    assert normalized_nfc == normalized_nfd, "NFC and NFD should normalize to same result"
    assert normalized_nfc == "café", "Should normalize to NFC form"
    assert unicodedata.is_normalized('NFC', normalized_nfc), "Result should be in NFC form"


def test_normalize_column_name_nonbreaking_space(file_loader):
    """Test _normalize_column_name with non-breaking spaces.
    
    Verifies:
    - Non-breaking spaces (U+00A0) converted to regular spaces (U+0020)
    - Regular and non-breaking spaces normalize to same result
    """
    print("\n🔤 Testing Unicode normalization - non-breaking space")
    
    ops = BaseOperations(file_loader)
    
    with_regular = "Нетто, кг"  # Regular space (U+0020)
    with_nonbreaking = "Нетто,\u00A0кг"  # Non-breaking space (U+00A0)
    
    normalized_regular = ops._normalize_column_name(with_regular)
    normalized_nonbreaking = ops._normalize_column_name(with_nonbreaking)
    
    print(f"  Regular space: {repr(with_regular)} → {repr(normalized_regular)}")
    print(f"  Non-breaking: {repr(with_nonbreaking)} → {repr(normalized_nonbreaking)}")
    
    assert normalized_regular == normalized_nonbreaking, "Should normalize to same result"
    assert "\u00A0" not in normalized_nonbreaking, "Should remove non-breaking spaces"
    assert " " in normalized_nonbreaking, "Should have regular space"


def test_normalize_column_name_whitespace(file_loader):
    """Test _normalize_column_name with various whitespace issues.
    
    Verifies:
    - Leading/trailing whitespace removed
    - Multiple consecutive spaces collapsed to one
    - Tabs and newlines handled
    """
    print("\n🔤 Testing Unicode normalization - whitespace")
    
    ops = BaseOperations(file_loader)
    
    test_cases = [
        (" Name ", "Name"),
        ("  Price  ", "Price"),
        ("Name  Value", "Name Value"),
        ("\tValue\t", "Value"),
        ("A    B    C", "A B C"),
    ]
    
    for input_str, expected in test_cases:
        result = ops._normalize_column_name(input_str)
        print(f"  {repr(input_str)} → {repr(result)}")
        assert result == expected, f"Should normalize {repr(input_str)} to {repr(expected)}"


def test_normalize_column_name_combined_issues(file_loader):
    """Test _normalize_column_name with combined Unicode + whitespace.
    
    Verifies:
    - Handles NFC + non-breaking space + whitespace together
    - All normalizations applied correctly
    """
    print("\n🔤 Testing Unicode normalization - combined issues")
    
    import unicodedata
    ops = BaseOperations(file_loader)
    
    # NFC + non-breaking space + leading/trailing spaces
    cafe_nfc = "café"
    input_str = f" {cafe_nfc}\u00A0bar "
    
    result = ops._normalize_column_name(input_str)
    
    print(f"  Input: {repr(input_str)}")
    print(f"  Output: {repr(result)}")
    
    assert result == "café bar", "Should handle combined normalization"
    assert "\u00A0" not in result, "Should remove non-breaking space"
    assert not result.startswith(" "), "Should remove leading space"
    assert not result.endswith(" "), "Should remove trailing space"


def test_find_column_nfc_vs_nfd(file_loader):
    """Test _find_column with different Unicode forms.
    
    Verifies:
    - Finds column when request uses NFD but DataFrame has NFC
    - Returns original column name from DataFrame
    """
    print("\n🔍 Testing find_column - NFC vs NFD")
    
    import unicodedata
    ops = BaseOperations(file_loader)
    
    # DataFrame has NFC
    df = pd.DataFrame({"café": [1, 2, 3]})
    
    # Request uses NFD
    cafe_nfd = unicodedata.normalize('NFD', "café")
    
    actual_column = ops._find_column(df, cafe_nfd, context="test")
    
    print(f"  Request (NFD): {repr(cafe_nfd)}")
    print(f"  Found: {repr(actual_column)}")
    
    assert actual_column == "café", "Should find column despite Unicode form difference"
    assert actual_column in df.columns, "Should return original column name"


def test_find_column_nonbreaking_space(file_loader):
    """Test _find_column with non-breaking space mismatch.
    
    Verifies:
    - Finds column when request has regular space but DataFrame has non-breaking
    - Returns original column name with non-breaking space
    """
    print("\n🔍 Testing find_column - non-breaking space")
    
    ops = BaseOperations(file_loader)
    
    # DataFrame has non-breaking space
    df = pd.DataFrame({"Нетто,\u00A0кг": [1, 2, 3]})
    
    # Request uses regular space
    request_with_regular = "Нетто, кг"
    
    actual_column = ops._find_column(df, request_with_regular, context="test")
    
    print(f"  Request (regular space): {repr(request_with_regular)}")
    print(f"  Found: {repr(actual_column)}")
    
    assert actual_column == "Нетто,\u00A0кг", "Should find column despite space difference"
    assert "\u00A0" in actual_column, "Should return original with non-breaking space"


def test_find_column_whitespace_variations(file_loader):
    """Test _find_column with whitespace variations.
    
    Verifies:
    - Finds column when whitespace differs
    - Returns original column name with original whitespace
    """
    print("\n🔍 Testing find_column - whitespace variations")
    
    ops = BaseOperations(file_loader)
    
    # DataFrame has leading/trailing spaces
    df = pd.DataFrame({" Name ": [1, 2, 3]})
    
    # Request without spaces
    actual_column = ops._find_column(df, "Name", context="test")
    
    print(f"  Request: 'Name'")
    print(f"  Found: {repr(actual_column)}")
    
    assert actual_column == " Name ", "Should find column despite whitespace difference"


def test_find_column_not_found_with_suggestions(file_loader):
    """Test _find_column error message with fuzzy suggestions.
    
    Verifies:
    - ValueError raised when column not found
    - Error message includes fuzzy suggestions
    - Mentions context in error
    """
    print("\n🔍 Testing find_column - not found with suggestions")
    
    ops = BaseOperations(file_loader)
    
    df = pd.DataFrame({"café": [1, 2, 3], "Москва": [4, 5, 6]})
    
    with pytest.raises(ValueError) as exc_info:
        ops._find_column(df, "caffe", context="test_operation")
    
    error_msg = str(exc_info.value)
    print(f"  Error message: {error_msg}")
    
    assert "not found in test_operation" in error_msg, "Should mention context"
    assert "Did you mean" in error_msg, "Should provide suggestions"
    assert "café" in error_msg, "Should suggest similar column"


def test_find_column_not_found_no_suggestions(file_loader):
    """Test _find_column error when no close matches exist.
    
    Verifies:
    - Error message lists available columns
    - No suggestions when no close matches
    """
    print("\n🔍 Testing find_column - not found without suggestions")
    
    ops = BaseOperations(file_loader)
    
    df = pd.DataFrame({"café": [1, 2, 3]})
    
    with pytest.raises(ValueError) as exc_info:
        ops._find_column(df, "xyz123", context="test")
    
    error_msg = str(exc_info.value)
    print(f"  Error message: {error_msg}")
    
    assert "not found in test" in error_msg, "Should mention context"
    assert "Available columns" in error_msg, "Should list available columns"
    assert "Did you mean" not in error_msg, "Should not suggest when no close matches"


def test_find_columns_multiple_unicode(file_loader):
    """Test _find_columns with multiple Unicode columns.
    
    Verifies:
    - Finds all columns with Unicode variations
    - Returns original column names from DataFrame
    - Preserves order
    """
    print("\n🔍 Testing find_columns - multiple Unicode columns")
    
    import unicodedata
    ops = BaseOperations(file_loader)
    
    df = pd.DataFrame({
        "café": [1, 2, 3],
        "Москва": [4, 5, 6],
        "naïve": [7, 8, 9],
    })
    
    # Request with NFD forms
    cafe_nfd = unicodedata.normalize('NFD', "café")
    naive_nfd = unicodedata.normalize('NFD', "naïve")
    
    actual_columns = ops._find_columns(
        df,
        [cafe_nfd, "Москва", naive_nfd],
        context="test"
    )
    
    print(f"  Found: {actual_columns}")
    
    assert len(actual_columns) == 3, "Should find all columns"
    assert actual_columns == ["café", "Москва", "naïve"], "Should return original names"


def test_find_columns_mixed_issues(file_loader):
    """Test _find_columns with mixed Unicode and whitespace issues.
    
    Verifies:
    - Handles combination of NFC/NFD, spaces, whitespace
    - Returns all original column names correctly
    """
    print("\n🔍 Testing find_columns - mixed issues")
    
    import unicodedata
    ops = BaseOperations(file_loader)
    
    df = pd.DataFrame({
        "café": [1, 2, 3],
        " Name ": [4, 5, 6],
        "Нетто,\u00A0кг": [7, 8, 9],
    })
    
    # Request with different forms
    cafe_nfd = unicodedata.normalize('NFD', "café")
    
    actual_columns = ops._find_columns(
        df,
        [cafe_nfd, "Name", "Нетто, кг"],  # NFD, no spaces, regular space
        context="test"
    )
    
    print(f"  Found: {actual_columns}")
    
    assert len(actual_columns) == 3, "Should find all columns"
    assert actual_columns[0] == "café", "Should find NFC from NFD request"
    assert actual_columns[1] == " Name ", "Should find column with spaces"
    assert actual_columns[2] == "Нетто,\u00A0кг", "Should find column with non-breaking space"


def test_find_columns_one_not_found(file_loader):
    """Test _find_columns when one column not found.
    
    Verifies:
    - Raises ValueError on first missing column
    - Error message mentions missing column
    """
    print("\n🔍 Testing find_columns - one not found")
    
    ops = BaseOperations(file_loader)
    
    df = pd.DataFrame({"café": [1, 2, 3], "Москва": [4, 5, 6]})
    
    with pytest.raises(ValueError) as exc_info:
        ops._find_columns(
            df,
            ["café", "NotExist", "Москва"],
            context="test"
        )
    
    error_msg = str(exc_info.value)
    print(f"  Error message: {error_msg}")
    
    assert "NotExist" in error_msg, "Should mention missing column"
    assert "not found in test" in error_msg, "Should mention context"


def test_find_columns_empty_list(file_loader):
    """Test _find_columns with empty list.
    
    Verifies:
    - Returns empty list for empty input
    - No error raised
    """
    print("\n🔍 Testing find_columns - empty list")
    
    ops = BaseOperations(file_loader)
    
    df = pd.DataFrame({"café": [1, 2, 3]})
    
    actual_columns = ops._find_columns(df, [], context="test")
    
    print(f"  Found: {actual_columns}")
    
    assert actual_columns == [], "Should return empty list"


def test_normalize_column_name_cyrillic(file_loader):
    """Test _normalize_column_name with Cyrillic Unicode.
    
    Verifies:
    - Cyrillic NFC/NFD forms normalize to same result
    - Result is in NFC form
    """
    print("\n🔤 Testing Unicode normalization - Cyrillic")
    
    import unicodedata
    ops = BaseOperations(file_loader)
    
    moscow_nfc = "Москва"
    moscow_nfd = unicodedata.normalize('NFD', moscow_nfc)
    
    normalized_nfc = ops._normalize_column_name(moscow_nfc)
    normalized_nfd = ops._normalize_column_name(moscow_nfd)
    
    print(f"  NFC: {repr(moscow_nfc)} → {repr(normalized_nfc)}")
    print(f"  NFD: {repr(moscow_nfd)} → {repr(normalized_nfd)}")
    
    assert normalized_nfc == normalized_nfd, "Cyrillic NFC/NFD should normalize to same"
    assert unicodedata.is_normalized('NFC', normalized_nfc), "Should be in NFC form"


def test_normalize_column_name_empty_string(file_loader):
    """Test _normalize_column_name with empty string.
    
    Verifies:
    - Empty string returns empty string
    - No error raised
    """
    print("\n🔤 Testing Unicode normalization - empty string")
    
    ops = BaseOperations(file_loader)
    
    result = ops._normalize_column_name("")
    
    print(f"  Result: {repr(result)}")
    
    assert result == "", "Should return empty string"


def test_normalize_column_name_only_whitespace(file_loader):
    """Test _normalize_column_name with only whitespace.
    
    Verifies:
    - Whitespace-only strings become empty after normalization
    - Various whitespace types handled
    """
    print("\n🔤 Testing Unicode normalization - only whitespace")
    
    ops = BaseOperations(file_loader)
    
    test_cases = ["   ", "\t\t", "\n\n", " \t\n "]
    
    for input_str in test_cases:
        result = ops._normalize_column_name(input_str)
        print(f"  {repr(input_str)} → {repr(result)}")
        assert result == "", "Should return empty string after stripping"


def test_find_column_case_sensitive(file_loader):
    """Test that _find_column matching is case-sensitive.
    
    Verifies:
    - "Name" and "name" are treated as different columns
    - Exact case match required
    """
    print("\n🔍 Testing find_column - case sensitivity")
    
    ops = BaseOperations(file_loader)
    
    df = pd.DataFrame({
        "Name": [1, 2, 3],
        "name": [4, 5, 6],
    })
    
    actual_upper = ops._find_column(df, "Name", context="test")
    actual_lower = ops._find_column(df, "name", context="test")
    
    print(f"  'Name' → {repr(actual_upper)}")
    print(f"  'name' → {repr(actual_lower)}")
    
    assert actual_upper == "Name", "Should find uppercase version"
    assert actual_lower == "name", "Should find lowercase version"


# ============================================================================
# _ADD_SAMPLE_ROWS TESTS
# ============================================================================

def test_add_sample_rows_with_none(file_loader):
    """Test _add_sample_rows returns None when sample_size is None.
    
    Verifies:
    - Returns None when sample_size=None (default behavior)
    - No data processing occurs
    """
    print("\n🔍 Testing _add_sample_rows with None")
    
    ops = BaseOperations(file_loader)
    df = pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})
    
    result = ops._add_sample_rows(df, sample_size=None)
    
    print(f"  Result: {result}")
    
    assert result is None, "Should return None when sample_size is None"


def test_add_sample_rows_with_count(file_loader):
    """Test _add_sample_rows returns N rows when sample_size=N.
    
    Verifies:
    - Returns exactly N rows when N < len(df)
    - Rows are formatted as list of dicts
    - Values are formatted (integers without .0)
    """
    print("\n🔍 Testing _add_sample_rows with count")
    
    ops = BaseOperations(file_loader)
    df = pd.DataFrame({
        "Name": ["Alice", "Bob", "Charlie"],
        "Age": [25, 30, 35],
        "Score": [95.5, 88.0, 92.3]
    })
    
    result = ops._add_sample_rows(df, sample_size=2)
    
    print(f"  Result: {result}")
    
    assert result is not None, "Should return list"
    assert isinstance(result, list), "Should be list"
    assert len(result) == 2, "Should return exactly 2 rows"
    
    # Verify structure
    assert all(isinstance(row, dict) for row in result), "Each row should be dict"
    assert all("Name" in row and "Age" in row and "Score" in row for row in result), "All columns present"
    
    # Verify formatting: 88.0 should become 88 (int)
    assert result[1]["Age"] == 30, "Age should be int"
    assert isinstance(result[1]["Age"], int), "Age should be int type"
    assert result[1]["Score"] == 88, "Score 88.0 should be formatted as 88"
    assert isinstance(result[1]["Score"], int), "Score 88.0 should be int type"


def test_add_sample_rows_exceeds_size(file_loader):
    """Test _add_sample_rows when requested size exceeds DataFrame size.
    
    Verifies:
    - Returns all rows when sample_size > len(df)
    - No error raised
    - Returns min(sample_size, len(df)) rows
    """
    print("\n🔍 Testing _add_sample_rows exceeds size")
    
    ops = BaseOperations(file_loader)
    df = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
    
    result = ops._add_sample_rows(df, sample_size=10)
    
    print(f"  Requested: 10, Available: 2, Got: {len(result)}")
    
    assert result is not None, "Should return list"
    assert len(result) == 2, "Should return all 2 rows (not 10)"


def test_add_sample_rows_empty_dataframe(file_loader):
    """Test _add_sample_rows with empty DataFrame.
    
    Verifies:
    - Returns empty list for empty DataFrame
    - No error raised
    """
    print("\n🔍 Testing _add_sample_rows with empty DataFrame")
    
    ops = BaseOperations(file_loader)
    df = pd.DataFrame({"A": [], "B": []})
    
    result = ops._add_sample_rows(df, sample_size=5)
    
    print(f"  Result: {result}")
    
    assert result is not None, "Should return list"
    assert isinstance(result, list), "Should be list"
    assert len(result) == 0, "Should return empty list"


def test_add_sample_rows_formats_integers(file_loader):
    """Test _add_sample_rows formats float integers correctly.
    
    Verifies:
    - Float values like 50089416.0 are formatted as int 50089416
    - Uses _format_value() internally
    - No .0 suffix in output
    """
    print("\n🔍 Testing _add_sample_rows formats integers")
    
    ops = BaseOperations(file_loader)
    df = pd.DataFrame({
        "ID": [50089416.0, 50089417.0],
        "Value": [100.5, 200.0]
    })
    
    result = ops._add_sample_rows(df, sample_size=2)
    
    print(f"  Result: {result}")
    
    # Verify ID is formatted as int (no .0)
    assert result[0]["ID"] == 50089416, "ID should be int"
    assert isinstance(result[0]["ID"], int), "ID should be int type"
    
    # Verify Value with .0 is also formatted as int
    assert result[1]["Value"] == 200, "Value 200.0 should be int"
    assert isinstance(result[1]["Value"], int), "Value 200.0 should be int type"
    
    # Verify Value with decimal stays float
    assert result[0]["Value"] == 100.5, "Value 100.5 should stay float"
    assert isinstance(result[0]["Value"], float), "Value 100.5 should be float type"
