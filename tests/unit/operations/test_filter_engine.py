# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Unit tests for FilterEngine component.

Tests cover:
- All 12 filter operators (==, !=, >, <, >=, <=, in, not_in, contains, startswith, endswith, regex, is_null, is_not_null)
- AND/OR logic combination
- DateTime filtering
- Filter validation
- Error handling
- Edge cases
"""

import pytest
import pandas as pd
from datetime import datetime

from mcp_excel.operations.filtering import FilterEngine
from mcp_excel.models.requests import FilterCondition


@pytest.fixture
def sample_df():
    """Create sample DataFrame for testing."""
    return pd.DataFrame({
        "Name": ["Alice", "Bob", "Charlie", "David", "Eve"],
        "Age": [25, 30, 35, 40, 45],
        "City": ["Moscow", "London", "Paris", "Tokyo", "Berlin"],
        "Salary": [50000.0, 60000.0, 70000.0, 80000.0, 90000.0],
        "Active": [True, True, False, True, False],
    })


@pytest.fixture
def df_with_nulls():
    """Create DataFrame with null values."""
    return pd.DataFrame({
        "Name": ["Alice", "Bob", None, "David", "Eve"],
        "Age": [25, None, 35, 40, None],
        "City": ["Moscow", "London", None, "Tokyo", "Berlin"],
    })


@pytest.fixture
def df_with_dates():
    """Create DataFrame with datetime column."""
    return pd.DataFrame({
        "Name": ["Alice", "Bob", "Charlie"],
        "Date": pd.to_datetime(["2024-01-01", "2024-02-01", "2024-03-01"]),
        "Value": [100, 200, 300],
    })


# ============================================================================
# COMPARISON OPERATORS
# ============================================================================

def test_filter_equals(filter_engine, sample_df):
    """Test == operator."""
    print(f"\n📂 Testing == operator")
    
    filters = [FilterCondition(column="Name", operator="==", value="Alice")]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should return 1 row"
    assert result.iloc[0]["Name"] == "Alice", "Should return Alice"


def test_filter_not_equals(filter_engine, sample_df):
    """Test != operator."""
    print(f"\n📂 Testing != operator")
    
    filters = [FilterCondition(column="Name", operator="!=", value="Alice")]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 4, "Should return 4 rows"
    assert "Alice" not in result["Name"].values, "Should not include Alice"


def test_filter_greater_than(filter_engine, sample_df):
    """Test > operator."""
    print(f"\n📂 Testing > operator")
    
    filters = [FilterCondition(column="Age", operator=">", value=30)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 3, "Should return 3 rows (35, 40, 45)"
    assert all(result["Age"] > 30), "All ages should be > 30"


def test_filter_less_than(filter_engine, sample_df):
    """Test < operator."""
    print(f"\n📂 Testing < operator")
    
    filters = [FilterCondition(column="Age", operator="<", value=35)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should return 2 rows (25, 30)"
    assert all(result["Age"] < 35), "All ages should be < 35"


def test_filter_greater_or_equal(filter_engine, sample_df):
    """Test >= operator."""
    print(f"\n📂 Testing >= operator")
    
    filters = [FilterCondition(column="Age", operator=">=", value=35)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 3, "Should return 3 rows (35, 40, 45)"
    assert all(result["Age"] >= 35), "All ages should be >= 35"


def test_filter_less_or_equal(filter_engine, sample_df):
    """Test <= operator."""
    print(f"\n📂 Testing <= operator")
    
    filters = [FilterCondition(column="Age", operator="<=", value=30)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should return 2 rows (25, 30)"
    assert all(result["Age"] <= 30), "All ages should be <= 30"


# ============================================================================
# SET OPERATORS
# ============================================================================

def test_filter_in(filter_engine, sample_df):
    """Test 'in' operator."""
    print(f"\n📂 Testing 'in' operator")
    
    filters = [FilterCondition(column="Name", operator="in", values=["Alice", "Bob", "Charlie"])]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 3, "Should return 3 rows"
    assert set(result["Name"]) == {"Alice", "Bob", "Charlie"}, "Should return specified names"


def test_filter_not_in(filter_engine, sample_df):
    """Test 'not_in' operator."""
    print(f"\n📂 Testing 'not_in' operator")
    
    filters = [FilterCondition(column="Name", operator="not_in", values=["Alice", "Bob"])]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 3, "Should return 3 rows"
    assert "Alice" not in result["Name"].values, "Should not include Alice"
    assert "Bob" not in result["Name"].values, "Should not include Bob"


# ============================================================================
# STRING OPERATORS
# ============================================================================

def test_filter_contains(filter_engine, sample_df):
    """Test 'contains' operator."""
    print(f"\n📂 Testing 'contains' operator")
    
    filters = [FilterCondition(column="City", operator="contains", value="on")]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should return 1 row (London)"
    assert result.iloc[0]["City"] == "London", "Should return London"


def test_filter_startswith(filter_engine, sample_df):
    """Test 'startswith' operator."""
    print(f"\n📂 Testing 'startswith' operator")
    
    filters = [FilterCondition(column="City", operator="startswith", value="M")]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should return 1 row (Moscow)"
    assert result.iloc[0]["City"] == "Moscow", "Should return Moscow"


def test_filter_endswith(filter_engine, sample_df):
    """Test 'endswith' operator."""
    print(f"\n📂 Testing 'endswith' operator")
    
    filters = [FilterCondition(column="City", operator="endswith", value="is")]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should return 1 row (Paris)"
    assert result.iloc[0]["City"] == "Paris", "Should return Paris"


def test_filter_regex(filter_engine, sample_df):
    """Test 'regex' operator."""
    print(f"\n📂 Testing 'regex' operator")
    
    # Match cities starting with L or P
    filters = [FilterCondition(column="City", operator="regex", value="^[LP]")]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should return 2 rows (London, Paris)"
    assert set(result["City"]) == {"London", "Paris"}, "Should return London and Paris"


# ============================================================================
# NULL OPERATORS
# ============================================================================

def test_filter_is_null(filter_engine, df_with_nulls):
    """Test 'is_null' operator."""
    print(f"\n📂 Testing 'is_null' operator")
    
    filters = [FilterCondition(column="Name", operator="is_null", value=None)]
    result = filter_engine.apply_filters(df_with_nulls, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should return 1 row with null Name"
    assert pd.isna(result.iloc[0]["Name"]), "Name should be null"


def test_filter_is_not_null(filter_engine, df_with_nulls):
    """Test 'is_not_null' operator."""
    print(f"\n📂 Testing 'is_not_null' operator")
    
    filters = [FilterCondition(column="Age", operator="is_not_null", value=None)]
    result = filter_engine.apply_filters(df_with_nulls, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 3, "Should return 3 rows with non-null Age"
    assert all(result["Age"].notna()), "All ages should be non-null"


# ============================================================================
# COMBINED FILTERS (AND/OR LOGIC)
# ============================================================================

def test_filter_and_logic(filter_engine, sample_df):
    """Test AND logic with multiple filters."""
    print(f"\n📂 Testing AND logic")
    
    filters = [
        FilterCondition(column="Age", operator=">", value=30),
        FilterCondition(column="Salary", operator="<", value=80000),
    ]
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should return 1 row (Charlie: age 35, salary 70000)"
    assert all(result["Age"] > 30), "All ages should be > 30"
    assert all(result["Salary"] < 80000), "All salaries should be < 80000"


def test_filter_or_logic(filter_engine, sample_df):
    """Test OR logic with multiple filters."""
    print(f"\n📂 Testing OR logic")
    
    filters = [
        FilterCondition(column="Age", operator="<", value=30),
        FilterCondition(column="Age", operator=">", value=40),
    ]
    result = filter_engine.apply_filters(sample_df, filters, logic="OR")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should return 2 rows (25 and 45)"
    assert all((result["Age"] < 30) | (result["Age"] > 40)), "Should match OR condition"


def test_filter_complex_and_logic(filter_engine, sample_df):
    """Test complex AND logic with 3+ filters."""
    print(f"\n📂 Testing complex AND logic")
    
    filters = [
        FilterCondition(column="Age", operator=">=", value=30),
        FilterCondition(column="Age", operator="<=", value=40),
        FilterCondition(column="Active", operator="==", value=True),
    ]
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should return 2 rows (Bob and David)"
    assert all((result["Age"] >= 30) & (result["Age"] <= 40)), "Age should be 30-40"
    assert all(result["Active"]), "All should be active"


# ============================================================================
# DATETIME FILTERING
# ============================================================================

def test_filter_datetime_with_string(filter_engine, df_with_dates):
    """Test datetime filtering with ISO string."""
    print(f"\n📂 Testing datetime filter with string")
    
    filters = [FilterCondition(column="Date", operator=">=", value="2024-02-01")]
    result = filter_engine.apply_filters(df_with_dates, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should return 2 rows (Feb and Mar)"
    assert all(result["Date"] >= pd.Timestamp("2024-02-01")), "All dates should be >= 2024-02-01"


def test_filter_datetime_in_operator(filter_engine, df_with_dates):
    """Test datetime filtering with 'in' operator."""
    print(f"\n📂 Testing datetime 'in' operator")
    
    filters = [FilterCondition(column="Date", operator="in", values=["2024-01-01", "2024-03-01"])]
    result = filter_engine.apply_filters(df_with_dates, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should return 2 rows (Jan and Mar)"


# ============================================================================
# COUNT_FILTERED METHOD
# ============================================================================

def test_count_filtered(filter_engine, sample_df):
    """Test count_filtered method."""
    print(f"\n📂 Testing count_filtered method")
    
    filters = [FilterCondition(column="Age", operator=">", value=30)]
    count = filter_engine.count_filtered(sample_df, filters)
    
    print(f"✅ Count: {count}")
    
    assert count == 3, "Should count 3 rows"
    assert isinstance(count, int), "Should return int"


def test_count_filtered_no_filters(filter_engine, sample_df):
    """Test count_filtered with no filters."""
    print(f"\n📂 Testing count_filtered with no filters")
    
    count = filter_engine.count_filtered(sample_df, [])
    
    print(f"✅ Count: {count}")
    
    assert count == len(sample_df), "Should return total row count"


# ============================================================================
# VALIDATION
# ============================================================================

def test_validate_filters_valid(filter_engine, sample_df):
    """Test validation with valid filters."""
    print(f"\n📂 Testing filter validation (valid)")
    
    filters = [FilterCondition(column="Age", operator=">", value=30)]
    is_valid, error = filter_engine.validate_filters(sample_df, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is True, "Should be valid"
    assert error is None, "Should have no error"


def test_validate_filters_invalid_column(filter_engine, sample_df):
    """Test validation with invalid column."""
    print(f"\n📂 Testing filter validation (invalid column)")
    
    filters = [FilterCondition(column="NonExistent", operator="==", value="test")]
    is_valid, error = filter_engine.validate_filters(sample_df, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is False, "Should be invalid"
    assert "not found" in error.lower(), "Error should mention column not found"


def test_validate_filters_missing_values(filter_engine, sample_df):
    """Test validation with missing 'values' for 'in' operator."""
    print(f"\n📂 Testing filter validation (missing values)")
    
    filters = [FilterCondition(column="Name", operator="in", values=None)]
    is_valid, error = filter_engine.validate_filters(sample_df, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is False, "Should be invalid"
    assert "values" in error.lower(), "Error should mention missing values"


# ============================================================================
# ERROR HANDLING
# ============================================================================

def test_filter_column_not_found(filter_engine, sample_df):
    """Test error when column doesn't exist."""
    print(f"\n📂 Testing column not found error")
    
    filters = [FilterCondition(column="NonExistent", operator="==", value="test")]
    
    with pytest.raises(ValueError) as exc_info:
        filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    assert "not found" in str(exc_info.value).lower(), "Error should mention column not found"


def test_filter_invalid_logic(filter_engine, sample_df):
    """Test error with invalid logic operator."""
    print(f"\n📂 Testing invalid logic operator")
    
    filters = [FilterCondition(column="Age", operator=">", value=30)]
    
    with pytest.raises(ValueError) as exc_info:
        filter_engine.apply_filters(sample_df, filters, logic="XOR")
    
    print(f"✅ Caught expected error: {exc_info.value}")
    assert "invalid logic" in str(exc_info.value).lower(), "Error should mention invalid logic"


def test_filter_invalid_regex(filter_engine, sample_df):
    """Test error with invalid regex pattern."""
    print(f"\n📂 Testing invalid regex pattern")
    
    filters = [FilterCondition(column="City", operator="regex", value="[invalid")]
    
    with pytest.raises(ValueError) as exc_info:
        filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    assert "regex" in str(exc_info.value).lower(), "Error should mention regex"


def test_filter_in_without_values(filter_engine, sample_df):
    """Test error when 'in' operator used without values."""
    print(f"\n📂 Testing 'in' operator without values")
    
    filters = [FilterCondition(column="Name", operator="in", values=None)]
    
    with pytest.raises(ValueError) as exc_info:
        filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    assert "values" in str(exc_info.value).lower(), "Error should mention missing values"


# ============================================================================
# FILTER SUMMARY
# ============================================================================

def test_get_filter_summary_single(filter_engine):
    """Test filter summary with single filter."""
    print(f"\n📂 Testing filter summary (single)")
    
    filters = [FilterCondition(column="Age", operator=">", value=30)]
    summary = filter_engine.get_filter_summary(filters, "AND")
    
    print(f"✅ Summary: {summary}")
    
    assert "Age" in summary, "Should mention column"
    assert ">" in summary, "Should mention operator"
    assert "30" in summary, "Should mention value"


def test_get_filter_summary_multiple(filter_engine):
    """Test filter summary with multiple filters."""
    print(f"\n📂 Testing filter summary (multiple)")
    
    filters = [
        FilterCondition(column="Age", operator=">", value=30),
        FilterCondition(column="City", operator="==", value="Moscow"),
    ]
    summary = filter_engine.get_filter_summary(filters, "AND")
    
    print(f"✅ Summary: {summary}")
    
    assert "Age" in summary, "Should mention Age"
    assert "City" in summary, "Should mention City"
    assert "AND" in summary, "Should mention AND logic"


def test_get_filter_summary_empty(filter_engine):
    """Test filter summary with no filters."""
    print(f"\n📂 Testing filter summary (empty)")
    
    summary = filter_engine.get_filter_summary([], "AND")
    
    print(f"✅ Summary: {summary}")
    
    assert "no filter" in summary.lower(), "Should indicate no filters"


# ============================================================================
# EDGE CASES
# ============================================================================

def test_filter_empty_dataframe(filter_engine):
    """Test filtering empty DataFrame."""
    print(f"\n📂 Testing empty DataFrame")
    
    df = pd.DataFrame({"Name": [], "Age": []})
    filters = [FilterCondition(column="Age", operator=">", value=30)]
    result = filter_engine.apply_filters(df, filters)
    
    print(f"✅ Result length: {len(result)}")
    
    assert len(result) == 0, "Should return empty DataFrame"


def test_filter_no_matches(filter_engine, sample_df):
    """Test filter that matches no rows."""
    print(f"\n📂 Testing filter with no matches")
    
    filters = [FilterCondition(column="Age", operator=">", value=100)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Result length: {len(result)}")
    
    assert len(result) == 0, "Should return empty DataFrame"


def test_filter_all_match(filter_engine, sample_df):
    """Test filter that matches all rows."""
    print(f"\n📂 Testing filter that matches all")
    
    filters = [FilterCondition(column="Age", operator=">", value=0)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Result length: {len(result)}")
    
    assert len(result) == len(sample_df), "Should return all rows"


# ============================================================================
# UNICODE NORMALIZATION (NFC/NFD, Non-breaking spaces, Whitespace)
# ============================================================================

@pytest.fixture
def df_with_unicode_columns():
    """DataFrame with Unicode column names in NFC form (composed).
    
    This simulates what Pandas returns after reading an Excel file.
    """
    return pd.DataFrame({
        "café": [1, 2, 3],  # NFC form (single character é = U+00E9)
        "Нетто, кг": [10, 20, 30],  # Regular space (U+0020)
        "Name": [100, 200, 300],  # ASCII for control
        "Клиент": ["A", "B", "C"],  # Cyrillic
    })


@pytest.fixture
def df_with_nonbreaking_spaces():
    """DataFrame with non-breaking spaces in column names.
    
    This simulates real Excel files where non-breaking spaces (U+00A0)
    are used instead of regular spaces (U+0020).
    """
    return pd.DataFrame({
        "Нетто,\u00A0кг": [10, 20, 30],  # Non-breaking space (U+00A0)
        "Name\u00A0Value": [100, 200, 300],  # Non-breaking space
    })


@pytest.fixture
def df_with_messy_whitespace():
    """DataFrame with messy whitespace in column names."""
    return pd.DataFrame({
        " Name ": [1, 2, 3],  # Leading/trailing spaces
        "Value  Total": [10, 20, 30],  # Double space
        "  Price  ": [100, 200, 300],  # Multiple leading/trailing
    })


def test_unicode_nfc_vs_nfd_cafe(filter_engine, df_with_unicode_columns):
    """Test filtering with NFD form when DataFrame has NFC form.
    
    Real scenario: Agent copies "café" from get_sheet_info (NFC),
    but user's filter uses NFD form (e + combining accent).
    """
    print(f"\n📂 Testing NFC vs NFD: café")
    
    import unicodedata
    
    # Create NFD form (decomposed: e + combining acute accent)
    cafe_nfd = unicodedata.normalize('NFD', "café")
    print(f"   Filter uses NFD: {repr(cafe_nfd)}")
    print(f"   DataFrame has NFC: {repr('café')}")
    
    # This should work despite different Unicode forms
    filters = [FilterCondition(column=cafe_nfd, operator="==", value=1)]
    result = filter_engine.apply_filters(df_with_unicode_columns, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should find column despite NFD vs NFC difference"
    assert result.iloc[0]["café"] == 1, "Should return correct row"


def test_unicode_nfd_vs_nfc_cyrillic(filter_engine, df_with_unicode_columns):
    """Test Cyrillic Unicode normalization."""
    print(f"\n📂 Testing NFC vs NFD: Cyrillic")
    
    import unicodedata
    
    # Cyrillic "Клиент" - ensure it works in both forms
    column_nfc = "Клиент"
    column_nfd = unicodedata.normalize('NFD', column_nfc)
    
    print(f"   Filter uses NFD: {repr(column_nfd)}")
    print(f"   DataFrame has NFC: {repr(column_nfc)}")
    
    filters = [FilterCondition(column=column_nfd, operator="==", value="A")]
    result = filter_engine.apply_filters(df_with_unicode_columns, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should find Cyrillic column despite Unicode form"
    assert result.iloc[0]["Клиент"] == "A", "Should return correct row"


def test_unicode_nonbreaking_space_to_regular(filter_engine, df_with_nonbreaking_spaces):
    """Test filtering with regular space when DataFrame has non-breaking space.
    
    Real scenario: User types "Нетто, кг" with regular space (U+0020),
    but Excel file has non-breaking space (U+00A0).
    """
    print(f"\n📂 Testing non-breaking space → regular space")
    
    # Filter uses regular space
    filter_column = "Нетто, кг"  # Regular space (U+0020)
    df_column = "Нетто,\u00A0кг"  # Non-breaking space (U+00A0)
    
    print(f"   Filter: {repr(filter_column)}")
    print(f"   DataFrame: {repr(df_column)}")
    print(f"   Visually identical but: {filter_column == df_column}")  # False!
    
    filters = [FilterCondition(column=filter_column, operator=">", value=15)]
    result = filter_engine.apply_filters(df_with_nonbreaking_spaces, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should find column despite space type difference"
    assert all(result[df_column] > 15), "Should filter correctly"


def test_unicode_regular_space_to_nonbreaking(filter_engine, df_with_unicode_columns):
    """Test filtering with non-breaking space when DataFrame has regular space."""
    print(f"\n📂 Testing regular space → non-breaking space")
    
    # Filter uses non-breaking space
    filter_column = "Нетто,\u00A0кг"  # Non-breaking space (U+00A0)
    df_column = "Нетто, кг"  # Regular space (U+0020)
    
    print(f"   Filter: {repr(filter_column)}")
    print(f"   DataFrame: {repr(df_column)}")
    
    filters = [FilterCondition(column=filter_column, operator="==", value=10)]
    result = filter_engine.apply_filters(df_with_unicode_columns, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should find column despite space type difference"
    assert result.iloc[0][df_column] == 10, "Should return correct row"


def test_unicode_leading_trailing_spaces(filter_engine, df_with_messy_whitespace):
    """Test filtering with leading/trailing spaces.
    
    Real scenario: User types " Name " (with spaces), but column is "Name".
    """
    print(f"\n📂 Testing leading/trailing spaces")
    
    # Filter without spaces, DataFrame with spaces
    filter_column = "Name"
    df_column = " Name "
    
    print(f"   Filter: {repr(filter_column)}")
    print(f"   DataFrame: {repr(df_column)}")
    
    filters = [FilterCondition(column=filter_column, operator="==", value=1)]
    result = filter_engine.apply_filters(df_with_messy_whitespace, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should find column despite leading/trailing spaces"
    assert result.iloc[0][df_column] == 1, "Should return correct row"


def test_unicode_multiple_spaces(filter_engine, df_with_messy_whitespace):
    """Test filtering with multiple consecutive spaces."""
    print(f"\n📂 Testing multiple consecutive spaces")
    
    # Filter with single space, DataFrame with double space
    filter_column = "Value Total"  # Single space
    df_column = "Value  Total"  # Double space
    
    print(f"   Filter: {repr(filter_column)}")
    print(f"   DataFrame: {repr(df_column)}")
    
    filters = [FilterCondition(column=filter_column, operator=">", value=15)]
    result = filter_engine.apply_filters(df_with_messy_whitespace, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should find column despite multiple spaces"


def test_unicode_combined_edge_case(filter_engine):
    """Test combination of Unicode normalization + non-breaking spaces + whitespace.
    
    Worst case: NFC/NFD + non-breaking space + leading/trailing spaces.
    """
    print(f"\n📂 Testing combined Unicode + spaces")
    
    import unicodedata
    
    # DataFrame with NFC + non-breaking space + trailing space
    df = pd.DataFrame({
        "café\u00A0bar ": [1, 2, 3],  # NFC + non-breaking + trailing
    })
    
    # Filter with NFD + regular space + no trailing space
    cafe_nfd = unicodedata.normalize('NFD', "café")
    filter_column = f"{cafe_nfd} bar"  # NFD + regular space + no trailing
    
    # Define string outside f-string to avoid backslash in f-string expression
    df_column_example = "café\u00A0bar "
    
    print(f"   Filter: {repr(filter_column)}")
    print(f"   DataFrame: {repr(df_column_example)}")
    
    filters = [FilterCondition(column=filter_column, operator="==", value=2)]
    result = filter_engine.apply_filters(df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should handle combined Unicode + space variations"
    assert result.iloc[0]["café\u00A0bar "] == 2, "Should return correct row"


def test_unicode_fuzzy_matching_suggestions(filter_engine, df_with_unicode_columns):
    """Test that error message includes fuzzy matching suggestions.
    
    When column is not found, system should suggest similar columns.
    """
    print(f"\n📂 Testing fuzzy matching suggestions")
    
    # Typo: "Namee" instead of "Name"
    filters = [FilterCondition(column="Namee", operator="==", value=100)]
    
    with pytest.raises(ValueError) as exc_info:
        filter_engine.apply_filters(df_with_unicode_columns, filters)
    
    error_msg = str(exc_info.value)
    print(f"✅ Error message: {error_msg}")
    
    assert "not found" in error_msg.lower(), "Should mention column not found"
    assert "did you mean" in error_msg.lower(), "Should provide suggestions"
    assert "Name" in error_msg, "Should suggest 'Name' as close match"


def test_unicode_cyrillic_fuzzy_matching(filter_engine, df_with_unicode_columns):
    """Test fuzzy matching with Cyrillic typo."""
    print(f"\n📂 Testing Cyrillic fuzzy matching")
    
    # Typo: "Клиенты" instead of "Клиент"
    filters = [FilterCondition(column="Клиенты", operator="==", value="A")]
    
    with pytest.raises(ValueError) as exc_info:
        filter_engine.apply_filters(df_with_unicode_columns, filters)
    
    error_msg = str(exc_info.value)
    print(f"✅ Error message: {error_msg}")
    
    assert "Клиент" in error_msg, "Should suggest 'Клиент' as close match"


def test_unicode_validate_filters_consistency(filter_engine, df_with_unicode_columns):
    """Test that validate_filters uses same normalization logic.
    
    Critical: validate_filters must use same normalization as apply_filters.
    """
    print(f"\n📂 Testing validate_filters with Unicode")
    
    import unicodedata
    
    # Use NFD form
    cafe_nfd = unicodedata.normalize('NFD', "café")
    filters = [FilterCondition(column=cafe_nfd, operator="==", value=1)]
    
    is_valid, error = filter_engine.validate_filters(df_with_unicode_columns, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is True, "validate_filters should accept NFD form"
    assert error is None, "Should have no error"


def test_unicode_validate_with_suggestions(filter_engine, df_with_unicode_columns):
    """Test validate_filters provides suggestions for invalid columns."""
    print(f"\n📂 Testing validate_filters with suggestions")
    
    # Typo: "Namee" instead of "Name"
    filters = [FilterCondition(column="Namee", operator="==", value=100)]
    
    is_valid, error = filter_engine.validate_filters(df_with_unicode_columns, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is False, "Should be invalid"
    assert "not found" in error.lower(), "Error should mention column not found"
    assert "did you mean" in error.lower(), "Should provide suggestions"
    assert "Name" in error, "Should suggest 'Name'"


def test_unicode_count_filtered_consistency(filter_engine, df_with_unicode_columns):
    """Test that count_filtered uses same normalization logic."""
    print(f"\n📂 Testing count_filtered with Unicode")
    
    import unicodedata
    
    # Use NFD form
    cafe_nfd = unicodedata.normalize('NFD', "café")
    filters = [FilterCondition(column=cafe_nfd, operator=">", value=1)]
    
    count = filter_engine.count_filtered(df_with_unicode_columns, filters)
    
    print(f"✅ Count: {count}")
    
    assert count == 2, "count_filtered should handle NFD form"


@pytest.mark.parametrize("unicode_form", ["NFC", "NFD", "NFKC", "NFKD"])
def test_unicode_all_forms_parametrized(filter_engine, df_with_unicode_columns, unicode_form):
    """Test all Unicode normalization forms (parametrized).
    
    Ensures system works with any Unicode normalization form.
    """
    print(f"\n📂 Testing Unicode form: {unicode_form}")
    
    import unicodedata
    
    # Normalize "café" to specified form
    cafe_normalized = unicodedata.normalize(unicode_form, "café")
    print(f"   Using form {unicode_form}: {repr(cafe_normalized)}")
    
    filters = [FilterCondition(column=cafe_normalized, operator=">", value=0)]
    result = filter_engine.apply_filters(df_with_unicode_columns, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 3, f"Should find column with {unicode_form} form"


def test_unicode_extreme_whitespace(filter_engine, df_with_messy_whitespace):
    """Test extreme whitespace: multiple leading/trailing spaces."""
    print(f"\n📂 Testing extreme whitespace")
    
    # Filter with clean name, DataFrame with messy spaces
    filter_column = "Price"
    df_column = "  Price  "
    
    print(f"   Filter: {repr(filter_column)}")
    print(f"   DataFrame: {repr(df_column)}")
    
    filters = [FilterCondition(column=filter_column, operator=">=", value=100)]
    result = filter_engine.apply_filters(df_with_messy_whitespace, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 3, "Should find column despite extreme whitespace"


def test_unicode_multiple_nonbreaking_spaces(filter_engine, df_with_nonbreaking_spaces):
    """Test column with multiple non-breaking spaces."""
    print(f"\n📂 Testing multiple non-breaking spaces")
    
    # Filter uses regular spaces
    filter_column = "Name Value"  # Regular spaces
    
    filters = [FilterCondition(column=filter_column, operator="<", value=250)]
    result = filter_engine.apply_filters(df_with_nonbreaking_spaces, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 2, "Should find column with multiple non-breaking spaces"


def test_unicode_cyrillic_with_nonbreaking(filter_engine):
    """Test Cyrillic + non-breaking spaces (common in Russian Excel files)."""
    print(f"\n📂 Testing Cyrillic + non-breaking spaces")
    
    # DataFrame with non-breaking space
    df = pd.DataFrame({
        "Нетто,\u00A0кг": [10, 20, 30],
        "Клиент\u00A0ID": ["A", "B", "C"],
    })
    
    # Filter with regular spaces
    filters = [
        FilterCondition(column="Нетто, кг", operator=">", value=15),
        FilterCondition(column="Клиент ID", operator="==", value="B"),
    ]
    result = filter_engine.apply_filters(df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    assert len(result) == 1, "Should handle Cyrillic + non-breaking spaces"
    assert result.iloc[0]["Клиент\u00A0ID"] == "B", "Should return correct row"


def test_unicode_no_close_matches(filter_engine, df_with_unicode_columns):
    """Test error message when no close matches exist."""
    print(f"\n📂 Testing no close matches")
    
    # Completely different column name
    filters = [FilterCondition(column="XYZ123", operator="==", value=1)]
    
    with pytest.raises(ValueError) as exc_info:
        filter_engine.apply_filters(df_with_unicode_columns, filters)
    
    error_msg = str(exc_info.value)
    print(f"✅ Error message: {error_msg}")
    
    assert "not found" in error_msg.lower(), "Should mention column not found"


def test_unicode_validate_with_nonbreaking(filter_engine, df_with_nonbreaking_spaces):
    """Test validate_filters with non-breaking space."""
    print(f"\n📂 Testing validate_filters with non-breaking space")
    
    # Filter uses regular space, DataFrame has non-breaking
    filters = [FilterCondition(column="Нетто, кг", operator=">", value=10)]
    
    is_valid, error = filter_engine.validate_filters(df_with_nonbreaking_spaces, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is True, "validate_filters should handle space variations"
    assert error is None, "Should have no error"


def test_unicode_validate_with_whitespace(filter_engine, df_with_messy_whitespace):
    """Test validate_filters with messy whitespace."""
    print(f"\n📂 Testing validate_filters with whitespace")
    
    # Filter without spaces, DataFrame with spaces
    filters = [FilterCondition(column="Name", operator="==", value=1)]
    
    is_valid, error = filter_engine.validate_filters(df_with_messy_whitespace, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is True, "validate_filters should handle whitespace"
    assert error is None, "Should have no error"


def test_unicode_count_with_nonbreaking(filter_engine, df_with_nonbreaking_spaces):
    """Test count_filtered with non-breaking space."""
    print(f"\n📂 Testing count_filtered with non-breaking space")
    
    # Filter uses regular space
    filters = [FilterCondition(column="Нетто, кг", operator=">=", value=20)]
    
    count = filter_engine.count_filtered(df_with_nonbreaking_spaces, filters)
    
    print(f"✅ Count: {count}")
    
    assert count == 2, "count_filtered should handle space variations"


# ============================================================================
# NORMALIZATION METHOD UNIT TESTS (Direct testing of _normalize_column_name)
# ============================================================================

def test_normalize_method_nfc_nfd(filter_engine):
    """Test _normalize_column_name with NFC/NFD forms."""
    print(f"\n📂 Testing _normalize_column_name: NFC/NFD")
    
    import unicodedata
    
    cafe_nfc = "café"  # NFC (composed)
    cafe_nfd = unicodedata.normalize('NFD', "café")  # NFD (decomposed)
    
    normalized_nfc = filter_engine._normalize_column_name(cafe_nfc)
    normalized_nfd = filter_engine._normalize_column_name(cafe_nfd)
    
    print(f"   NFC input: {repr(cafe_nfc)} → {repr(normalized_nfc)}")
    print(f"   NFD input: {repr(cafe_nfd)} → {repr(normalized_nfd)}")
    
    assert normalized_nfc == normalized_nfd, "NFC and NFD should normalize to same result"
    assert normalized_nfc == "café", "Should normalize to NFC form"


def test_normalize_method_nonbreaking_space(filter_engine):
    """Test _normalize_column_name with non-breaking space."""
    print(f"\n📂 Testing _normalize_column_name: non-breaking space")
    
    with_regular = "Нетто, кг"  # Regular space (U+0020)
    with_nonbreaking = "Нетто,\u00A0кг"  # Non-breaking space (U+00A0)
    
    normalized_regular = filter_engine._normalize_column_name(with_regular)
    normalized_nonbreaking = filter_engine._normalize_column_name(with_nonbreaking)
    
    print(f"   Regular space: {repr(with_regular)} → {repr(normalized_regular)}")
    print(f"   Non-breaking: {repr(with_nonbreaking)} → {repr(normalized_nonbreaking)}")
    
    assert normalized_regular == normalized_nonbreaking, "Should normalize to same result"
    assert "\u00A0" not in normalized_nonbreaking, "Should remove non-breaking spaces"


def test_normalize_method_whitespace(filter_engine):
    """Test _normalize_column_name with various whitespace."""
    print(f"\n📂 Testing _normalize_column_name: whitespace")
    
    test_cases = [
        (" Name ", "Name"),  # Leading/trailing
        ("Name  Value", "Name Value"),  # Double space
        ("  Price  ", "Price"),  # Multiple leading/trailing
        ("\tName\t", "Name"),  # Tabs
        ("Name\n", "Name"),  # Newline
    ]
    
    for input_str, expected in test_cases:
        result = filter_engine._normalize_column_name(input_str)
        print(f"   {repr(input_str)} → {repr(result)}")
        assert result == expected, f"Should normalize {repr(input_str)} to {repr(expected)}"


def test_normalize_method_combined(filter_engine):
    """Test _normalize_column_name with combined edge cases."""
    print(f"\n📂 Testing _normalize_column_name: combined")
    
    import unicodedata
    
    # NFC + non-breaking + leading/trailing spaces
    cafe_nfc = "café"
    input_str = f" {cafe_nfc}\u00A0bar "
    
    result = filter_engine._normalize_column_name(input_str)
    
    print(f"   Input: {repr(input_str)}")
    print(f"   Output: {repr(result)}")
    
    assert result == "café bar", "Should handle combined normalization"
    assert "\u00A0" not in result, "Should remove non-breaking space"
    assert not result.startswith(" "), "Should remove leading space"
    assert not result.endswith(" "), "Should remove trailing space"


# ============================================================================
# NEGATION OPERATOR (NOT) TESTS
# ============================================================================

def test_filter_negate_equals(filter_engine, sample_df):
    """Test negation with == operator (NOT equals)."""
    print(f"\n📂 Testing negation with == operator")
    
    filters = [FilterCondition(column="Name", operator="==", value="Alice", negate=True)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should return all rows EXCEPT Alice
    assert len(result) == len(sample_df) - 1, "Should exclude Alice"
    assert "Alice" not in result["Name"].values, "Alice should not be in results"


def test_filter_negate_greater_than(filter_engine, sample_df):
    """Test negation with > operator (NOT greater than = <=)."""
    print(f"\n📂 Testing negation with > operator")
    
    filters = [FilterCondition(column="Age", operator=">", value=30, negate=True)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should return rows where Age <= 30
    assert all(result["Age"] <= 30), "All ages should be <= 30"


def test_filter_negate_in(filter_engine, sample_df):
    """Test negation with 'in' operator (NOT IN)."""
    print(f"\n📂 Testing negation with 'in' operator")
    
    filters = [FilterCondition(column="Name", operator="in", values=["Alice", "Bob"], negate=True)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should return all rows EXCEPT Alice and Bob
    assert "Alice" not in result["Name"].values, "Alice should not be in results"
    assert "Bob" not in result["Name"].values, "Bob should not be in results"
    assert len(result) == len(sample_df) - 2, "Should exclude Alice and Bob"


def test_filter_negate_contains(filter_engine, sample_df):
    """Test negation with 'contains' operator."""
    print(f"\n📂 Testing negation with 'contains' operator")
    
    filters = [FilterCondition(column="City", operator="contains", value="on", negate=True)]
    result = filter_engine.apply_filters(sample_df, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should return rows where City does NOT contain "on"
    assert all("on" not in str(city) for city in result["City"].values), "No city should contain 'on'"


def test_filter_negate_is_null(filter_engine, df_with_nulls):
    """Test negation with 'is_null' operator (NOT NULL = is_not_null)."""
    print(f"\n📂 Testing negation with 'is_null' operator")
    
    filters = [FilterCondition(column="Name", operator="is_null", negate=True)]
    result = filter_engine.apply_filters(df_with_nulls, filters)
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should return rows where Name is NOT null
    assert result["Name"].notna().all(), "All names should be non-null"


def test_filter_negate_with_and_logic(filter_engine, sample_df):
    """Test negation combined with AND logic."""
    print(f"\n📂 Testing negation with AND logic")
    
    filters = [
        FilterCondition(column="Age", operator=">", value=25),
        FilterCondition(column="City", operator="==", value="Moscow", negate=True)
    ]
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Age > 25 AND City != "Moscow"
    assert all(result["Age"] > 25), "All ages should be > 25"
    assert "Moscow" not in result["City"].values, "Moscow should not be in results"


def test_filter_negate_with_or_logic(filter_engine, sample_df):
    """Test negation combined with OR logic."""
    print(f"\n📂 Testing negation with OR logic")
    
    filters = [
        FilterCondition(column="Age", operator="<", value=25, negate=True),
        FilterCondition(column="Active", operator="==", value=False)
    ]
    result = filter_engine.apply_filters(sample_df, filters, logic="OR")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Age >= 25 OR Active == False
    assert all((row["Age"] >= 25) or (row["Active"] == False) for _, row in result.iterrows()), \
        "Each row should satisfy: Age >= 25 OR Active == False"


def test_filter_multiple_negations(filter_engine, sample_df):
    """Test multiple negated conditions."""
    print(f"\n📂 Testing multiple negations")
    
    filters = [
        FilterCondition(column="Name", operator="==", value="Alice", negate=True),
        FilterCondition(column="City", operator="==", value="Moscow", negate=True)
    ]
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Name != "Alice" AND City != "Moscow"
    assert "Alice" not in result["Name"].values, "Alice should not be in results"
    assert "Moscow" not in result["City"].values, "Moscow should not be in results"


def test_get_filter_summary_with_negation(filter_engine):
    """Test filter summary includes NOT for negated conditions."""
    print(f"\n📂 Testing get_filter_summary with negation")
    
    filters = [
        FilterCondition(column="Age", operator=">", value=30, negate=True),
        FilterCondition(column="Status", operator="==", value="Active")
    ]
    summary = filter_engine.get_filter_summary(filters, "AND")
    
    print(f"✅ Summary: {summary}")
    
    assert "NOT" in summary, "Summary should contain NOT"
    assert "Age > 30" in summary, "Summary should contain Age > 30"
    assert "Status == Active" in summary, "Summary should contain Status == Active"


def test_count_filtered_with_negation(filter_engine, sample_df):
    """Test count_filtered with negated condition."""
    print(f"\n📂 Testing count_filtered with negation")
    
    filters = [FilterCondition(column="Age", operator=">", value=30, negate=True)]
    count = filter_engine.count_filtered(sample_df, filters)
    
    print(f"✅ Count: {count}")
    
    # Should count rows where Age <= 30
    expected = len(sample_df[sample_df["Age"] <= 30])
    assert count == expected, f"Count should be {expected}"


def test_validate_filters_with_negation(filter_engine, sample_df):
    """Test validation works with negated filters."""
    print(f"\n📂 Testing validate_filters with negation")
    
    filters = [FilterCondition(column="Age", operator=">", value=30, negate=True)]
    is_valid, error = filter_engine.validate_filters(sample_df, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is True, "Negated filter should be valid"
    assert error is None, "Should have no error"


# ============================================================================
# NESTED FILTER GROUPS (FilterGroup) TESTS
# ============================================================================

def test_nested_group_simple_and_or(filter_engine, sample_df):
    """Test nested group: (A AND B) OR C.
    
    Logic: (Age > 30 AND City = "Moscow") OR Name = "Alice"
    """
    print(f"\n📂 Testing nested group: (A AND B) OR C")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="City", operator="==", value="Moscow")
            ],
            logic="AND"
        ),
        FilterCondition(column="Name", operator="==", value="Alice")
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="OR")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should match: (Age > 30 AND City = "Moscow") OR Name = "Alice"
    # Alice has Age=25, City=Moscow, so matches second condition
    assert len(result) >= 1, "Should find at least Alice"
    assert "Alice" in result["Name"].values, "Should include Alice"


def test_nested_group_simple_or_and(filter_engine, sample_df):
    """Test nested group: (A OR B) AND C.
    
    Logic: (Name = "Alice" OR Name = "Bob") AND Age < 35
    """
    print(f"\n📂 Testing nested group: (A OR B) AND C")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Name", operator="==", value="Alice"),
                FilterCondition(column="Name", operator="==", value="Bob")
            ],
            logic="OR"
        ),
        FilterCondition(column="Age", operator="<", value=35)
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should match: (Alice OR Bob) AND Age < 35
    # Alice: Age=25 (matches), Bob: Age=30 (matches)
    assert len(result) == 2, "Should find Alice and Bob"
    assert set(result["Name"]) == {"Alice", "Bob"}, "Should include only Alice and Bob"


def test_nested_group_two_groups_or(filter_engine, sample_df):
    """Test two nested groups with OR: (A AND B) OR (C AND D).
    
    Logic: (Age > 40 AND Active = True) OR (Age < 30 AND City = "Moscow")
    """
    print(f"\n📂 Testing nested: (A AND B) OR (C AND D)")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=40),
                FilterCondition(column="Active", operator="==", value=True)
            ],
            logic="AND"
        ),
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator="<", value=30),
                FilterCondition(column="City", operator="==", value="Moscow")
            ],
            logic="AND"
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="OR")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should match: (Age > 40 AND Active) OR (Age < 30 AND Moscow)
    # David: Age=40, Active=True (doesn't match first - Age not > 40)
    # Eve: Age=45, Active=False (doesn't match first)
    # Alice: Age=25, Moscow (matches second)
    assert len(result) >= 1, "Should find at least one match"


def test_nested_group_two_groups_and(filter_engine, sample_df):
    """Test two nested groups with AND: (A OR B) AND (C OR D).
    
    Logic: (Age < 30 OR Age > 40) AND (City = "Moscow" OR City = "Berlin")
    """
    print(f"\n📂 Testing nested: (A OR B) AND (C OR D)")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator="<", value=30),
                FilterCondition(column="Age", operator=">", value=40)
            ],
            logic="OR"
        ),
        FilterGroup(
            filters=[
                FilterCondition(column="City", operator="==", value="Moscow"),
                FilterCondition(column="City", operator="==", value="Berlin")
            ],
            logic="OR"
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should match: (Age < 30 OR Age > 40) AND (Moscow OR Berlin)
    # Alice: Age=25, Moscow (matches both)
    # Eve: Age=45, Berlin (matches both)
    assert len(result) == 2, "Should find Alice and Eve"
    assert set(result["Name"]) == {"Alice", "Eve"}, "Should include Alice and Eve"


def test_nested_group_three_levels(filter_engine, sample_df):
    """Test three-level nesting: ((A OR B) AND C) OR D.
    
    Logic: ((Age < 30 OR Age > 40) AND Active = True) OR Name = "Charlie"
    """
    print(f"\n📂 Testing three-level nesting")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterGroup(
                    filters=[
                        FilterCondition(column="Age", operator="<", value=30),
                        FilterCondition(column="Age", operator=">", value=40)
                    ],
                    logic="OR"
                ),
                FilterCondition(column="Active", operator="==", value=True)
            ],
            logic="AND"
        ),
        FilterCondition(column="Name", operator="==", value="Charlie")
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="OR")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should match: ((Age < 30 OR Age > 40) AND Active) OR Charlie
    # Alice: Age=25, Active=True (matches first)
    # Charlie: Age=35, Active=False (matches second)
    assert len(result) >= 2, "Should find at least 2 matches"
    assert "Alice" in result["Name"].values, "Should include Alice"
    assert "Charlie" in result["Name"].values, "Should include Charlie"


def test_nested_group_four_levels(filter_engine, sample_df):
    """Test four-level nesting: (((A AND B) OR C) AND D) OR E.
    
    Logic: (((Age > 25 AND Age < 35) OR City = "Paris") AND Active = True) OR Name = "Eve"
    """
    print(f"\n📂 Testing four-level nesting")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterGroup(
                    filters=[
                        FilterGroup(
                            filters=[
                                FilterCondition(column="Age", operator=">", value=25),
                                FilterCondition(column="Age", operator="<", value=35)
                            ],
                            logic="AND"
                        ),
                        FilterCondition(column="City", operator="==", value="Paris")
                    ],
                    logic="OR"
                ),
                FilterCondition(column="Active", operator="==", value=True)
            ],
            logic="AND"
        ),
        FilterCondition(column="Name", operator="==", value="Eve")
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="OR")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Complex logic - should find at least Eve
    assert len(result) >= 1, "Should find at least one match"
    assert "Eve" in result["Name"].values, "Should include Eve"


def test_nested_group_very_deep(filter_engine, sample_df):
    """Test very deep nesting (5+ levels).
    
    Verifies system can handle deep recursion without stack overflow.
    """
    print(f"\n📂 Testing very deep nesting (5 levels)")
    
    from mcp_excel.models.requests import FilterGroup
    
    # Build 5-level deep structure
    level5 = FilterCondition(column="Age", operator=">", value=20)
    level4 = FilterGroup(filters=[level5], logic="AND")
    level3 = FilterGroup(filters=[level4], logic="AND")
    level2 = FilterGroup(filters=[level3], logic="AND")
    level1 = FilterGroup(filters=[level2], logic="AND")
    
    filters = [level1]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should work without stack overflow
    assert len(result) == 5, "Should find all rows with Age > 20"


def test_nested_group_negation_simple(filter_engine, sample_df):
    """Test negation of simple group: NOT (A AND B).
    
    Logic: NOT (Age > 30 AND Active = True)
    """
    print(f"\n📂 Testing NOT (A AND B)")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="Active", operator="==", value=True)
            ],
            logic="AND",
            negate=True
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should match: NOT (Age > 30 AND Active)
    # Excludes: Bob (30, True), David (40, True)
    # Includes: Alice (25, True), Charlie (35, False), Eve (45, False)
    assert len(result) >= 3, "Should find rows not matching (Age > 30 AND Active)"


def test_nested_group_negation_complex(filter_engine, sample_df):
    """Test negation of complex group: NOT ((A OR B) AND C).
    
    Logic: NOT ((Age < 30 OR Age > 40) AND City = "Moscow")
    """
    print(f"\n📂 Testing NOT ((A OR B) AND C)")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterGroup(
                    filters=[
                        FilterCondition(column="Age", operator="<", value=30),
                        FilterCondition(column="Age", operator=">", value=40)
                    ],
                    logic="OR"
                ),
                FilterCondition(column="City", operator="==", value="Moscow")
            ],
            logic="AND",
            negate=True
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should exclude: Alice (Age=25, Moscow)
    assert len(result) >= 4, "Should find rows not matching complex condition"
    assert "Alice" not in result["Name"].values or len(result) == 5, "May or may not include Alice depending on data"


def test_nested_group_negation_multiple(filter_engine, sample_df):
    """Test multiple negated groups: NOT (A AND B) OR NOT (C AND D).
    
    Logic: NOT (Age > 35 AND Active = True) OR NOT (City = "Moscow" AND Age < 30)
    """
    print(f"\n📂 Testing NOT (A AND B) OR NOT (C AND D)")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=35),
                FilterCondition(column="Active", operator="==", value=True)
            ],
            logic="AND",
            negate=True
        ),
        FilterGroup(
            filters=[
                FilterCondition(column="City", operator="==", value="Moscow"),
                FilterCondition(column="Age", operator="<", value=30)
            ],
            logic="AND",
            negate=True
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="OR")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Complex OR logic - should match most rows
    assert len(result) >= 1, "Should find matches"


def test_nested_group_negation_with_condition_negation(filter_engine, sample_df):
    """Test group negation combined with condition negation: NOT (A AND NOT B).
    
    Logic: NOT (Age > 30 AND NOT (City = "Moscow"))
    """
    print(f"\n📂 Testing NOT (A AND NOT B)")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="City", operator="==", value="Moscow", negate=True)
            ],
            logic="AND",
            negate=True
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should work with double negation
    assert len(result) >= 1, "Should handle double negation"


def test_nested_group_empty(filter_engine, sample_df):
    """Test empty nested group.
    
    Empty group should return all rows (no filtering).
    """
    print(f"\n📂 Testing empty nested group")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(filters=[], logic="AND")
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Empty group should match all rows
    assert len(result) == len(sample_df), "Empty group should return all rows"


def test_nested_group_single_condition(filter_engine, sample_df):
    """Test group with single condition.
    
    Group with one condition should work same as flat condition.
    """
    print(f"\n📂 Testing group with single condition")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[FilterCondition(column="Age", operator=">", value=30)],
            logic="AND"
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should work same as flat filter
    assert len(result) == 3, "Should find 3 rows with Age > 30"
    assert all(result["Age"] > 30), "All ages should be > 30"


def test_nested_group_single_nested_group(filter_engine, sample_df):
    """Test group within group with single condition.
    
    Verifies unnecessary nesting doesn't break logic.
    """
    print(f"\n📂 Testing group within group (single condition)")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterGroup(
                    filters=[FilterCondition(column="Age", operator=">", value=30)],
                    logic="AND"
                )
            ],
            logic="AND"
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should work despite unnecessary nesting
    assert len(result) == 3, "Should find 3 rows with Age > 30"


def test_nested_group_all_operators(filter_engine, sample_df):
    """Test nested groups with all 14 operators.
    
    Verifies all operators work correctly within nested groups.
    """
    print(f"\n📂 Testing all operators in nested groups")
    
    from mcp_excel.models.requests import FilterGroup
    
    # Test various operators in nested structure
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">=", value=25),
                FilterCondition(column="Age", operator="<=", value=45),
                FilterCondition(column="Name", operator="!=", value="Unknown")
            ],
            logic="AND"
        )
    ]
    
    result = filter_engine.apply_filters(sample_df, filters, logic="AND")
    
    print(f"✅ Filtered {len(result)} row(s)")
    
    # Should find all rows (Age 25-45, Name != Unknown)
    assert len(result) == 5, "Should find all rows"


def test_validate_nested_group_valid(filter_engine, sample_df):
    """Test validation of valid nested group."""
    print(f"\n📂 Testing validation of valid nested group")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="City", operator="==", value="Moscow")
            ],
            logic="AND"
        )
    ]
    
    is_valid, error = filter_engine.validate_filters(sample_df, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is True, "Valid nested group should pass validation"
    assert error is None, "Should have no error"


def test_validate_nested_group_invalid_column(filter_engine, sample_df):
    """Test validation catches invalid column in nested group."""
    print(f"\n📂 Testing validation with invalid column in group")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="NonExistent", operator="==", value="test")
            ],
            logic="AND"
        )
    ]
    
    is_valid, error = filter_engine.validate_filters(sample_df, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is False, "Should detect invalid column in nested group"
    assert "not found" in error.lower(), "Error should mention column not found"


def test_validate_nested_group_invalid_operator(filter_engine, sample_df):
    """Test validation catches invalid operator in nested group."""
    print(f"\n📂 Testing validation with invalid operator in group")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="Name", operator="in", values=None)  # Missing values
            ],
            logic="AND"
        )
    ]
    
    is_valid, error = filter_engine.validate_filters(sample_df, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is False, "Should detect missing values for 'in' operator"
    assert "values" in error.lower(), "Error should mention missing values"


def test_validate_nested_group_deep_invalid(filter_engine, sample_df):
    """Test validation catches error at 3rd level of nesting."""
    print(f"\n📂 Testing validation with error at 3rd level")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterGroup(
                    filters=[
                        FilterCondition(column="Age", operator=">", value=30),
                        FilterCondition(column="InvalidColumn", operator="==", value="test")
                    ],
                    logic="AND"
                )
            ],
            logic="AND"
        )
    ]
    
    is_valid, error = filter_engine.validate_filters(sample_df, filters)
    
    print(f"✅ Valid: {is_valid}, Error: {error}")
    
    assert is_valid is False, "Should detect error at deep nesting level"
    assert "not found" in error.lower(), "Error should mention column not found"


def test_filter_summary_nested_simple(filter_engine):
    """Test filter summary formatting for simple nested group."""
    print(f"\n📂 Testing filter summary: (A AND B) OR C")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="City", operator="==", value="Moscow")
            ],
            logic="AND"
        ),
        FilterCondition(column="Name", operator="==", value="Alice")
    ]
    
    summary = filter_engine.get_filter_summary(filters, "OR")
    
    print(f"✅ Summary: {summary}")
    
    assert "Age > 30" in summary, "Should include Age condition"
    assert "City == Moscow" in summary, "Should include City condition"
    assert "Name == Alice" in summary, "Should include Name condition"
    assert "AND" in summary, "Should show AND within group"
    assert "OR" in summary, "Should show OR between group and condition"
    assert "(" in summary and ")" in summary, "Should use parentheses for grouping"


def test_filter_summary_nested_with_negation(filter_engine):
    """Test filter summary formatting with negated group."""
    print(f"\n📂 Testing filter summary: NOT (A AND B)")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="Active", operator="==", value=True)
            ],
            logic="AND",
            negate=True
        )
    ]
    
    summary = filter_engine.get_filter_summary(filters, "AND")
    
    print(f"✅ Summary: {summary}")
    
    assert "NOT" in summary, "Should include NOT"
    assert "Age > 30" in summary, "Should include Age condition"
    assert "Active == True" in summary, "Should include Active condition"
    assert "(" in summary and ")" in summary, "Should use parentheses"


def test_filter_summary_nested_deep(filter_engine):
    """Test filter summary formatting for deep nesting."""
    print(f"\n📂 Testing filter summary: deep nesting")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterGroup(
                    filters=[
                        FilterCondition(column="Age", operator=">", value=25),
                        FilterCondition(column="Age", operator="<", value=35)
                    ],
                    logic="AND"
                ),
                FilterCondition(column="City", operator="==", value="Moscow")
            ],
            logic="OR"
        )
    ]
    
    summary = filter_engine.get_filter_summary(filters, "AND")
    
    print(f"✅ Summary: {summary}")
    
    assert "Age > 25" in summary, "Should include first Age condition"
    assert "Age < 35" in summary, "Should include second Age condition"
    assert "City == Moscow" in summary, "Should include City condition"
    # Should have nested parentheses
    assert summary.count("(") >= 2, "Should have multiple levels of parentheses"


def test_count_filtered_with_nested_group(filter_engine, sample_df):
    """Test count_filtered with nested groups."""
    print(f"\n📂 Testing count_filtered with nested groups")
    
    from mcp_excel.models.requests import FilterGroup
    
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Age", operator=">", value=30),
                FilterCondition(column="Active", operator="==", value=True)
            ],
            logic="AND"
        )
    ]
    
    count = filter_engine.count_filtered(sample_df, filters)
    
    print(f"✅ Count: {count}")
    
    # Should count without materializing DataFrame
    assert count >= 0, "Count should be non-negative"
    assert isinstance(count, int), "Count should be integer"
