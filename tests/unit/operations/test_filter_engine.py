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
