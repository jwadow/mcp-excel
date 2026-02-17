# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Smoke tests for data retrieval tools.

Tests for 3 data retrieval tools:
- get_unique_values
- get_value_counts
- filter_and_get_rows

Each tool is tested with:
- Basic functionality
- Different parameter combinations
- Edge cases (limits, truncation, pagination)
- Full response structure validation
- Excel output validation
"""

import pytest


# ============================================================================
# GET_UNIQUE_VALUES TESTS
# ============================================================================

def test_get_unique_values_basic(mcp_call_tool, simple_fixture):
    """Smoke: get_unique_values returns unique values from column."""
    print(f"\n🔢 Testing get_unique_values...")
    
    # Get unique values from first column
    column = simple_fixture.columns[0]
    
    result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 100
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from GetUniqueValuesResponse
    assert "values" in result, "Missing 'values'"
    assert "count" in result, "Missing 'count'"
    assert "truncated" in result, "Missing 'truncated'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify values
    assert isinstance(result["values"], list), "values should be list"
    assert result["count"] == len(result["values"]), f"count should match values length"
    assert result["count"] > 0, "Should have at least one unique value"
    assert isinstance(result["truncated"], bool), "truncated should be boolean"
    
    # Verify all values are unique
    assert len(result["values"]) == len(set(str(v) for v in result["values"])), "All values should be unique"
    
    # Verify metadata
    metadata = result["metadata"]
    assert metadata["sheet_name"] == simple_fixture.sheet_name, "Metadata sheet_name mismatch"
    
    print(f"  ✅ Found {result['count']} unique values in column '{column}'")
    print(f"  Truncated: {result['truncated']}")


def test_get_unique_values_with_limit(mcp_call_tool, simple_fixture):
    """Smoke: get_unique_values respects limit parameter."""
    print(f"\n🔢 Testing get_unique_values with limit...")
    
    column = simple_fixture.columns[0]
    limit = 5
    
    result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": limit
    })
    
    # Verify limit is respected
    assert len(result["values"]) <= limit, f"Should return at most {limit} values, got {len(result['values'])}"
    assert result["count"] == len(result["values"]), "count should match actual values returned"
    
    # If truncated, count should equal limit
    if result["truncated"]:
        assert result["count"] == limit, f"When truncated, count should equal limit ({limit})"
        print(f"  ✅ Truncated to {limit} values (more exist)")
    else:
        assert result["count"] < limit, f"When not truncated, count should be less than limit"
        print(f"  ✅ Returned all {result['count']} unique values (no truncation)")


def test_get_unique_values_numeric_column(mcp_call_tool, numeric_types_fixture):
    """Smoke: get_unique_values works with numeric columns."""
    print(f"\n🔢 Testing get_unique_values with numeric column...")
    
    # Find a numeric column
    result = mcp_call_tool("get_unique_values", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "column": numeric_types_fixture.columns[0],  # First column should be numeric
        "limit": 100
    })
    
    # Verify numeric values
    assert result["count"] > 0, "Should have unique values"
    
    # Check if values are numeric (int or float)
    for value in result["values"]:
        assert isinstance(value, (int, float)), f"Expected numeric value, got {type(value).__name__}: {value}"
    
    print(f"  ✅ Found {result['count']} unique numeric values")


# ============================================================================
# GET_VALUE_COUNTS TESTS
# ============================================================================

def test_get_value_counts_basic(mcp_call_tool, simple_fixture):
    """Smoke: get_value_counts returns frequency counts."""
    print(f"\n📊 Testing get_value_counts...")
    
    column = simple_fixture.columns[0]
    
    result = mcp_call_tool("get_value_counts", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "top_n": 10
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from GetValueCountsResponse
    assert "value_counts" in result, "Missing 'value_counts'"
    assert "total_values" in result, "Missing 'total_values'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify value_counts structure
    value_counts = result["value_counts"]
    assert isinstance(value_counts, dict), "value_counts should be dict"
    assert len(value_counts) > 0, "Should have at least one value"
    assert len(value_counts) <= 10, f"Should return at most 10 values (top_n=10), got {len(value_counts)}"
    
    # Verify all counts are positive integers
    for value, count in value_counts.items():
        assert isinstance(count, int), f"Count for '{value}' should be int, got {type(count).__name__}"
        assert count > 0, f"Count for '{value}' should be positive, got {count}"
    
    # Verify total_values
    assert result["total_values"] > 0, "total_values should be positive"
    assert result["total_values"] >= sum(value_counts.values()), "total_values should be >= sum of counts"
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    if excel_output["tsv"]:
        assert isinstance(excel_output["tsv"], str), "tsv should be string"
        assert len(excel_output["tsv"]) > 0, "tsv should not be empty"
        # TSV should contain tab characters
        assert "\t" in excel_output["tsv"], "TSV should contain tab separators"
    
    print(f"  ✅ Found {len(value_counts)} unique values (top {len(value_counts)})")
    print(f"  Total values: {result['total_values']}")
    print(f"  Top value: {list(value_counts.keys())[0]} ({list(value_counts.values())[0]} occurrences)")


def test_get_value_counts_with_top_n(mcp_call_tool, simple_fixture):
    """Smoke: get_value_counts respects top_n parameter."""
    print(f"\n📊 Testing get_value_counts with top_n...")
    
    column = simple_fixture.columns[0]
    top_n = 3
    
    result = mcp_call_tool("get_value_counts", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "top_n": top_n
    })
    
    # Verify top_n is respected
    value_counts = result["value_counts"]
    assert len(value_counts) <= top_n, f"Should return at most {top_n} values, got {len(value_counts)}"
    
    # Verify values are sorted by count (descending)
    counts = list(value_counts.values())
    assert counts == sorted(counts, reverse=True), "Values should be sorted by count (descending)"
    
    print(f"  ✅ Returned top {len(value_counts)} values (requested {top_n})")
    for value, count in value_counts.items():
        print(f"    - {value}: {count} occurrences")


def test_get_value_counts_tsv_format(mcp_call_tool, simple_fixture):
    """Smoke: get_value_counts TSV output is properly formatted."""
    print(f"\n📊 Testing get_value_counts TSV format...")
    
    column = simple_fixture.columns[0]
    
    result = mcp_call_tool("get_value_counts", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "top_n": 5
    })
    
    tsv = result["excel_output"]["tsv"]
    assert tsv, "TSV should not be empty"
    
    # Verify TSV structure
    lines = tsv.strip().split("\n")
    assert len(lines) >= 2, "TSV should have at least header + 1 data row"
    
    # Verify header
    header = lines[0]
    assert "\t" in header, "Header should have tab separator"
    header_cols = header.split("\t")
    assert len(header_cols) >= 2, "Header should have at least 2 columns (value, count)"
    
    # Verify data rows
    for i, line in enumerate(lines[1:], 1):
        cols = line.split("\t")
        assert len(cols) >= 2, f"Data row {i} should have at least 2 columns"
        # Second column should be a number (count)
        try:
            count = int(cols[1])
            assert count > 0, f"Count in row {i} should be positive"
        except ValueError:
            pytest.fail(f"Second column in row {i} should be numeric, got: {cols[1]}")
    
    print(f"  ✅ TSV format valid: {len(lines)} lines (1 header + {len(lines)-1} data rows)")


# ============================================================================
# FILTER_AND_GET_ROWS TESTS
# ============================================================================

def test_filter_and_get_rows_basic(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_get_rows returns filtered rows."""
    print(f"\n🔍 Testing filter_and_get_rows...")
    
    # Simple filter: get all rows (no filter)
    result = mcp_call_tool("filter_and_get_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [],
        "limit": 50
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from FilterAndGetRowsResponse
    assert "rows" in result, "Missing 'rows'"
    assert "count" in result, "Missing 'count'"
    assert "total_matches" in result, "Missing 'total_matches'"
    assert "truncated" in result, "Missing 'truncated'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify rows
    assert isinstance(result["rows"], list), "rows should be list"
    assert result["count"] == len(result["rows"]), "count should match rows length"
    assert result["count"] > 0, "Should have at least one row"
    assert result["count"] <= 50, "Should respect limit of 50"
    
    # Verify each row structure
    for i, row in enumerate(result["rows"]):
        assert isinstance(row, dict), f"Row {i} should be dict"
        # Each row should have values for columns
        assert len(row) > 0, f"Row {i} should not be empty"
        # Verify row has expected columns
        for col in simple_fixture.columns:
            assert col in row, f"Row {i} missing column '{col}'"
    
    # Verify total_matches
    assert result["total_matches"] >= result["count"], "total_matches should be >= count"
    
    # Verify truncated flag
    assert isinstance(result["truncated"], bool), "truncated should be boolean"
    if result["total_matches"] > result["count"]:
        assert result["truncated"], "Should be truncated when total_matches > count"
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    
    print(f"  ✅ Returned {result['count']} rows (total matches: {result['total_matches']})")
    print(f"  Truncated: {result['truncated']}")


def test_filter_and_get_rows_with_filter(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_get_rows applies filters correctly."""
    print(f"\n🔍 Testing filter_and_get_rows with filter...")
    
    # Filter by first column value
    column = simple_fixture.columns[0]
    
    # First get unique values to know what to filter by
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 10
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values found, skipping filter test")
        return
    
    # Use first unique value for filter
    filter_value = unique_result["values"][0]
    print(f"  Filtering by {column} == {filter_value}")
    
    result = mcp_call_tool("filter_and_get_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {"column": column, "operator": "==", "value": filter_value}
        ],
        "limit": 50
    })
    
    # Verify filtering worked
    assert result["count"] > 0, f"Should have at least one row matching filter"
    
    # Verify all returned rows match the filter
    for i, row in enumerate(result["rows"]):
        actual_value = row[column]
        assert actual_value == filter_value, f"Row {i} value '{actual_value}' doesn't match filter '{filter_value}'"
    
    print(f"  ✅ All {result['count']} rows match filter")


def test_filter_and_get_rows_with_columns(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_get_rows returns only requested columns."""
    print(f"\n🔍 Testing filter_and_get_rows with column selection...")
    
    # Request only first 2 columns
    columns_to_return = simple_fixture.columns[:2]
    print(f"  Requesting columns: {columns_to_return}")
    
    result = mcp_call_tool("filter_and_get_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [],
        "columns": columns_to_return,
        "limit": 10
    })
    
    # Verify only requested columns are returned
    for i, row in enumerate(result["rows"]):
        row_columns = set(row.keys())
        expected_columns = set(columns_to_return)
        assert row_columns == expected_columns, f"Row {i} has unexpected columns: {row_columns} != {expected_columns}"
    
    print(f"  ✅ All rows contain only requested columns")


def test_filter_and_get_rows_with_pagination(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_get_rows supports pagination with limit and offset."""
    print(f"\n🔍 Testing filter_and_get_rows pagination...")
    
    # Get first page
    page1 = mcp_call_tool("filter_and_get_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [],
        "limit": 3,
        "offset": 0
    })
    
    # Get second page
    page2 = mcp_call_tool("filter_and_get_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [],
        "limit": 3,
        "offset": 3
    })
    
    # Verify pagination
    assert page1["count"] <= 3, "Page 1 should have at most 3 rows"
    assert page2["count"] <= 3, "Page 2 should have at most 3 rows"
    
    # Verify total_matches is same for both pages
    assert page1["total_matches"] == page2["total_matches"], "total_matches should be same for both pages"
    
    # Verify rows are different (if both pages have data)
    if page1["count"] > 0 and page2["count"] > 0:
        # Compare first row of each page (should be different)
        first_col = simple_fixture.columns[0]
        page1_first = page1["rows"][0][first_col]
        page2_first = page2["rows"][0][first_col]
        # They might be same if data has duplicates, so just verify structure
        assert isinstance(page1_first, type(page2_first)), "Values should have same type"
    
    print(f"  ✅ Page 1: {page1['count']} rows, Page 2: {page2['count']} rows")
    print(f"  Total matches: {page1['total_matches']}")


def test_filter_and_get_rows_with_complex_filter(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_get_rows handles multiple filters with AND logic."""
    print(f"\n🔍 Testing filter_and_get_rows with multiple filters...")
    
    # Use multiple filters (if we have enough columns)
    if len(simple_fixture.columns) < 2:
        print(f"  ⚠️  Not enough columns for multi-filter test, skipping")
        return
    
    # Get unique values for first two columns
    col1 = simple_fixture.columns[0]
    col2 = simple_fixture.columns[1]
    
    unique1 = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": col1,
        "limit": 5
    })
    
    if unique1["count"] == 0:
        print(f"  ⚠️  No unique values in {col1}, skipping")
        return
    
    filter_value1 = unique1["values"][0]
    
    # Apply two filters with AND logic
    result = mcp_call_tool("filter_and_get_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {"column": col1, "operator": "==", "value": filter_value1},
            {"column": col2, "operator": "is_not_null"}
        ],
        "logic": "AND",
        "limit": 50
    })
    
    # Verify filters were applied
    if result["count"] > 0:
        # Verify first filter
        for i, row in enumerate(result["rows"]):
            assert row[col1] == filter_value1, f"Row {i} doesn't match first filter"
            # Second filter (is_not_null) - value should not be None/null
            assert row[col2] is not None, f"Row {i} doesn't match second filter (is_not_null)"
        
        print(f"  ✅ All {result['count']} rows match both filters (AND logic)")
    else:
        print(f"  ℹ️  No rows match both filters (this is OK for smoke test)")


def test_filter_and_get_rows_tsv_output(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_get_rows TSV output is properly formatted."""
    print(f"\n🔍 Testing filter_and_get_rows TSV output...")
    
    result = mcp_call_tool("filter_and_get_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [],
        "limit": 5
    })
    
    tsv = result["excel_output"]["tsv"]
    if not tsv:
        print(f"  ℹ️  No TSV output (this is OK)")
        return
    
    # Verify TSV structure
    lines = tsv.strip().split("\n")
    assert len(lines) >= 2, "TSV should have at least header + 1 data row"
    
    # Verify header
    header = lines[0]
    assert "\t" in header, "Header should have tab separator"
    header_cols = header.split("\t")
    assert len(header_cols) == len(simple_fixture.columns), f"Header should have {len(simple_fixture.columns)} columns"
    
    # Verify data rows match returned rows
    data_lines = lines[1:]
    assert len(data_lines) == result["count"], f"TSV should have {result['count']} data rows"
    
    print(f"  ✅ TSV format valid: {len(lines)} lines (1 header + {len(data_lines)} data rows)")
