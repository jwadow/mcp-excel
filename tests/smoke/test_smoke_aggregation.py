# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Smoke tests for aggregation tools.

Tests for 2 aggregation tools:
- aggregate (8 operations: sum, mean, median, min, max, std, var, count)
- group_by (pivot table functionality)

Each tool is tested with:
- All aggregation operations
- With and without filters
- Different data types
- Edge cases
- Full response structure validation
- Excel formula generation
"""

import pytest


# ============================================================================
# AGGREGATE TESTS - ALL OPERATIONS
# ============================================================================

def test_aggregate_sum(mcp_call_tool, numeric_types_fixture):
    """Smoke: aggregate with sum operation."""
    print(f"\n➕ Testing aggregate sum...")
    
    # Use numeric column
    column = numeric_types_fixture.columns[0]
    
    result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "sum",
        "target_column": column
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from AggregateResponse
    assert "value" in result, "Missing 'value'"
    assert "operation" in result, "Missing 'operation'"
    assert "target_column" in result, "Missing 'target_column'"
    assert "filters_applied" in result, "Missing 'filters_applied'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify value
    assert isinstance(result["value"], (int, float)), f"value should be numeric, got {type(result['value']).__name__}"
    
    # Verify operation
    assert result["operation"] == "sum", f"Expected operation 'sum', got '{result['operation']}'"
    assert result["target_column"] == column
    
    # Verify Excel formula
    formula = result["excel_output"]["formula"]
    assert formula, "Should have Excel formula"
    assert formula.startswith("="), "Formula should start with ="
    assert "SUM" in formula.upper(), "Should use SUM function"
    
    print(f"  ✅ Sum: {result['value']}, Formula: {formula[:50]}...")


def test_aggregate_mean(mcp_call_tool, numeric_types_fixture):
    """Smoke: aggregate with mean (average) operation."""
    print(f"\n📊 Testing aggregate mean...")
    
    column = numeric_types_fixture.columns[0]
    
    result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "mean",
        "target_column": column
    })
    
    assert isinstance(result["value"], (int, float))
    assert result["operation"] == "mean"
    
    # Verify Excel formula uses AVERAGE
    formula = result["excel_output"]["formula"]
    assert "AVERAGE" in formula.upper(), "Should use AVERAGE function"
    
    print(f"  ✅ Mean: {result['value']}")


def test_aggregate_median(mcp_call_tool, numeric_types_fixture):
    """Smoke: aggregate with median operation."""
    print(f"\n📊 Testing aggregate median...")
    
    column = numeric_types_fixture.columns[0]
    
    result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "median",
        "target_column": column
    })
    
    assert isinstance(result["value"], (int, float))
    assert result["operation"] == "median"
    
    print(f"  ✅ Median: {result['value']}")


def test_aggregate_min_max(mcp_call_tool, numeric_types_fixture):
    """Smoke: aggregate with min and max operations."""
    print(f"\n📊 Testing aggregate min/max...")
    
    column = numeric_types_fixture.columns[0]
    
    # Test min
    min_result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "min",
        "target_column": column
    })
    
    assert isinstance(min_result["value"], (int, float))
    assert min_result["operation"] == "min"
    assert "MIN" in min_result["excel_output"]["formula"].upper()
    
    # Test max
    max_result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "max",
        "target_column": column
    })
    
    assert isinstance(max_result["value"], (int, float))
    assert max_result["operation"] == "max"
    assert "MAX" in max_result["excel_output"]["formula"].upper()
    
    # Verify min <= max
    assert min_result["value"] <= max_result["value"], "Min should be <= Max"
    
    print(f"  ✅ Min: {min_result['value']}, Max: {max_result['value']}")


def test_aggregate_std_var(mcp_call_tool, numeric_types_fixture):
    """Smoke: aggregate with std (standard deviation) and var (variance) operations."""
    print(f"\n📊 Testing aggregate std/var...")
    
    column = numeric_types_fixture.columns[0]
    
    # Test std
    std_result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "std",
        "target_column": column
    })
    
    assert isinstance(std_result["value"], (int, float))
    assert std_result["value"] >= 0, "Std should be non-negative"
    assert std_result["operation"] == "std"
    
    # Test var
    var_result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "var",
        "target_column": column
    })
    
    assert isinstance(var_result["value"], (int, float))
    assert var_result["value"] >= 0, "Var should be non-negative"
    assert var_result["operation"] == "var"
    
    # Verify var = std^2 (approximately)
    expected_var = std_result["value"] ** 2
    assert abs(var_result["value"] - expected_var) < 0.01, f"Var should equal std^2: {var_result['value']} != {expected_var}"
    
    print(f"  ✅ Std: {std_result['value']:.2f}, Var: {var_result['value']:.2f}")


def test_aggregate_count(mcp_call_tool, simple_fixture):
    """Smoke: aggregate with count operation."""
    print(f"\n🔢 Testing aggregate count...")
    
    column = simple_fixture.columns[0]
    
    result = mcp_call_tool("aggregate", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "operation": "count",
        "target_column": column
    })
    
    assert isinstance(result["value"], int), "count should return int"
    assert result["value"] > 0, "count should be positive"
    assert result["operation"] == "count"
    
    # Verify Excel formula uses COUNT or COUNTA
    formula = result["excel_output"]["formula"]
    assert "COUNT" in formula.upper(), "Should use COUNT function"
    
    print(f"  ✅ Count: {result['value']}")


# ============================================================================
# AGGREGATE TESTS - WITH FILTERS
# ============================================================================

def test_aggregate_with_filter(mcp_call_tool, numeric_types_fixture):
    """Smoke: aggregate with filters applied."""
    print(f"\n📊 Testing aggregate with filter...")
    
    column = numeric_types_fixture.columns[0]
    
    # Get unique values for filter
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    filter_value = unique_result["values"][0]
    
    # Aggregate with filter
    result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "sum",
        "target_column": column,
        "filters": [
            {"column": column, "operator": "==", "value": filter_value}
        ]
    })
    
    # Verify filters were applied
    assert len(result["filters_applied"]) == 1, "Should have 1 filter applied"
    assert result["filters_applied"][0]["column"] == column
    
    # Verify Excel formula includes filter (SUMIF)
    formula = result["excel_output"]["formula"]
    assert "SUMIF" in formula.upper() or "SUMIFS" in formula.upper(), "Should use SUMIF/SUMIFS for filtered sum"
    
    print(f"  ✅ Filtered sum: {result['value']}")


def test_aggregate_with_nested_filter(mcp_call_tool, numeric_types_fixture):
    """Smoke: aggregate with nested filter group.
    
    CRITICAL: Verifies nested filters work with aggregate.
    """
    print(f"\n📊 Testing aggregate with nested filter...")
    
    column = numeric_types_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    filter_value = unique_result["values"][0]
    
    result = mcp_call_tool("aggregate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "operation": "count",
        "target_column": column,
        "filters": [
            {
                "filters": [
                    {"column": column, "operator": "==", "value": filter_value}
                ],
                "logic": "AND"
            }
        ]
    })
    
    assert "value" in result
    assert isinstance(result["value"], int)
    
    print(f"  ✅ Nested filter works with aggregate! Count: {result['value']}")


# ============================================================================
# GROUP_BY TESTS
# ============================================================================

def test_group_by_single_column(mcp_call_tool, simple_fixture):
    """Smoke: group_by with single grouping column."""
    print(f"\n📊 Testing group_by with single column...")
    
    if len(simple_fixture.columns) < 2:
        print(f"  ⚠️  Need at least 2 columns, skipping")
        return
    
    group_col = simple_fixture.columns[0]
    agg_col = simple_fixture.columns[1]
    
    result = mcp_call_tool("group_by", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "group_columns": [group_col],
        "agg_column": agg_col,
        "agg_operation": "count"
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from GroupByResponse
    assert "groups" in result, "Missing 'groups'"
    assert "group_columns" in result, "Missing 'group_columns'"
    assert "agg_column" in result, "Missing 'agg_column'"
    assert "agg_operation" in result, "Missing 'agg_operation'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify groups
    assert isinstance(result["groups"], list), "groups should be list"
    assert len(result["groups"]) > 0, "Should have at least one group"
    
    # Verify each group structure
    for i, group in enumerate(result["groups"]):
        assert isinstance(group, dict), f"Group {i} should be dict"
        # Should have group column and aggregated value
        assert group_col in group, f"Group {i} missing group column '{group_col}'"
        # Should have aggregated value (column name varies)
        assert len(group) >= 2, f"Group {i} should have at least 2 fields (group + agg)"
    
    # Verify metadata
    assert result["group_columns"] == [group_col]
    assert result["agg_column"] == agg_col
    assert result["agg_operation"] == "count"
    
    # Verify Excel output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    
    print(f"  ✅ Found {len(result['groups'])} groups")


def test_group_by_multiple_columns(mcp_call_tool, simple_fixture):
    """Smoke: group_by with multiple grouping columns (hierarchical)."""
    print(f"\n📊 Testing group_by with multiple columns...")
    
    if len(simple_fixture.columns) < 3:
        print(f"  ⚠️  Need at least 3 columns, skipping")
        return
    
    group_cols = simple_fixture.columns[:2]
    agg_col = simple_fixture.columns[2]
    
    print(f"  Grouping by: {group_cols}")
    
    result = mcp_call_tool("group_by", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "group_columns": group_cols,
        "agg_column": agg_col,
        "agg_operation": "count"
    })
    
    # Verify multiple group columns
    assert result["group_columns"] == group_cols, f"Expected {group_cols}, got {result['group_columns']}"
    assert len(result["groups"]) > 0, "Should have at least one group"
    
    # Verify each group has all group columns
    for i, group in enumerate(result["groups"]):
        for col in group_cols:
            assert col in group, f"Group {i} missing group column '{col}'"
    
    print(f"  ✅ Hierarchical grouping works! Found {len(result['groups'])} groups")


def test_group_by_all_operations(mcp_call_tool, numeric_types_fixture):
    """Smoke: group_by with different aggregation operations."""
    print(f"\n📊 Testing group_by with different operations...")
    
    if len(numeric_types_fixture.columns) < 2:
        print(f"  ⚠️  Need at least 2 columns, skipping")
        return
    
    group_col = numeric_types_fixture.columns[0]
    agg_col = numeric_types_fixture.columns[1]
    
    operations = ["sum", "mean", "count"]
    
    for operation in operations:
        print(f"  Testing {operation}...")
        
        result = mcp_call_tool("group_by", {
            "file_path": str(numeric_types_fixture.path_str),
            "sheet_name": numeric_types_fixture.sheet_name,
            "group_columns": [group_col],
            "agg_column": agg_col,
            "agg_operation": operation
        })
        
        assert result["agg_operation"] == operation, f"Expected operation '{operation}'"
        assert len(result["groups"]) > 0, f"Should have groups for {operation}"
        
        print(f"    ✅ {operation}: {len(result['groups'])} groups")


def test_group_by_with_filter(mcp_call_tool, simple_fixture):
    """Smoke: group_by with filters applied."""
    print(f"\n📊 Testing group_by with filter...")
    
    if len(simple_fixture.columns) < 2:
        print(f"  ⚠️  Need at least 2 columns, skipping")
        return
    
    group_col = simple_fixture.columns[0]
    agg_col = simple_fixture.columns[1]
    
    result = mcp_call_tool("group_by", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "group_columns": [group_col],
        "agg_column": agg_col,
        "agg_operation": "count",
        "filters": [
            {"column": group_col, "operator": "is_not_null"}
        ]
    })
    
    assert len(result["groups"]) > 0, "Should have groups after filtering"
    
    print(f"  ✅ Filtered grouping works! Found {len(result['groups'])} groups")


def test_group_by_tsv_output(mcp_call_tool, simple_fixture):
    """Smoke: group_by TSV output is properly formatted."""
    print(f"\n📊 Testing group_by TSV output...")
    
    if len(simple_fixture.columns) < 2:
        print(f"  ⚠️  Need at least 2 columns, skipping")
        return
    
    group_col = simple_fixture.columns[0]
    agg_col = simple_fixture.columns[1]
    
    result = mcp_call_tool("group_by", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "group_columns": [group_col],
        "agg_column": agg_col,
        "agg_operation": "count"
    })
    
    tsv = result["excel_output"]["tsv"]
    if not tsv:
        print(f"  ⚠️  No TSV output")
        return
    
    # Verify TSV structure
    lines = tsv.strip().split("\n")
    assert len(lines) >= 2, "TSV should have at least header + 1 data row"
    
    # Verify header
    header = lines[0]
    assert "\t" in header, "Header should have tab separator"
    header_cols = header.split("\t")
    assert len(header_cols) >= 2, "Header should have at least 2 columns (group + agg)"
    
    # Verify data rows match groups
    data_lines = lines[1:]
    assert len(data_lines) == len(result["groups"]), f"TSV should have {len(result['groups'])} data rows"
    
    print(f"  ✅ TSV format valid: {len(lines)} lines (1 header + {len(data_lines)} data rows)")
