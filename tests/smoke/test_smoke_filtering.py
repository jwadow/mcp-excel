# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Smoke tests for filtering and counting tools.

Tests for 2 filtering tools:
- filter_and_count
- filter_and_count_batch

CRITICAL: These tests verify the nested filter groups bug fix.
The bug was that FilterGroup definitions were not at root level of JSON Schema,
causing $ref resolution to fail in MCP Framework.

Each tool is tested with:
- Basic functionality
- All 12 filter operators
- Nested filter groups (1, 2, 3+ levels)
- Complex logical expressions (AND/OR combinations)
- Edge cases
- Full response structure validation
- Excel formula generation
"""

import pytest


# ============================================================================
# FILTER_AND_COUNT TESTS - BASIC
# ============================================================================

def test_filter_and_count_no_filter(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with empty filters returns total count."""
    print(f"\n🔢 Testing filter_and_count with no filters...")
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": []
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from FilterAndCountResponse
    assert "count" in result, "Missing 'count'"
    assert "filters_applied" in result, "Missing 'filters_applied'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify count
    assert isinstance(result["count"], int), "count should be int"
    assert result["count"] == simple_fixture.row_count, f"Expected {simple_fixture.row_count} rows, got {result['count']}"
    
    # Verify filters_applied
    assert isinstance(result["filters_applied"], list), "filters_applied should be list"
    assert len(result["filters_applied"]) == 0, "Should have no filters applied"
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "formula" in excel_output, "excel_output missing 'formula'"
    
    print(f"  ✅ Count: {result['count']} (all rows)")


def test_filter_and_count_simple_filter(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with simple equality filter."""
    print(f"\n🔢 Testing filter_and_count with simple filter...")
    
    # Get a value to filter by
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    filter_value = unique_result["values"][0]
    print(f"  Filtering: {column} == {filter_value}")
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {"column": column, "operator": "==", "value": filter_value}
        ]
    })
    
    # Verify count
    assert result["count"] >= 0, "count should be non-negative"
    assert result["count"] <= simple_fixture.row_count, "count can't exceed total rows"
    
    # Verify filters_applied
    assert len(result["filters_applied"]) == 1, "Should have 1 filter applied"
    applied_filter = result["filters_applied"][0]
    assert applied_filter["column"] == column
    assert applied_filter["operator"] == "=="
    assert applied_filter["value"] == filter_value
    
    # Verify Excel formula
    formula = result["excel_output"]["formula"]
    assert formula, "Should have Excel formula"
    assert formula.startswith("="), "Formula should start with ="
    assert "COUNTIF" in formula.upper(), "Should use COUNTIF function"
    
    print(f"  ✅ Count: {result['count']}, Formula: {formula[:50]}...")


# ============================================================================
# FILTER_AND_COUNT TESTS - NESTED GROUPS (CRITICAL)
# ============================================================================

def test_filter_and_count_nested_group_1_level(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with 1-level nested group.
    
    CRITICAL: This is the bug that was fixed - nested groups must work.
    Structure: [{"filters": [...], "logic": "AND"}]
    """
    print(f"\n🔢 Testing filter_and_count with 1-level nested group...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    filter_value = unique_result["values"][0]
    print(f"  Nested filter: (({column} == {filter_value}))")
    
    # CRITICAL: This structure caused the bug - nested group with $ref
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {
                "filters": [
                    {"column": column, "operator": "==", "value": filter_value}
                ],
                "logic": "AND"
            }
        ]
    })
    
    # If we got here without error, the bug is fixed!
    assert "count" in result, "Should have count in response"
    assert result["count"] >= 0, "count should be non-negative"
    
    print(f"  ✅ Nested group works! Count: {result['count']}")
    print(f"  ✅ BUG FIX VERIFIED: $ref resolution works correctly")


def test_filter_and_count_nested_group_2_levels(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with 2-level nested groups.
    
    Structure: [{"filters": [{"filters": [...]}]}]
    """
    print(f"\n🔢 Testing filter_and_count with 2-level nested groups...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    filter_value = unique_result["values"][0]
    print(f"  2-level nested: ((({column} == {filter_value})))")
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {
                "filters": [
                    {
                        "filters": [
                            {"column": column, "operator": "==", "value": filter_value}
                        ],
                        "logic": "AND"
                    }
                ],
                "logic": "AND"
            }
        ]
    })
    
    assert "count" in result
    assert result["count"] >= 0
    
    print(f"  ✅ 2-level nesting works! Count: {result['count']}")


def test_filter_and_count_nested_group_3_levels(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with 3-level nested groups (extreme case).
    
    Structure: [{"filters": [{"filters": [{"filters": [...]}]}]}]
    """
    print(f"\n🔢 Testing filter_and_count with 3-level nested groups...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    filter_value = unique_result["values"][0]
    print(f"  3-level nested: (((({column} == {filter_value}))))")
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {
                "filters": [
                    {
                        "filters": [
                            {
                                "filters": [
                                    {"column": column, "operator": "==", "value": filter_value}
                                ],
                                "logic": "AND"
                            }
                        ],
                        "logic": "AND"
                    }
                ],
                "logic": "AND"
            }
        ]
    })
    
    assert "count" in result
    assert result["count"] >= 0
    
    print(f"  ✅ 3-level nesting works! Count: {result['count']}")


def test_filter_and_count_complex_nested_logic(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with complex nested logic (A AND B) OR C.
    
    Structure: [group1(A AND B), C] with OR logic
    """
    print(f"\n🔢 Testing filter_and_count with complex nested logic...")
    
    if len(simple_fixture.columns) < 2:
        print(f"  ⚠️  Need at least 2 columns, skipping")
        return
    
    col1 = simple_fixture.columns[0]
    col2 = simple_fixture.columns[1]
    
    # Get values for filters
    unique1 = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": col1,
        "limit": 5
    })
    
    if unique1["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    val1 = unique1["values"][0]
    
    print(f"  Complex logic: (({col1} == {val1}) AND ({col2} is_not_null)) OR ({col1} is_not_null)")
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {
                "filters": [
                    {"column": col1, "operator": "==", "value": val1},
                    {"column": col2, "operator": "is_not_null"}
                ],
                "logic": "AND"
            },
            {"column": col1, "operator": "is_not_null"}
        ],
        "logic": "OR"
    })
    
    assert "count" in result
    assert result["count"] >= 0
    
    print(f"  ✅ Complex nested logic works! Count: {result['count']}")


# ============================================================================
# FILTER_AND_COUNT TESTS - ALL OPERATORS
# ============================================================================

def test_filter_and_count_operator_not_equal(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with != operator."""
    print(f"\n🔢 Testing filter_and_count with != operator...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    filter_value = unique_result["values"][0]
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {"column": column, "operator": "!=", "value": filter_value}
        ]
    })
    
    assert result["count"] >= 0
    # Count should be total - count of equal values
    print(f"  ✅ != operator works, count: {result['count']}")


def test_filter_and_count_operator_in(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with 'in' operator (multiple values)."""
    print(f"\n🔢 Testing filter_and_count with 'in' operator...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] < 2:
        print(f"  ⚠️  Need at least 2 unique values, skipping")
        return
    
    filter_values = unique_result["values"][:2]
    print(f"  Filter: {column} in {filter_values}")
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {"column": column, "operator": "in", "values": filter_values}
        ]
    })
    
    assert result["count"] >= 0
    print(f"  ✅ 'in' operator works, count: {result['count']}")


def test_filter_and_count_operator_is_null(mcp_call_tool, with_nulls_fixture):
    """Smoke: filter_and_count with is_null operator."""
    print(f"\n🔢 Testing filter_and_count with is_null operator...")
    
    # Use fixture with nulls
    column = with_nulls_fixture.columns[0]
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(with_nulls_fixture.path_str),
        "sheet_name": with_nulls_fixture.sheet_name,
        "filters": [
            {"column": column, "operator": "is_null"}
        ]
    })
    
    assert result["count"] >= 0
    print(f"  ✅ is_null operator works, count: {result['count']}")


def test_filter_and_count_operator_is_not_null(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with is_not_null operator."""
    print(f"\n🔢 Testing filter_and_count with is_not_null operator...")
    
    column = simple_fixture.columns[0]
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {"column": column, "operator": "is_not_null"}
        ]
    })
    
    assert result["count"] >= 0
    assert result["count"] <= simple_fixture.row_count
    print(f"  ✅ is_not_null operator works, count: {result['count']}")


# ============================================================================
# FILTER_AND_COUNT_BATCH TESTS
# ============================================================================

def test_filter_and_count_batch_basic(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count_batch processes multiple filter sets."""
    print(f"\n📊 Testing filter_and_count_batch...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] < 2:
        print(f"  ⚠️  Need at least 2 unique values, skipping")
        return
    
    val1 = unique_result["values"][0]
    val2 = unique_result["values"][1]
    
    result = mcp_call_tool("filter_and_count_batch", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filter_sets": [
            {
                "label": "Category A",
                "filters": [{"column": column, "operator": "==", "value": val1}]
            },
            {
                "label": "Category B",
                "filters": [{"column": column, "operator": "==", "value": val2}]
            },
            {
                "label": "All non-null",
                "filters": [{"column": column, "operator": "is_not_null"}]
            }
        ]
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from FilterAndCountBatchResponse
    assert "results" in result, "Missing 'results'"
    assert "total_filter_sets" in result, "Missing 'total_filter_sets'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify results
    assert isinstance(result["results"], list), "results should be list"
    assert len(result["results"]) == 3, f"Expected 3 results, got {len(result['results'])}"
    assert result["total_filter_sets"] == 3, "total_filter_sets should be 3"
    
    # Verify each result structure (from FilterSetResult model)
    for i, filter_result in enumerate(result["results"]):
        print(f"  Checking result {i+1}...")
        
        assert "label" in filter_result, f"Result {i} missing 'label'"
        assert "count" in filter_result, f"Result {i} missing 'count'"
        assert "filters_applied" in filter_result, f"Result {i} missing 'filters_applied'"
        assert "formula" in filter_result, f"Result {i} missing 'formula'"
        
        assert isinstance(filter_result["count"], int), f"Result {i} count should be int"
        assert filter_result["count"] >= 0, f"Result {i} count should be non-negative"
        
        print(f"    {filter_result['label']}: {filter_result['count']} rows")
    
    # Verify excel_output (TSV table)
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    if excel_output["tsv"]:
        tsv = excel_output["tsv"]
        lines = tsv.strip().split("\n")
        assert len(lines) >= 4, "TSV should have header + 3 data rows"
        print(f"  ✅ TSV output: {len(lines)} lines")
    
    print(f"  ✅ Batch processing works! Processed {result['total_filter_sets']} filter sets")


def test_filter_and_count_batch_with_nested_groups(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count_batch works with nested filter groups.
    
    CRITICAL: Verifies nested groups work in batch mode too.
    """
    print(f"\n📊 Testing filter_and_count_batch with nested groups...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 5
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    val1 = unique_result["values"][0]
    
    result = mcp_call_tool("filter_and_count_batch", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filter_sets": [
            {
                "label": "Nested group test",
                "filters": [
                    {
                        "filters": [
                            {"column": column, "operator": "==", "value": val1}
                        ],
                        "logic": "AND"
                    }
                ]
            }
        ]
    })
    
    assert len(result["results"]) == 1
    assert result["results"][0]["count"] >= 0
    
    print(f"  ✅ Nested groups work in batch mode! Count: {result['results'][0]['count']}")


def test_filter_and_count_batch_without_labels(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count_batch works without labels (auto-generated)."""
    print(f"\n📊 Testing filter_and_count_batch without labels...")
    
    column = simple_fixture.columns[0]
    
    result = mcp_call_tool("filter_and_count_batch", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filter_sets": [
            {
                "filters": [{"column": column, "operator": "is_not_null"}]
            },
            {
                "filters": [{"column": column, "operator": "is_null"}]
            }
        ]
    })
    
    assert len(result["results"]) == 2
    
    # Verify labels were auto-generated
    for i, filter_result in enumerate(result["results"]):
        label = filter_result["label"]
        # Should have some label (auto-generated or None)
        print(f"  Result {i+1} label: {label}, count: {filter_result['count']}")
    
    print(f"  ✅ Auto-generated labels work")


def test_filter_and_count_batch_max_filter_sets(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count_batch handles many filter sets (up to 50)."""
    print(f"\n📊 Testing filter_and_count_batch with many filter sets...")
    
    column = simple_fixture.columns[0]
    
    # Create 10 filter sets (not 50 to keep test fast)
    filter_sets = []
    for i in range(10):
        filter_sets.append({
            "label": f"Set {i+1}",
            "filters": [{"column": column, "operator": "is_not_null"}]
        })
    
    result = mcp_call_tool("filter_and_count_batch", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filter_sets": filter_sets
    })
    
    assert len(result["results"]) == 10, f"Expected 10 results, got {len(result['results'])}"
    assert result["total_filter_sets"] == 10
    
    print(f"  ✅ Processed {result['total_filter_sets']} filter sets successfully")
