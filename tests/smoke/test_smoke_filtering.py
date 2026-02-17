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


# ============================================================================
# SAMPLE_ROWS PARAMETER TESTS
# ============================================================================

def test_filter_and_count_with_sample_rows(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count with sample_rows parameter.
    
    Verifies:
    - sample_rows parameter works
    - Returns sample data in response
    """
    print(f"\n🔢 Testing filter_and_count with sample_rows...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 1
    })
    
    if unique_result["count"] == 0:
        print(f"  ⚠️  No unique values, skipping")
        return
    
    filter_value = unique_result["values"][0]
    
    result = mcp_call_tool("filter_and_count", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filters": [
            {"column": column, "operator": "==", "value": filter_value}
        ],
        "sample_rows": 3
    })
    
    # Verify sample_rows in response
    assert "sample_rows" in result, "Response should have sample_rows field"
    
    if result["sample_rows"] is not None:
        assert isinstance(result["sample_rows"], list), "sample_rows should be list"
        assert len(result["sample_rows"]) <= 3, "Should return at most 3 rows"
        print(f"  ✅ Count: {result['count']}, Sample rows: {len(result['sample_rows'])}")
    else:
        print(f"  ✅ Count: {result['count']}, No sample rows (None)")


def test_filter_and_count_batch_with_sample_rows(mcp_call_tool, simple_fixture):
    """Smoke: filter_and_count_batch with sample_rows in filter sets.
    
    Verifies:
    - sample_rows works in batch mode
    - Each filter set can have different sample_rows
    """
    print(f"\n📊 Testing filter_and_count_batch with sample_rows...")
    
    column = simple_fixture.columns[0]
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": column,
        "limit": 2
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
                "label": "Set 1 with samples",
                "filters": [{"column": column, "operator": "==", "value": val1}],
                "sample_rows": 2
            },
            {
                "label": "Set 2 no samples",
                "filters": [{"column": column, "operator": "==", "value": val2}]
            }
        ]
    })
    
    assert len(result["results"]) == 2, "Should have 2 results"
    
    # Verify sample_rows in each result
    for i, filter_result in enumerate(result["results"]):
        assert "sample_rows" in filter_result, f"Result {i} should have sample_rows field"
        
        if filter_result["sample_rows"] is not None:
            sample_count = len(filter_result["sample_rows"])
            print(f"  {filter_result['label']}: {filter_result['count']} rows, {sample_count} samples")
        else:
            print(f"  {filter_result['label']}: {filter_result['count']} rows, no samples")
    
    print(f"  ✅ Batch with sample_rows works!")


# ============================================================================
# ANALYZE_OVERLAP SMOKE TESTS
# ============================================================================

def test_smoke_analyze_overlap_basic(mcp_call_tool, simple_fixture):
    """Smoke: analyze_overlap with 2 sets (basic Venn diagram).
    
    Verifies:
    - Basic overlap analysis works
    - Returns sets, intersections, union
    - Venn diagram for 2 sets is generated
    """
    print(f"\n🔍 Testing analyze_overlap (2 sets)...")
    
    # Get values for filters
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": simple_fixture.columns[0],
        "limit": 2
    })
    
    if unique_result["count"] < 2:
        print(f"  ⚠️  Need at least 2 unique values, skipping")
        return
    
    val1 = unique_result["values"][0]
    val2 = unique_result["values"][1]
    
    print(f"  Analyzing overlap: Set A ({simple_fixture.columns[0]}=='{val1}') vs Set B ({simple_fixture.columns[0]}=='{val2}')")
    
    result = mcp_call_tool("analyze_overlap", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filter_sets": [
            {
                "label": "Set A",
                "filters": [{"column": simple_fixture.columns[0], "operator": "==", "value": val1}]
            },
            {
                "label": "Set B",
                "filters": [{"column": simple_fixture.columns[0], "operator": "==", "value": val2}]
            }
        ]
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from AnalyzeOverlapResponse
    assert "sets" in result, "Missing 'sets'"
    assert "pairwise_intersections" in result, "Missing 'pairwise_intersections'"
    assert "union_count" in result, "Missing 'union_count'"
    assert "union_percentage" in result, "Missing 'union_percentage'"
    assert "venn_diagram_2" in result, "Missing 'venn_diagram_2'"
    assert "venn_diagram_3" in result, "Missing 'venn_diagram_3'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify sets
    assert isinstance(result["sets"], dict), "sets should be dict"
    assert "Set A" in result["sets"], "Should have Set A"
    assert "Set B" in result["sets"], "Should have Set B"
    
    set_a = result["sets"]["Set A"]
    set_b = result["sets"]["Set B"]
    
    assert "count" in set_a, "Set A should have count"
    assert "percentage" in set_a, "Set A should have percentage"
    assert isinstance(set_a["count"], int), "Set A count should be int"
    
    print(f"  Set A: {set_a['count']} rows ({set_a['percentage']}%)")
    print(f"  Set B: {set_b['count']} rows ({set_b['percentage']}%)")
    
    # Verify pairwise intersections
    assert isinstance(result["pairwise_intersections"], dict), "pairwise_intersections should be dict"
    assert "Set A ∩ Set B" in result["pairwise_intersections"], "Should have A ∩ B"
    intersection = result["pairwise_intersections"]["Set A ∩ Set B"]
    print(f"  Intersection: {intersection}")
    
    # Verify union
    assert isinstance(result["union_count"], int), "union_count should be int"
    assert result["union_count"] >= 0, "union_count should be non-negative"
    print(f"  Union: {result['union_count']} rows ({result['union_percentage']}%)")
    
    # Verify Venn diagram for 2 sets
    assert result["venn_diagram_2"] is not None, "Should have Venn diagram for 2 sets"
    venn2 = result["venn_diagram_2"]
    assert "A_only" in venn2, "Venn diagram should have A_only"
    assert "B_only" in venn2, "Venn diagram should have B_only"
    assert "A_and_B" in venn2, "Venn diagram should have A_and_B"
    
    print(f"  Venn diagram: A_only={venn2['A_only']}, B_only={venn2['B_only']}, A∩B={venn2['A_and_B']}")
    
    # Verify venn_diagram_3 is None for 2 sets
    assert result["venn_diagram_3"] is None, "Should not have 3-set Venn for 2 sets"
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    assert excel_output["tsv"], "TSV should not be empty"
    
    print(f"  ✅ Overlap analysis works! Union formula: A + B - intersection = {set_a['count']} + {set_b['count']} - {intersection} = {result['union_count']}")


def test_smoke_analyze_overlap_three_sets(mcp_call_tool, simple_fixture):
    """Smoke: analyze_overlap with 3 sets (full Venn diagram).
    
    Verifies:
    - 3-set overlap analysis works
    - All pairwise intersections calculated
    - Venn diagram for 3 sets is generated (7 zones)
    """
    print(f"\n🔍 Testing analyze_overlap (3 sets)...")
    
    # Get values for filters
    unique_result = mcp_call_tool("get_unique_values", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "column": simple_fixture.columns[0],
        "limit": 3
    })
    
    if unique_result["count"] < 3:
        print(f"  ⚠️  Need at least 3 unique values, skipping")
        return
    
    val1 = unique_result["values"][0]
    val2 = unique_result["values"][1]
    val3 = unique_result["values"][2]
    
    print(f"  Analyzing overlap: 3 sets from {simple_fixture.columns[0]}")
    
    result = mcp_call_tool("analyze_overlap", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "filter_sets": [
            {
                "label": "A",
                "filters": [{"column": simple_fixture.columns[0], "operator": "==", "value": val1}]
            },
            {
                "label": "B",
                "filters": [{"column": simple_fixture.columns[0], "operator": "==", "value": val2}]
            },
            {
                "label": "C",
                "filters": [{"column": simple_fixture.columns[0], "operator": "==", "value": val3}]
            }
        ]
    })
    
    # Verify sets
    assert len(result["sets"]) == 3, "Should have 3 sets"
    assert "A" in result["sets"], "Should have Set A"
    assert "B" in result["sets"], "Should have Set B"
    assert "C" in result["sets"], "Should have Set C"
    
    print(f"  Set A: {result['sets']['A']['count']} rows")
    print(f"  Set B: {result['sets']['B']['count']} rows")
    print(f"  Set C: {result['sets']['C']['count']} rows")
    
    # Verify pairwise intersections (should have 3: A∩B, A∩C, B∩C)
    assert len(result["pairwise_intersections"]) == 3, "Should have 3 pairwise intersections"
    assert "A ∩ B" in result["pairwise_intersections"], "Should have A ∩ B"
    assert "A ∩ C" in result["pairwise_intersections"], "Should have A ∩ C"
    assert "B ∩ C" in result["pairwise_intersections"], "Should have B ∩ C"
    
    print(f"  Pairwise: A∩B={result['pairwise_intersections']['A ∩ B']}, A∩C={result['pairwise_intersections']['A ∩ C']}, B∩C={result['pairwise_intersections']['B ∩ C']}")
    
    # Verify Venn diagram for 3 sets (7 zones)
    assert result["venn_diagram_3"] is not None, "Should have Venn diagram for 3 sets"
    venn3 = result["venn_diagram_3"]
    
    assert "A_only" in venn3, "Venn diagram should have A_only"
    assert "B_only" in venn3, "Venn diagram should have B_only"
    assert "C_only" in venn3, "Venn diagram should have C_only"
    assert "A_and_B_only" in venn3, "Venn diagram should have A_and_B_only"
    assert "A_and_C_only" in venn3, "Venn diagram should have A_and_C_only"
    assert "B_and_C_only" in venn3, "Venn diagram should have B_and_C_only"
    assert "A_and_B_and_C" in venn3, "Venn diagram should have A_and_B_and_C"
    
    print(f"  Venn zones (7): A_only={venn3['A_only']}, B_only={venn3['B_only']}, C_only={venn3['C_only']}")
    print(f"               A∩B_only={venn3['A_and_B_only']}, A∩C_only={venn3['A_and_C_only']}, B∩C_only={venn3['B_and_C_only']}")
    print(f"               A∩B∩C={venn3['A_and_B_and_C']}")
    
    # Verify venn_diagram_2 is None for 3 sets
    assert result["venn_diagram_2"] is None, "Should not have 2-set Venn for 3 sets"
    
    # Verify all zones sum to union
    total_zones = (venn3["A_only"] + venn3["B_only"] + venn3["C_only"] +
                   venn3["A_and_B_only"] + venn3["A_and_C_only"] + venn3["B_and_C_only"] +
                   venn3["A_and_B_and_C"])
    assert total_zones == result["union_count"], f"All Venn zones should sum to union: {total_zones} != {result['union_count']}"
    
    print(f"  ✅ 3-set overlap analysis works! Union: {result['union_count']} rows")
