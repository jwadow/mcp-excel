# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Smoke tests for statistics, validation, and multi-sheet tools.

Tests for 7 tools:
- Statistics (3): get_column_stats, correlate, detect_outliers
- Validation (2): find_duplicates, find_nulls
- Multi-sheet (2): search_across_sheets, compare_sheets

Each tool is tested with full response validation and edge cases.
"""

import pytest


# ============================================================================
# STATISTICS TESTS - GET_COLUMN_STATS
# ============================================================================

def test_get_column_stats_basic(mcp_call_tool, numeric_types_fixture):
    """Smoke: get_column_stats returns complete statistical summary."""
    print(f"\n📊 Testing get_column_stats...")
    
    column = numeric_types_fixture.columns[0]
    
    result = mcp_call_tool("get_column_stats", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "column": column
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from GetColumnStatsResponse
    assert "column" in result, "Missing 'column'"
    assert "stats" in result, "Missing 'stats'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify column
    assert result["column"] == column
    
    # Verify stats structure (from ColumnStats model)
    stats = result["stats"]
    assert "count" in stats, "stats missing 'count'"
    assert "null_count" in stats, "stats missing 'null_count'"
    
    # For numeric columns, should have statistical measures
    if "mean" in stats and stats["mean"] is not None:
        assert isinstance(stats["mean"], (int, float)), "mean should be numeric"
    if "median" in stats and stats["median"] is not None:
        assert isinstance(stats["median"], (int, float)), "median should be numeric"
    if "std" in stats and stats["std"] is not None:
        assert isinstance(stats["std"], (int, float)), "std should be numeric"
        assert stats["std"] >= 0, "std should be non-negative"
    if "min" in stats and stats["min"] is not None:
        assert isinstance(stats["min"], (int, float)), "min should be numeric"
    if "max" in stats and stats["max"] is not None:
        assert isinstance(stats["max"], (int, float)), "max should be numeric"
    
    # If both min and max exist, verify min <= max
    if stats.get("min") is not None and stats.get("max") is not None:
        assert stats["min"] <= stats["max"], "min should be <= max"
    
    print(f"  ✅ Stats: count={stats['count']}, mean={stats.get('mean')}, median={stats.get('median')}")


# ============================================================================
# STATISTICS TESTS - CORRELATE
# ============================================================================

def test_correlate_two_columns(mcp_call_tool, numeric_types_fixture):
    """Smoke: correlate calculates correlation between 2 columns."""
    print(f"\n📊 Testing correlate with 2 columns...")
    
    if len(numeric_types_fixture.columns) < 2:
        print(f"  ⚠️  Need at least 2 columns, skipping")
        return
    
    columns = numeric_types_fixture.columns[:2]
    
    result = mcp_call_tool("correlate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "columns": columns,
        "method": "pearson"
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from CorrelateResponse
    assert "correlation_matrix" in result, "Missing 'correlation_matrix'"
    assert "method" in result, "Missing 'method'"
    assert "columns" in result, "Missing 'columns'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify correlation matrix structure
    matrix = result["correlation_matrix"]
    assert isinstance(matrix, dict), "correlation_matrix should be dict"
    
    # Matrix should have entries for both columns
    for col in columns:
        assert col in matrix, f"Matrix missing column '{col}'"
        assert isinstance(matrix[col], dict), f"Matrix['{col}'] should be dict"
        
        # Each column should have correlation with all columns
        for col2 in columns:
            assert col2 in matrix[col], f"Matrix['{col}'] missing '{col2}'"
            corr_value = matrix[col][col2]
            assert isinstance(corr_value, (int, float)), f"Correlation should be numeric"
            assert -1 <= corr_value <= 1, f"Correlation should be between -1 and 1, got {corr_value}"
        
        # Diagonal should be 1 (correlation with self)
        assert abs(matrix[col][col] - 1.0) < 0.01, f"Correlation of '{col}' with itself should be 1"
    
    # Verify method
    assert result["method"] == "pearson"
    assert result["columns"] == columns
    
    print(f"  ✅ Correlation matrix calculated for {len(columns)} columns")


def test_correlate_multiple_columns(mcp_call_tool, numeric_types_fixture):
    """Smoke: correlate works with 3+ columns."""
    print(f"\n📊 Testing correlate with multiple columns...")
    
    if len(numeric_types_fixture.columns) < 3:
        print(f"  ⚠️  Need at least 3 columns, skipping")
        return
    
    columns = numeric_types_fixture.columns[:3]
    
    result = mcp_call_tool("correlate", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "columns": columns,
        "method": "spearman"
    })
    
    # Verify matrix is complete (3x3)
    matrix = result["correlation_matrix"]
    assert len(matrix) == 3, "Matrix should have 3 rows"
    for col in columns:
        assert len(matrix[col]) == 3, f"Matrix['{col}'] should have 3 columns"
    
    assert result["method"] == "spearman"
    
    print(f"  ✅ 3x3 correlation matrix calculated")


# ============================================================================
# STATISTICS TESTS - DETECT_OUTLIERS
# ============================================================================

def test_detect_outliers_iqr(mcp_call_tool, numeric_types_fixture):
    """Smoke: detect_outliers with IQR method."""
    print(f"\n📊 Testing detect_outliers (IQR method)...")
    
    column = numeric_types_fixture.columns[0]
    
    result = mcp_call_tool("detect_outliers", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "column": column,
        "method": "iqr",
        "threshold": 1.5
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from DetectOutliersResponse
    assert "outliers" in result, "Missing 'outliers'"
    assert "outlier_count" in result, "Missing 'outlier_count'"
    assert "method" in result, "Missing 'method'"
    assert "threshold" in result, "Missing 'threshold'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify outliers
    assert isinstance(result["outliers"], list), "outliers should be list"
    assert result["outlier_count"] == len(result["outliers"]), "outlier_count should match outliers length"
    assert result["outlier_count"] >= 0, "outlier_count should be non-negative"
    
    # Verify method and threshold
    assert result["method"] == "iqr"
    assert result["threshold"] == 1.5
    
    print(f"  ✅ Found {result['outlier_count']} outliers using IQR method")


def test_detect_outliers_zscore(mcp_call_tool, numeric_types_fixture):
    """Smoke: detect_outliers with Z-score method."""
    print(f"\n📊 Testing detect_outliers (Z-score method)...")
    
    column = numeric_types_fixture.columns[0]
    
    result = mcp_call_tool("detect_outliers", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "column": column,
        "method": "zscore",
        "threshold": 3.0
    })
    
    assert result["method"] == "zscore"
    assert result["threshold"] == 3.0
    assert result["outlier_count"] >= 0
    
    print(f"  ✅ Found {result['outlier_count']} outliers using Z-score method")


# ============================================================================
# VALIDATION TESTS - FIND_DUPLICATES
# ============================================================================

def test_find_duplicates_single_column(mcp_call_tool, with_duplicates_fixture):
    """Smoke: find_duplicates detects duplicates in single column."""
    print(f"\n🔍 Testing find_duplicates (single column)...")
    
    column = with_duplicates_fixture.columns[0]
    
    result = mcp_call_tool("find_duplicates", {
        "file_path": str(with_duplicates_fixture.path_str),
        "sheet_name": with_duplicates_fixture.sheet_name,
        "columns": [column]
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from FindDuplicatesResponse
    assert "duplicates" in result, "Missing 'duplicates'"
    assert "duplicate_count" in result, "Missing 'duplicate_count'"
    assert "columns_checked" in result, "Missing 'columns_checked'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify duplicates
    assert isinstance(result["duplicates"], list), "duplicates should be list"
    assert result["duplicate_count"] == len(result["duplicates"]), "duplicate_count should match duplicates length"
    assert result["duplicate_count"] >= 0, "duplicate_count should be non-negative"
    
    # Verify columns_checked
    assert result["columns_checked"] == [column]
    
    # If duplicates found, verify structure
    for i, dup in enumerate(result["duplicates"]):
        assert isinstance(dup, dict), f"Duplicate {i} should be dict"
        assert column in dup, f"Duplicate {i} missing column '{column}'"
    
    print(f"  ✅ Found {result['duplicate_count']} duplicate rows")


def test_find_duplicates_multiple_columns(mcp_call_tool, with_duplicates_fixture):
    """Smoke: find_duplicates detects duplicates across multiple columns."""
    print(f"\n🔍 Testing find_duplicates (multiple columns)...")
    
    if len(with_duplicates_fixture.columns) < 2:
        print(f"  ⚠️  Need at least 2 columns, skipping")
        return
    
    columns = with_duplicates_fixture.columns[:2]
    
    result = mcp_call_tool("find_duplicates", {
        "file_path": str(with_duplicates_fixture.path_str),
        "sheet_name": with_duplicates_fixture.sheet_name,
        "columns": columns
    })
    
    assert result["columns_checked"] == columns
    assert result["duplicate_count"] >= 0
    
    print(f"  ✅ Found {result['duplicate_count']} duplicate combinations")


# ============================================================================
# VALIDATION TESTS - FIND_NULLS
# ============================================================================

def test_find_nulls_basic(mcp_call_tool, with_nulls_fixture):
    """Smoke: find_nulls detects null values in columns."""
    print(f"\n🔍 Testing find_nulls...")
    
    columns = with_nulls_fixture.columns[:2] if len(with_nulls_fixture.columns) >= 2 else with_nulls_fixture.columns
    
    result = mcp_call_tool("find_nulls", {
        "file_path": str(with_nulls_fixture.path_str),
        "sheet_name": with_nulls_fixture.sheet_name,
        "columns": columns
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from FindNullsResponse
    assert "null_info" in result, "Missing 'null_info'"
    assert "total_nulls" in result, "Missing 'total_nulls'"
    assert "columns_checked" in result, "Missing 'columns_checked'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify null_info structure
    null_info = result["null_info"]
    assert isinstance(null_info, dict), "null_info should be dict"
    
    # Should have info for each checked column
    for col in columns:
        assert col in null_info, f"null_info missing column '{col}'"
        col_info = null_info[col]
        
        assert "count" in col_info, f"null_info['{col}'] missing 'count'"
        assert "percentage" in col_info, f"null_info['{col}'] missing 'percentage'"
        assert isinstance(col_info["count"], int), "null count should be int"
        assert col_info["count"] >= 0, "null count should be non-negative"
        assert 0 <= col_info["percentage"] <= 100, "percentage should be 0-100"
    
    # Verify total_nulls
    assert isinstance(result["total_nulls"], int), "total_nulls should be int"
    assert result["total_nulls"] >= 0, "total_nulls should be non-negative"
    
    # Verify columns_checked
    assert result["columns_checked"] == columns
    
    print(f"  ✅ Total nulls: {result['total_nulls']}")
    for col in columns:
        print(f"    {col}: {null_info[col]['count']} nulls ({null_info[col]['percentage']:.1f}%)")


# ============================================================================
# MULTI-SHEET TESTS - SEARCH_ACROSS_SHEETS
# ============================================================================

def test_search_across_sheets_basic(mcp_call_tool, multi_sheet_fixture):
    """Smoke: search_across_sheets finds value across multiple sheets."""
    print(f"\n🔍 Testing search_across_sheets...")
    
    # Search for a value that might exist
    result = mcp_call_tool("search_across_sheets", {
        "file_path": str(multi_sheet_fixture.path_str),
        "column_name": "Name",
        "value": "Test"
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from SearchAcrossSheetsResponse
    assert "matches" in result, "Missing 'matches'"
    assert "total_matches" in result, "Missing 'total_matches'"
    assert "column_name" in result, "Missing 'column_name'"
    assert "value" in result, "Missing 'value'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify matches structure
    assert isinstance(result["matches"], list), "matches should be list"
    assert result["total_matches"] >= 0, "total_matches should be non-negative"
    
    # Verify metadata
    assert result["column_name"] == "Name"
    assert result["value"] == "Test"
    
    # If matches found, verify structure
    for i, match in enumerate(result["matches"]):
        assert isinstance(match, dict), f"Match {i} should be dict"
        assert "sheet" in match, f"Match {i} missing 'sheet'"
        assert "count" in match, f"Match {i} missing 'count'"
        assert isinstance(match["count"], int), f"Match {i} count should be int"
        assert match["count"] > 0, f"Match {i} count should be positive"
    
    print(f"  ✅ Total matches: {result['total_matches']} across {len(result['matches'])} sheets")


# ============================================================================
# MULTI-SHEET TESTS - COMPARE_SHEETS
# ============================================================================

def test_compare_sheets_basic(mcp_call_tool, multi_sheet_fixture):
    """Smoke: compare_sheets finds differences between sheets."""
    print(f"\n🔍 Testing compare_sheets...")
    
    # Compare first two sheets
    sheets = multi_sheet_fixture.expected["sheet_names"]
    if len(sheets) < 2:
        print(f"  ⚠️  Need at least 2 sheets, skipping")
        return
    
    sheet1 = sheets[0]
    sheet2 = sheets[1]
    
    # Try to compare with a common column (might not exist)
    result = mcp_call_tool("compare_sheets", {
        "file_path": str(multi_sheet_fixture.path_str),
        "sheet1": sheet1,
        "sheet2": sheet2,
        "key_column": "ID",
        "compare_columns": ["Name"]
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from CompareSheetsResponse
    assert "differences" in result, "Missing 'differences'"
    assert "difference_count" in result, "Missing 'difference_count'"
    assert "key_column" in result, "Missing 'key_column'"
    assert "compare_columns" in result, "Missing 'compare_columns'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify differences
    assert isinstance(result["differences"], list), "differences should be list"
    assert result["difference_count"] == len(result["differences"]), "difference_count should match differences length"
    assert result["difference_count"] >= 0, "difference_count should be non-negative"
    
    # Verify metadata
    assert result["key_column"] == "ID"
    assert result["compare_columns"] == ["Name"]
    
    print(f"  ✅ Found {result['difference_count']} differences between '{sheet1}' and '{sheet2}'")
