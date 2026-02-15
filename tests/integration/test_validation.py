# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Integration tests for Validation operations.

Tests cover:
- find_duplicates: Find duplicate rows based on specified columns
- find_nulls: Find null/empty values in specified columns

These are END-TO-END tests that verify the complete operation flow:
FileLoader -> HeaderDetector -> ValidationOperations -> Response
"""

import pytest

from mcp_excel.operations.validation import ValidationOperations
from mcp_excel.models.requests import (
    FindDuplicatesRequest,
    FindNullsRequest,
)


# ============================================================================
# find_nulls tests
# ============================================================================

def test_find_nulls_single_column(with_nulls_fixture, file_loader):
    """Test find_nulls with single column.
    
    Verifies:
    - Finds null values in specified column
    - Returns correct null count and percentage
    - Includes null indices (limited to 100)
    - Generates TSV output
    """
    print(f"\n📂 Testing find_nulls on single column")
    
    ops = ValidationOperations(file_loader)
    
    # Check column that has nulls (from fixture metadata)
    cols_with_nulls = with_nulls_fixture.expected["columns_with_nulls"]
    test_col = cols_with_nulls[0]
    print(f"   Checking column: '{test_col}'")
    
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=[test_col]
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Total nulls found: {response.total_nulls}")
    print(f"   Columns checked: {response.columns_checked}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.columns_checked == [test_col], "Should check specified column"
    assert test_col in response.null_info, "Should have info for checked column"
    
    # Check null info structure
    null_info = response.null_info[test_col]
    assert "null_count" in null_info, "Should have null_count"
    assert "null_percentage" in null_info, "Should have null_percentage"
    assert "total_rows" in null_info, "Should have total_rows"
    assert "null_indices" in null_info, "Should have null_indices"
    assert "truncated" in null_info, "Should have truncated flag"
    
    print(f"   Null count: {null_info['null_count']}")
    print(f"   Null percentage: {null_info['null_percentage']}%")
    print(f"   Total rows: {null_info['total_rows']}")
    
    # Check that null_count matches total_nulls for single column
    assert response.total_nulls == null_info["null_count"], "Total should match single column count"
    
    # Check null percentage calculation
    expected_percentage = (null_info["null_count"] / null_info["total_rows"] * 100)
    assert abs(null_info["null_percentage"] - expected_percentage) < 0.1, "Percentage should be accurate"
    
    # Check null indices
    assert isinstance(null_info["null_indices"], list), "Null indices should be list"
    assert len(null_info["null_indices"]) <= 100, "Should limit to 100 indices"
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert test_col in response.excel_output.tsv, "TSV should contain column name"


def test_find_nulls_multiple_columns(with_nulls_fixture, file_loader):
    """Test find_nulls with multiple columns.
    
    Verifies:
    - Checks all specified columns
    - Returns separate statistics for each column
    - Total nulls is sum across all columns
    """
    print(f"\n📂 Testing find_nulls with multiple columns")
    
    ops = ValidationOperations(file_loader)
    
    # Check multiple columns
    cols_with_nulls = with_nulls_fixture.expected["columns_with_nulls"]
    check_cols = cols_with_nulls[:3] if len(cols_with_nulls) >= 3 else cols_with_nulls
    print(f"   Checking columns: {check_cols}")
    
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=check_cols
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Total nulls found: {response.total_nulls}")
    
    assert response.columns_checked == check_cols, "Should check all specified columns"
    assert len(response.null_info) == len(check_cols), "Should have info for each column"
    
    # Check that total_nulls is sum of individual counts
    sum_nulls = sum(info["null_count"] for info in response.null_info.values())
    assert response.total_nulls == sum_nulls, "Total should be sum of individual counts"
    
    # Print statistics for each column
    for col in check_cols:
        info = response.null_info[col]
        print(f"   {col}: {info['null_count']} nulls ({info['null_percentage']}%)")


def test_find_nulls_all_columns(with_nulls_fixture, file_loader):
    """Test find_nulls checking all columns.
    
    Verifies:
    - Checks all columns in the sheet
    - Returns comprehensive null statistics
    """
    print(f"\n📂 Testing find_nulls with all columns")
    
    ops = ValidationOperations(file_loader)
    
    # Check all columns
    all_cols = with_nulls_fixture.columns
    print(f"   Checking all {len(all_cols)} columns")
    
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=all_cols
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Total nulls found: {response.total_nulls}")
    print(f"   Columns with nulls: {sum(1 for info in response.null_info.values() if info['null_count'] > 0)}")
    
    assert response.columns_checked == all_cols, "Should check all columns"
    assert len(response.null_info) == len(all_cols), "Should have info for all columns"


def test_find_nulls_no_nulls(simple_fixture, file_loader):
    """Test find_nulls when no nulls exist.
    
    Verifies:
    - Returns 0 null count for columns without nulls
    - Total nulls is 0
    - Null percentage is 0
    """
    print(f"\n📂 Testing find_nulls with no nulls")
    
    ops = ValidationOperations(file_loader)
    
    # Simple fixture has no nulls
    first_col = simple_fixture.columns[0]
    print(f"   Checking column: '{first_col}' (should have no nulls)")
    
    request = FindNullsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        columns=[first_col]
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Total nulls found: {response.total_nulls}")
    
    assert response.total_nulls == 0, "Should find no nulls"
    
    null_info = response.null_info[first_col]
    assert null_info["null_count"] == 0, "Null count should be 0"
    assert null_info["null_percentage"] == 0, "Null percentage should be 0"
    assert len(null_info["null_indices"]) == 0, "Null indices should be empty"


def test_find_nulls_invalid_column(with_nulls_fixture, file_loader):
    """Test find_nulls with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message lists available columns
    """
    print(f"\n📂 Testing find_nulls with invalid column")
    
    ops = ValidationOperations(file_loader)
    
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=["NonExistentColumn"]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.find_nulls(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower(), "Error should mention column not found"
    assert "NonExistentColumn" in error_msg, "Error should mention the invalid column"
    assert "Available columns" in error_msg, "Error should list available columns"


def test_find_nulls_tsv_format(with_nulls_fixture, file_loader):
    """Test that find_nulls generates proper TSV output.
    
    Verifies:
    - TSV contains summary table
    - TSV has headers (Column, Null Count, Percentage, Total Rows)
    - TSV uses tab separators
    - Can be pasted into Excel
    """
    print(f"\n📂 Testing find_nulls TSV output format")
    
    ops = ValidationOperations(file_loader)
    
    cols_with_nulls = with_nulls_fixture.expected["columns_with_nulls"]
    check_cols = cols_with_nulls[:2] if len(cols_with_nulls) >= 2 else cols_with_nulls
    
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=check_cols
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ TSV output generated")
    print(f"   Length: {len(response.excel_output.tsv)} chars")
    print(f"   Preview (first 200 chars): {response.excel_output.tsv[:200]}...")
    
    tsv = response.excel_output.tsv
    
    # Check structure
    assert "\t" in tsv, "TSV should use tab separators"
    assert "\n" in tsv, "TSV should have multiple lines"
    
    # Check headers
    assert "Column" in tsv, "TSV should have Column header"
    assert "Null Count" in tsv, "TSV should have Null Count header"
    assert "Percentage" in tsv, "TSV should have Percentage header"
    assert "Total Rows" in tsv, "TSV should have Total Rows header"
    
    # Check that checked columns are in TSV
    for col in check_cols:
        assert col in tsv, f"TSV should include column {col}"


def test_find_nulls_null_indices_limit(with_nulls_fixture, file_loader):
    """Test that find_nulls limits null indices to 100.
    
    Verifies:
    - Null indices are limited to first 100
    - Truncated flag is set correctly
    """
    print(f"\n📂 Testing find_nulls null indices limit")
    
    ops = ValidationOperations(file_loader)
    
    # Check all columns (might have many nulls)
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=with_nulls_fixture.columns
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Checking null indices limits")
    
    for col, info in response.null_info.items():
        indices_count = len(info["null_indices"])
        print(f"   {col}: {indices_count} indices (truncated: {info['truncated']})")
        
        # Should never exceed 100
        assert indices_count <= 100, f"Column {col} should have max 100 indices"
        
        # If truncated, should have exactly 100
        if info["truncated"]:
            assert indices_count == 100, f"Truncated column {col} should have exactly 100 indices"
            assert info["null_count"] > 100, f"Truncated column {col} should have more than 100 nulls"
        else:
            # If not truncated, indices count should match null count
            assert indices_count == info["null_count"], f"Non-truncated column {col} indices should match null count"


def test_find_nulls_percentage_calculation(with_nulls_fixture, file_loader):
    """Test that find_nulls calculates null percentage correctly.
    
    Verifies:
    - Percentage is calculated as (null_count / total_rows * 100)
    - Percentage is rounded to 2 decimal places
    - Percentage is between 0 and 100
    """
    print(f"\n📂 Testing find_nulls percentage calculation")
    
    ops = ValidationOperations(file_loader)
    
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=with_nulls_fixture.columns
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Checking percentage calculations")
    
    for col, info in response.null_info.items():
        null_count = info["null_count"]
        total_rows = info["total_rows"]
        percentage = info["null_percentage"]
        
        # Calculate expected percentage
        expected = (null_count / total_rows * 100) if total_rows > 0 else 0
        
        print(f"   {col}: {null_count}/{total_rows} = {percentage}% (expected: {expected:.2f}%)")
        
        # Check accuracy (within 0.1% due to rounding)
        assert abs(percentage - expected) < 0.1, f"Percentage for {col} should be accurate"
        
        # Check range
        assert 0 <= percentage <= 100, f"Percentage for {col} should be between 0 and 100"


def test_find_nulls_wide_table(wide_table_fixture, file_loader):
    """Test find_nulls on wide table (50 columns).
    
    Verifies:
    - Handles many columns correctly
    - Performance is acceptable
    """
    print(f"\n📂 Testing find_nulls on wide table")
    
    ops = ValidationOperations(file_loader)
    
    # Check all 50 columns
    request = FindNullsRequest(
        file_path=wide_table_fixture.path_str,
        sheet_name=wide_table_fixture.sheet_name,
        columns=wide_table_fixture.columns
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Total nulls found: {response.total_nulls}")
    print(f"   Columns checked: {len(response.columns_checked)}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert len(response.columns_checked) == 50, "Should check all 50 columns"
    assert len(response.null_info) == 50, "Should have info for all 50 columns"
    assert response.performance.execution_time_ms < 5000, "Should complete in reasonable time"


def test_find_nulls_single_column_table(single_column_fixture, file_loader):
    """Test find_nulls on minimal table (single column).
    
    Verifies:
    - Handles edge case of single column
    """
    print(f"\n📂 Testing find_nulls on single column table")
    
    ops = ValidationOperations(file_loader)
    
    request = FindNullsRequest(
        file_path=single_column_fixture.path_str,
        sheet_name=single_column_fixture.sheet_name,
        columns=single_column_fixture.columns
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Total nulls found: {response.total_nulls}")
    
    assert len(response.null_info) == 1, "Should check single column"


def test_find_nulls_mixed_languages(mixed_languages_fixture, file_loader):
    """Test find_nulls with unicode data.
    
    Verifies:
    - Handles unicode correctly
    - No encoding errors
    """
    print(f"\n📂 Testing find_nulls with mixed languages")
    
    ops = ValidationOperations(file_loader)
    
    request = FindNullsRequest(
        file_path=mixed_languages_fixture.path_str,
        sheet_name=mixed_languages_fixture.sheet_name,
        columns=[mixed_languages_fixture.columns[0]]
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Total nulls found: {response.total_nulls}")
    print(f"   (Unicode handling verified - no encoding errors)")
    
    # Just verify no errors occurred
    assert response.columns_checked == [mixed_languages_fixture.columns[0]]


def test_find_nulls_performance_metrics(with_nulls_fixture, file_loader):
    """Test that find_nulls includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    - Rows processed count is correct
    """
    print(f"\n📂 Testing find_nulls performance metrics")
    
    ops = ValidationOperations(file_loader)
    
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=[with_nulls_fixture.columns[0]]
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Performance metrics:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Rows processed: {response.performance.rows_processed}")
    print(f"   Cache hit: {response.performance.cache_hit}")
    print(f"   Memory used: {response.performance.memory_used_mb}MB")
    
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.rows_processed == with_nulls_fixture.row_count, "Should process all rows"
    assert response.performance.cache_hit in [True, False], "Should report cache status"
    assert response.performance.memory_used_mb >= 0, "Should report memory usage"


def test_find_nulls_metadata(with_nulls_fixture, file_loader):
    """Test that find_nulls includes correct metadata.
    
    Verifies:
    - Metadata includes sheet name
    - Metadata includes row/column counts
    - Metadata includes file format
    """
    print(f"\n📂 Testing find_nulls metadata")
    
    ops = ValidationOperations(file_loader)
    
    request = FindNullsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        columns=[with_nulls_fixture.columns[0]]
    )
    
    # Act
    response = ops.find_nulls(request)
    
    # Assert
    print(f"✅ Metadata:")
    print(f"   Sheet: {response.metadata.sheet_name}")
    print(f"   Rows: {response.metadata.rows_total}")
    print(f"   Columns: {response.metadata.columns_total}")
    print(f"   Format: {response.metadata.file_format}")
    
    assert response.metadata.sheet_name == with_nulls_fixture.sheet_name, "Should have correct sheet name"
    assert response.metadata.rows_total == with_nulls_fixture.row_count, "Should have correct row count"
    assert response.metadata.columns_total == len(with_nulls_fixture.columns), "Should have correct column count"
    assert response.metadata.file_format == with_nulls_fixture.format, "Should have correct format"


# ============================================================================
# find_duplicates tests
# ============================================================================

def test_find_duplicates_single_column(with_duplicates_fixture, file_loader):
    """Test find_duplicates with single column.
    
    Verifies:
    - Finds all duplicate rows (including first occurrence)
    - Returns correct duplicate count
    - Includes row indices
    - Generates TSV output
    """
    print(f"\n📂 Testing find_duplicates on single column")
    
    ops = ValidationOperations(file_loader)
    
    # Check first column for duplicates
    first_col = with_duplicates_fixture.columns[0]
    print(f"   Checking column: '{first_col}'")
    
    request = FindDuplicatesRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        columns=[first_col]
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Duplicates found: {response.duplicate_count}")
    print(f"   Columns checked: {response.columns_checked}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.columns_checked == [first_col], "Should check specified column"
    assert response.duplicate_count > 0, "Should find duplicates in test fixture"
    assert len(response.duplicates) == response.duplicate_count, "Count should match list length"
    
    # Check that duplicates have row indices
    for dup in response.duplicates:
        assert "_row_index" in dup, "Each duplicate should have row index"
        assert isinstance(dup["_row_index"], int), "Row index should be integer"
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    assert first_col in response.excel_output.tsv, "TSV should contain column name"
    
    # Check metadata
    assert response.metadata.sheet_name == with_duplicates_fixture.sheet_name
    assert response.metadata.rows_total == with_duplicates_fixture.row_count
    
    # Check performance metrics
    assert response.performance.execution_time_ms > 0
    assert response.performance.rows_processed == with_duplicates_fixture.row_count


def test_find_duplicates_multiple_columns(with_duplicates_fixture, file_loader):
    """Test find_duplicates with multiple columns.
    
    Verifies:
    - Finds duplicates based on combination of columns
    - Only rows with ALL columns matching are considered duplicates
    """
    print(f"\n📂 Testing find_duplicates with multiple columns")
    
    ops = ValidationOperations(file_loader)
    
    # Check first 2 columns for duplicates
    check_cols = with_duplicates_fixture.columns[:2]
    print(f"   Checking columns: {check_cols}")
    
    request = FindDuplicatesRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        columns=check_cols
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Duplicates found: {response.duplicate_count}")
    print(f"   Sample duplicate (first 3 fields): {dict(list(response.duplicates[0].items())[:3]) if response.duplicates else 'None'}")
    
    assert response.columns_checked == check_cols, "Should check specified columns"
    assert response.duplicate_count > 0, "Should find duplicates"
    
    # Verify that duplicates have matching values in checked columns
    if len(response.duplicates) >= 2:
        # Group duplicates by checked column values
        groups = {}
        for dup in response.duplicates:
            key = tuple(dup.get(col) for col in check_cols)
            if key not in groups:
                groups[key] = []
            groups[key].append(dup)
        
        # Each group should have at least 2 rows (otherwise not duplicates)
        for key, group in groups.items():
            assert len(group) >= 2, f"Duplicate group {key} should have at least 2 rows"
            print(f"   Duplicate group {key}: {len(group)} rows")


def test_find_duplicates_all_columns(with_duplicates_fixture, file_loader):
    """Test find_duplicates checking all columns.
    
    Verifies:
    - Finds exact duplicate rows (all columns match)
    - More strict than single/multi-column check
    """
    print(f"\n📂 Testing find_duplicates with all columns")
    
    ops = ValidationOperations(file_loader)
    
    # Check all columns
    all_cols = with_duplicates_fixture.columns
    print(f"   Checking all {len(all_cols)} columns")
    
    request = FindDuplicatesRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        columns=all_cols
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Exact duplicates found: {response.duplicate_count}")
    
    assert response.columns_checked == all_cols, "Should check all columns"
    # Note: might be 0 if no exact duplicates exist
    print(f"   (Exact duplicates require ALL columns to match)")


def test_find_duplicates_no_duplicates(simple_fixture, file_loader):
    """Test find_duplicates when no duplicates exist.
    
    Verifies:
    - Returns empty list when no duplicates
    - Count is 0
    - TSV indicates no duplicates
    """
    print(f"\n📂 Testing find_duplicates with no duplicates")
    
    ops = ValidationOperations(file_loader)
    
    # Simple fixture has unique values
    first_col = simple_fixture.columns[0]
    print(f"   Checking column: '{first_col}' (should have no duplicates)")
    
    request = FindDuplicatesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        columns=[first_col]
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Duplicates found: {response.duplicate_count}")
    
    assert response.duplicate_count == 0, "Should find no duplicates"
    assert len(response.duplicates) == 0, "Duplicates list should be empty"
    assert "No duplicates found" in response.excel_output.tsv, "TSV should indicate no duplicates"


def test_find_duplicates_invalid_column(with_duplicates_fixture, file_loader):
    """Test find_duplicates with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message lists available columns
    """
    print(f"\n📂 Testing find_duplicates with invalid column")
    
    ops = ValidationOperations(file_loader)
    
    request = FindDuplicatesRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        columns=["NonExistentColumn"]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.find_duplicates(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower(), "Error should mention column not found"
    assert "NonExistentColumn" in error_msg, "Error should mention the invalid column"
    assert "Available columns" in error_msg, "Error should list available columns"


def test_find_duplicates_tsv_format(with_duplicates_fixture, file_loader):
    """Test that find_duplicates generates proper TSV output.
    
    Verifies:
    - TSV contains headers
    - TSV contains row indices
    - TSV uses tab separators
    - Can be pasted into Excel
    """
    print(f"\n📂 Testing find_duplicates TSV output format")
    
    ops = ValidationOperations(file_loader)
    
    request = FindDuplicatesRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        columns=[with_duplicates_fixture.columns[0]]
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ TSV output generated")
    print(f"   Length: {len(response.excel_output.tsv)} chars")
    print(f"   Preview (first 200 chars): {response.excel_output.tsv[:200]}...")
    
    tsv = response.excel_output.tsv
    
    # Check structure
    assert "\t" in tsv, "TSV should use tab separators"
    assert "\n" in tsv, "TSV should have multiple lines"
    
    # Check headers
    assert "_row_index" in tsv, "TSV should include row index column"
    for col in with_duplicates_fixture.columns:
        assert col in tsv, f"TSV should include column {col}"
    
    # Check that first line is headers
    lines = tsv.split("\n")
    assert len(lines) > 1, "TSV should have header + data lines"
    header_line = lines[0]
    assert "_row_index" in header_line, "First line should be headers"


def test_find_duplicates_row_indices(with_duplicates_fixture, file_loader):
    """Test that find_duplicates returns correct row indices.
    
    Verifies:
    - Row indices are 0-based
    - Row indices correspond to original DataFrame positions
    - All duplicates have indices
    """
    print(f"\n📂 Testing find_duplicates row indices")
    
    ops = ValidationOperations(file_loader)
    
    request = FindDuplicatesRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        columns=[with_duplicates_fixture.columns[0]]
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Checking row indices for {response.duplicate_count} duplicates")
    
    indices = [dup["_row_index"] for dup in response.duplicates]
    print(f"   Row indices: {indices[:10]}..." if len(indices) > 10 else f"   Row indices: {indices}")
    
    # Check all have indices
    assert len(indices) == response.duplicate_count, "All duplicates should have indices"
    
    # Check indices are valid (0-based, within range)
    for idx in indices:
        assert isinstance(idx, int), "Index should be integer"
        assert idx >= 0, "Index should be non-negative"
        assert idx < with_duplicates_fixture.row_count, "Index should be within data range"
    
    # Check indices are sorted (duplicates should be in order)
    # Note: might not be sorted if duplicates are scattered
    print(f"   Indices are {'sorted' if indices == sorted(indices) else 'not sorted'}")


def test_find_duplicates_wide_table(wide_table_fixture, file_loader):
    """Test find_duplicates on wide table (50 columns).
    
    Verifies:
    - Handles many columns correctly
    - Performance is acceptable
    """
    print(f"\n📂 Testing find_duplicates on wide table")
    
    ops = ValidationOperations(file_loader)
    
    # Check first column (should have duplicates since values are "Значение_0_0", "Значение_1_0", etc.)
    request = FindDuplicatesRequest(
        file_path=wide_table_fixture.path_str,
        sheet_name=wide_table_fixture.sheet_name,
        columns=[wide_table_fixture.columns[0]]
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Duplicates found: {response.duplicate_count}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.performance.execution_time_ms < 5000, "Should complete in reasonable time"
    
    # Check that response includes all columns
    if response.duplicates:
        first_dup = response.duplicates[0]
        # Should have 50 columns + _row_index
        assert len(first_dup) == 51, "Should include all 50 columns + row index"


def test_find_duplicates_single_column_table(single_column_fixture, file_loader):
    """Test find_duplicates on minimal table (single column).
    
    Verifies:
    - Handles edge case of single column
    """
    print(f"\n📂 Testing find_duplicates on single column table")
    
    ops = ValidationOperations(file_loader)
    
    request = FindDuplicatesRequest(
        file_path=single_column_fixture.path_str,
        sheet_name=single_column_fixture.sheet_name,
        columns=single_column_fixture.columns
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Duplicates found: {response.duplicate_count}")
    
    # Single column table with unique values should have no duplicates
    assert response.duplicate_count == 0, "Single column with unique values should have no duplicates"


def test_find_duplicates_mixed_languages(mixed_languages_fixture, file_loader):
    """Test find_duplicates with unicode data.
    
    Verifies:
    - Handles unicode correctly in duplicate detection
    - No encoding errors
    """
    print(f"\n📂 Testing find_duplicates with mixed languages")
    
    ops = ValidationOperations(file_loader)
    
    request = FindDuplicatesRequest(
        file_path=mixed_languages_fixture.path_str,
        sheet_name=mixed_languages_fixture.sheet_name,
        columns=[mixed_languages_fixture.columns[0]]
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Duplicates found: {response.duplicate_count}")
    print(f"   (Unicode handling verified - no encoding errors)")
    
    # Just verify no errors occurred
    assert response.columns_checked == [mixed_languages_fixture.columns[0]]


def test_find_duplicates_performance_metrics(with_duplicates_fixture, file_loader):
    """Test that find_duplicates includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    - Rows processed count is correct
    """
    print(f"\n📂 Testing find_duplicates performance metrics")
    
    ops = ValidationOperations(file_loader)
    
    request = FindDuplicatesRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        columns=[with_duplicates_fixture.columns[0]]
    )
    
    # Act
    response = ops.find_duplicates(request)
    
    # Assert
    print(f"✅ Performance metrics:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Rows processed: {response.performance.rows_processed}")
    print(f"   Cache hit: {response.performance.cache_hit}")
    print(f"   Memory used: {response.performance.memory_used_mb}MB")
    
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.rows_processed == with_duplicates_fixture.row_count, "Should process all rows"
    assert response.performance.cache_hit in [True, False], "Should report cache status"
    assert response.performance.memory_used_mb >= 0, "Should report memory usage"
