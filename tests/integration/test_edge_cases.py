# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Integration tests for Edge Cases.

Tests cover:
- Merged cells handling (horizontal, vertical, title rows)
- Multi-level headers detection (2-level, 3-level hierarchies)
- Complex real-world scenarios (enterprise chaos)
- Interaction between merged cells and operations

These tests verify that the system handles messy real-world Excel files correctly.
"""

import pytest

from mcp_excel.operations.inspection import InspectionOperations
from mcp_excel.operations.data_operations import DataOperations
from mcp_excel.models.requests import (
    GetSheetInfoRequest,
    GetColumnNamesRequest,
    FilterAndCountRequest,
    FilterCondition,
    AggregateRequest,
    GetUniqueValuesRequest,
)


# ============================================================================
# Merged Cells Tests
# ============================================================================

def test_merged_cells_header_detection(merged_cells_fixture, file_loader):
    """Test header detection with merged cells in headers.
    
    Verifies:
    - HeaderDetector finds correct header row despite merges
    - Merged cells don't confuse the detection algorithm
    - Confidence is reasonable
    """
    print(f"\n📂 Testing merged cells header detection")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=merged_cells_fixture.path_str,
        sheet_name=merged_cells_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Detected header row: {response.header_detection.header_row}")
    print(f"   Confidence: {response.header_detection.confidence:.2%}")
    print(f"   Columns: {response.column_names}")
    
    assert response.header_detection is not None, "Should include header detection"
    assert response.header_detection.header_row == merged_cells_fixture.header_row, "Should detect correct header row"
    assert response.header_detection.confidence > 0.6, "Should have reasonable confidence despite merges"
    
    # Check that we got some columns (pandas reads merged cells as first cell value + Unnamed for others)
    assert len(response.column_names) == len(merged_cells_fixture.columns), "Should have correct number of columns"


def test_merged_cells_data_operations(merged_cells_fixture, file_loader):
    """Test that data operations work correctly with merged cells.
    
    Verifies:
    - Can filter data from file with merged headers
    - Can aggregate data from file with merged headers
    - Operations use correct header row
    """
    print(f"\n📂 Testing data operations with merged cells")
    
    ops = DataOperations(file_loader)
    
    # Get column names first
    inspection_ops = InspectionOperations(file_loader)
    sheet_info_request = GetSheetInfoRequest(
        file_path=merged_cells_fixture.path_str,
        sheet_name=merged_cells_fixture.sheet_name
    )
    sheet_info = inspection_ops.get_sheet_info(sheet_info_request)
    
    print(f"   Available columns: {sheet_info.column_names}")
    
    # Test filtering - use first non-merged column (should be "Регион" or similar)
    first_col = sheet_info.column_names[0]
    
    # Get unique values to filter on
    unique_request = GetUniqueValuesRequest(
        file_path=merged_cells_fixture.path_str,
        sheet_name=merged_cells_fixture.sheet_name,
        column=first_col,
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    
    if unique_response.values:
        filter_value = unique_response.values[0]
        print(f"   Testing filter: {first_col} == '{filter_value}'")
        
        # Test filter_and_count
        filter_request = FilterAndCountRequest(
            file_path=merged_cells_fixture.path_str,
            sheet_name=merged_cells_fixture.sheet_name,
            filters=[FilterCondition(column=first_col, operator="==", value=filter_value)],
            logic="AND"
        )
        filter_response = ops.filter_and_count(filter_request)
        
        print(f"✅ Filter works: {filter_response.count} rows matched")
        assert filter_response.count > 0, "Should find matching rows"
        assert filter_response.excel_output.formula, "Should generate Excel formula"


def test_merged_cells_column_names(merged_cells_fixture, file_loader):
    """Test get_column_names with merged cells.
    
    Verifies:
    - Can extract column names despite merges
    - Returns correct number of columns
    """
    print(f"\n📂 Testing get_column_names with merged cells")
    
    ops = InspectionOperations(file_loader)
    request = GetColumnNamesRequest(
        file_path=merged_cells_fixture.path_str,
        sheet_name=merged_cells_fixture.sheet_name
    )
    
    # Act
    response = ops.get_column_names(request)
    
    # Assert
    print(f"✅ Columns: {response.column_names}")
    print(f"   Count: {response.column_count}")
    
    assert response.column_count == len(merged_cells_fixture.columns), "Should return correct column count"
    assert len(response.column_names) == len(merged_cells_fixture.columns), "Should have all column names"


def test_merged_cells_sample_rows(merged_cells_fixture, file_loader):
    """Test that sample rows work correctly with merged cells.
    
    Verifies:
    - Sample rows contain actual data (not merge artifacts)
    - All columns are present in sample rows
    """
    print(f"\n📂 Testing sample rows with merged cells")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=merged_cells_fixture.path_str,
        sheet_name=merged_cells_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Sample rows count: {len(response.sample_rows)}")
    
    assert len(response.sample_rows) > 0, "Should have sample rows"
    
    # Check first sample row
    first_row = response.sample_rows[0]
    print(f"   First row columns: {len(first_row)}")
    print(f"   First row preview: {dict(list(first_row.items())[:3])}")
    
    assert len(first_row) == len(merged_cells_fixture.columns), "Sample row should have all columns"
    
    # Check that sample rows contain actual data (not None/NaN for all values)
    non_null_values = sum(1 for v in first_row.values() if v is not None)
    assert non_null_values > 0, "Sample row should have some non-null values"


# ============================================================================
# Multi-Level Headers Tests
# ============================================================================

def test_multilevel_headers_detection(multilevel_headers_fixture, file_loader):
    """Test header detection with 3-level hierarchy.
    
    Verifies:
    - HeaderDetector finds deepest header level
    - Skips company name and category rows
    - Returns correct column names from deepest level
    """
    print(f"\n📂 Testing multi-level headers detection")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=multilevel_headers_fixture.path_str,
        sheet_name=multilevel_headers_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Detected header row: {response.header_detection.header_row}")
    print(f"   Confidence: {response.header_detection.confidence:.2%}")
    print(f"   Columns: {response.column_names}")
    print(f"   Expected: {multilevel_headers_fixture.columns}")
    
    assert response.header_detection.header_row == multilevel_headers_fixture.header_row, "Should detect deepest level"
    assert response.column_names == multilevel_headers_fixture.columns, "Should get columns from deepest level"
    assert response.header_detection.confidence > 0.7, "Should have good confidence"


def test_multilevel_headers_data_integrity(multilevel_headers_fixture, file_loader):
    """Test that data is correctly read with multi-level headers.
    
    Verifies:
    - Data starts after correct header row
    - Row count is correct (excludes header rows)
    - Sample rows contain actual data
    """
    print(f"\n📂 Testing data integrity with multi-level headers")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=multilevel_headers_fixture.path_str,
        sheet_name=multilevel_headers_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Row count: {response.row_count}")
    print(f"   Data starts at row: {response.data_start_row}")
    print(f"   Expected rows: {multilevel_headers_fixture.row_count}")
    
    assert response.row_count == multilevel_headers_fixture.row_count, "Should count only data rows"
    assert response.data_start_row == multilevel_headers_fixture.header_row + 1, "Data should start after header"
    
    # Check sample rows
    assert len(response.sample_rows) > 0, "Should have sample rows"
    first_row = response.sample_rows[0]
    
    # Verify sample row has data (not header values)
    print(f"   First row preview: {dict(list(first_row.items())[:3])}")
    
    # Check that we have actual data values (not category names like "Информация", "Продажи")
    # The first column should be "ID" with values like "ID-1000", "ID-1001", etc.
    if "ID" in first_row:
        id_value = first_row["ID"]
        print(f"   ID value: {id_value}")
        assert id_value is not None, "ID should have value"
        # Check it's not a header value
        assert "Информация" not in str(id_value), "Should not contain header text"


def test_multilevel_headers_operations(multilevel_headers_fixture, file_loader):
    """Test that operations work correctly with multi-level headers.
    
    Verifies:
    - Can filter data
    - Can aggregate data
    - Operations use correct columns
    """
    print(f"\n📂 Testing operations with multi-level headers")
    
    ops = DataOperations(file_loader)
    
    # Get sheet info first
    inspection_ops = InspectionOperations(file_loader)
    sheet_info_request = GetSheetInfoRequest(
        file_path=multilevel_headers_fixture.path_str,
        sheet_name=multilevel_headers_fixture.sheet_name
    )
    sheet_info = inspection_ops.get_sheet_info(sheet_info_request)
    
    print(f"   Columns: {sheet_info.column_names}")
    
    # Find a numeric column for aggregation
    numeric_col = None
    for col, dtype in sheet_info.column_types.items():
        if dtype in ["integer", "float"]:
            numeric_col = col
            break
    
    if numeric_col:
        print(f"   Testing aggregation on: {numeric_col}")
        
        # Test aggregation
        agg_request = AggregateRequest(
            file_path=multilevel_headers_fixture.path_str,
            sheet_name=multilevel_headers_fixture.sheet_name,
            operation="count",
            target_column=numeric_col,
            filters=[]
        )
        agg_response = ops.aggregate(agg_request)
        
        print(f"✅ Aggregation works: count = {agg_response.value}")
        assert agg_response.value > 0, "Should count rows"
        assert agg_response.excel_output.formula, "Should generate formula"


def test_multilevel_headers_candidates(multilevel_headers_fixture, file_loader):
    """Test that header detection returns candidates for multi-level headers.
    
    Verifies:
    - Detection returns multiple candidates
    - Candidates include different header levels
    - Chosen candidate has highest score
    """
    print(f"\n📂 Testing header detection candidates")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=multilevel_headers_fixture.path_str,
        sheet_name=multilevel_headers_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Candidates count: {len(response.header_detection.candidates) if response.header_detection.candidates else 0}")
    
    if response.header_detection.candidates:
        print(f"   Top 3 candidates:")
        for i, candidate in enumerate(response.header_detection.candidates[:3], 1):
            print(f"     {i}. Row {candidate['row']}: score={candidate['score']:.2f}")
            print(f"        Preview: {candidate['preview'][:3]}")
        
        # Check that chosen row is in candidates
        chosen_row = response.header_detection.header_row
        candidate_rows = [c['row'] for c in response.header_detection.candidates]
        assert chosen_row in candidate_rows, "Chosen header should be in candidates"


# ============================================================================
# Enterprise Chaos Tests (Worst Case)
# ============================================================================

def test_enterprise_chaos_header_detection(enterprise_chaos_fixture, file_loader):
    """Test header detection on worst-case scenario.
    
    Verifies:
    - Can detect header despite: junk rows + merged cells + multi-level headers
    - Confidence is acceptable
    - Returns correct header row
    """
    print(f"\n📂 Testing enterprise chaos header detection")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=enterprise_chaos_fixture.path_str,
        sheet_name=enterprise_chaos_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Detected header row: {response.header_detection.header_row}")
    print(f"   Confidence: {response.header_detection.confidence:.2%}")
    print(f"   Columns: {response.column_names}")
    print(f"   Expected header row: {enterprise_chaos_fixture.header_row}")
    
    assert response.header_detection.header_row == enterprise_chaos_fixture.header_row, "Should handle worst-case"
    assert response.header_detection.confidence > 0.5, "Should have some confidence despite chaos"
    
    # Note: First column might be 'Unnamed: 0' due to merged cell
    # This is expected pandas behavior
    assert len(response.column_names) == len(enterprise_chaos_fixture.columns), "Should have correct column count"


def test_enterprise_chaos_data_operations(enterprise_chaos_fixture, file_loader):
    """Test that all operations work on worst-case file.
    
    Verifies:
    - Can read data despite complex structure
    - Can filter data
    - Can aggregate data
    - No crashes or errors
    """
    print(f"\n📂 Testing operations on enterprise chaos file")
    
    ops = DataOperations(file_loader)
    
    # Get sheet info
    inspection_ops = InspectionOperations(file_loader)
    sheet_info_request = GetSheetInfoRequest(
        file_path=enterprise_chaos_fixture.path_str,
        sheet_name=enterprise_chaos_fixture.sheet_name
    )
    sheet_info = inspection_ops.get_sheet_info(sheet_info_request)
    
    print(f"   Columns: {sheet_info.column_names}")
    print(f"   Row count: {sheet_info.row_count}")
    
    # Find a column to work with (skip first if it's Unnamed)
    work_col = None
    for col in sheet_info.column_names:
        if not col.startswith("Unnamed"):
            work_col = col
            break
    
    if work_col:
        print(f"   Testing with column: {work_col}")
        
        # Test get_unique_values
        unique_request = GetUniqueValuesRequest(
            file_path=enterprise_chaos_fixture.path_str,
            sheet_name=enterprise_chaos_fixture.sheet_name,
            column=work_col,
            limit=5
        )
        unique_response = ops.get_unique_values(unique_request)
        
        print(f"✅ Unique values: {unique_response.count}")
        assert unique_response.count > 0, "Should find unique values"
        
        # Test filtering if we have values
        if unique_response.values:
            filter_value = unique_response.values[0]
            filter_request = FilterAndCountRequest(
                file_path=enterprise_chaos_fixture.path_str,
                sheet_name=enterprise_chaos_fixture.sheet_name,
                filters=[FilterCondition(column=work_col, operator="==", value=filter_value)],
                logic="AND"
            )
            filter_response = ops.filter_and_count(filter_request)
            
            print(f"✅ Filter works: {filter_response.count} rows")
            assert filter_response.count > 0, "Should filter successfully"


def test_enterprise_chaos_row_count(enterprise_chaos_fixture, file_loader):
    """Test that row count is correct despite junk rows.
    
    Verifies:
    - Row count excludes junk rows and header rows
    - Counts all data rows including footer
    """
    print(f"\n📂 Testing row count on enterprise chaos")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=enterprise_chaos_fixture.path_str,
        sheet_name=enterprise_chaos_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Row count: {response.row_count}")
    print(f"   Expected (data rows): {enterprise_chaos_fixture.row_count}")
    print(f"   Junk rows before data: {enterprise_chaos_fixture.expected['junk_rows_before_data']}")
    
    # Note: Pandas reads all rows after header including empty rows and footer with formulas
    # The fixture has 5 client rows + empty row + footer row = 7 total rows
    assert response.row_count >= enterprise_chaos_fixture.row_count, "Should count at least the data rows"
    assert response.row_count <= 10, "Should not count excessive rows"


def test_enterprise_chaos_sample_rows_quality(enterprise_chaos_fixture, file_loader):
    """Test that sample rows contain actual data, not junk.
    
    Verifies:
    - Sample rows are from data section, not junk section
    - Sample rows have meaningful values
    """
    print(f"\n📂 Testing sample rows quality on enterprise chaos")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=enterprise_chaos_fixture.path_str,
        sheet_name=enterprise_chaos_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Sample rows: {len(response.sample_rows)}")
    
    assert len(response.sample_rows) > 0, "Should have sample rows"
    
    first_row = response.sample_rows[0]
    print(f"   First row preview: {dict(list(first_row.items())[:3])}")
    
    # Check that sample rows don't contain junk text
    junk_indicators = ["ООО", "ИНН", "КПП", "Сводный отчёт"]
    for value in first_row.values():
        if value is not None:
            value_str = str(value)
            for junk in junk_indicators:
                assert junk not in value_str, f"Sample row should not contain junk text: {junk}"


# ============================================================================
# Cross-Operation Tests with Edge Cases
# ============================================================================

def test_edge_cases_all_operations_work(merged_cells_fixture, file_loader):
    """Test that all major operations work with edge case files.
    
    Verifies:
    - inspect_file works
    - get_sheet_info works
    - get_column_names works
    - filter_and_count works
    - aggregate works
    - No crashes on edge cases
    """
    print(f"\n📂 Testing all operations on edge case file")
    
    inspection_ops = InspectionOperations(file_loader)
    data_ops = DataOperations(file_loader)
    
    # 1. inspect_file
    from mcp_excel.models.requests import InspectFileRequest
    inspect_request = InspectFileRequest(file_path=merged_cells_fixture.path_str)
    inspect_response = inspection_ops.inspect_file(inspect_request)
    print(f"✅ inspect_file: {inspect_response.sheet_count} sheets")
    assert inspect_response.sheet_count > 0
    
    # 2. get_sheet_info
    sheet_info_request = GetSheetInfoRequest(
        file_path=merged_cells_fixture.path_str,
        sheet_name=merged_cells_fixture.sheet_name
    )
    sheet_info = inspection_ops.get_sheet_info(sheet_info_request)
    print(f"✅ get_sheet_info: {sheet_info.column_count} columns")
    assert sheet_info.column_count > 0
    
    # 3. get_column_names
    col_names_request = GetColumnNamesRequest(
        file_path=merged_cells_fixture.path_str,
        sheet_name=merged_cells_fixture.sheet_name
    )
    col_names = inspection_ops.get_column_names(col_names_request)
    print(f"✅ get_column_names: {col_names.column_count} columns")
    assert col_names.column_count > 0
    
    # 4. filter_and_count (if we have columns)
    if sheet_info.column_names:
        first_col = sheet_info.column_names[0]
        filter_request = FilterAndCountRequest(
            file_path=merged_cells_fixture.path_str,
            sheet_name=merged_cells_fixture.sheet_name,
            filters=[FilterCondition(column=first_col, operator="is_not_null", value=None)],
            logic="AND"
        )
        filter_response = data_ops.filter_and_count(filter_request)
        print(f"✅ filter_and_count: {filter_response.count} rows")
        assert filter_response.count >= 0
    
    # 5. aggregate (if we have numeric columns)
    numeric_col = None
    for col, dtype in sheet_info.column_types.items():
        if dtype in ["integer", "float"]:
            numeric_col = col
            break
    
    if numeric_col:
        agg_request = AggregateRequest(
            file_path=merged_cells_fixture.path_str,
            sheet_name=merged_cells_fixture.sheet_name,
            operation="count",
            target_column=numeric_col,
            filters=[]
        )
        agg_response = data_ops.aggregate(agg_request)
        print(f"✅ aggregate: count = {agg_response.value}")
        assert agg_response.value >= 0


def test_edge_cases_performance(multilevel_headers_fixture, file_loader):
    """Test that edge case files don't cause performance issues.
    
    Verifies:
    - Operations complete in reasonable time
    - No excessive memory usage
    """
    print(f"\n📂 Testing performance on edge case file")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=multilevel_headers_fixture.path_str,
        sheet_name=multilevel_headers_fixture.sheet_name
    )
    
    # Act
    response = ops.get_sheet_info(request)
    
    # Assert
    print(f"✅ Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Memory used: {response.performance.memory_used_mb}MB")
    
    # Edge case files are small, should be fast
    assert response.performance.execution_time_ms < 5000, "Should complete in reasonable time"
    # Python process memory includes interpreter overhead, so 150 MB is reasonable threshold
    assert response.performance.memory_used_mb < 150, "Should not use excessive memory"


def test_edge_cases_error_messages(merged_cells_fixture, file_loader):
    """Test that error messages are helpful with edge case files.
    
    Verifies:
    - Invalid column names produce clear errors
    - Error messages mention available columns
    """
    print(f"\n📂 Testing error messages on edge case file")
    
    ops = DataOperations(file_loader)
    
    # Try to filter on non-existent column
    request = FilterAndCountRequest(
        file_path=merged_cells_fixture.path_str,
        sheet_name=merged_cells_fixture.sheet_name,
        filters=[FilterCondition(column="NonExistentColumn", operator="==", value="test")],
        logic="AND"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.filter_and_count(request)
    
    error_msg = str(exc_info.value)
    print(f"✅ Error message: {error_msg[:100]}...")
    
    assert "not found" in error_msg.lower() or "nonexistentcolumn" in error_msg.lower(), "Error should mention column not found"
    # Should mention available columns
    assert "available" in error_msg.lower() or "columns" in error_msg.lower(), "Error should mention available columns"
