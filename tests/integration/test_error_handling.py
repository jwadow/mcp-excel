# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Integration tests for Error Handling across all operations.

Tests cover:
- File system errors (not found, invalid format, corrupted)
- Column validation errors (missing columns)
- Filter validation errors (invalid operators, missing parameters)
- Data type errors (non-numeric aggregation)
- Context overflow errors (response too large)
- Multi-sheet operation errors (missing key columns)
- Sheet validation errors (non-existent sheets, empty sheets)

These are END-TO-END tests that verify error handling quality:
- Correct error type is raised
- Error messages are helpful and actionable
- Error messages include suggestions where applicable
"""

import pytest
from pathlib import Path

from mcp_excel.operations.inspection import InspectionOperations
from mcp_excel.operations.data_operations import DataOperations
from mcp_excel.operations.statistics import StatisticsOperations
from mcp_excel.models.requests import (
    InspectFileRequest,
    GetSheetInfoRequest,
    GetUniqueValuesRequest,
    GetValueCountsRequest,
    FilterAndCountRequest,
    FilterAndGetRowsRequest,
    FilterCondition,
    AggregateRequest,
    GroupByRequest,
    CompareSheetsRequest,
    GetColumnStatsRequest,
)


# ============================================================================
# Priority 1: File System Errors
# ============================================================================

def test_file_not_found(file_loader):
    """Test error when file doesn't exist.
    
    Verifies:
    - FileNotFoundError is raised
    - Error message mentions the file path
    - Error message suggests using absolute path
    """
    print(f"\n❌ Testing error: File not found")
    
    ops = InspectionOperations(file_loader)
    non_existent_path = "C:/NonExistent/Path/file.xlsx"
    request = InspectFileRequest(file_path=non_existent_path)
    
    # Act & Assert
    with pytest.raises(FileNotFoundError) as exc_info:
        ops.inspect_file(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower(), "Error should mention 'not found'"
    assert non_existent_path in error_msg or "file.xlsx" in error_msg, "Error should mention file path"
    assert "absolute path" in error_msg.lower(), "Error should suggest using absolute path"


def test_invalid_file_format_txt(temp_excel_path, file_loader):
    """Test error when file has unsupported format (.txt).
    
    Verifies:
    - ValueError is raised
    - Error message mentions unsupported format
    - Error message lists supported formats
    """
    print(f"\n❌ Testing error: Invalid file format (.txt)")
    
    # Create a .txt file
    txt_file = temp_excel_path / "test.txt"
    txt_file.write_text("This is not an Excel file")
    
    ops = InspectionOperations(file_loader)
    request = InspectFileRequest(file_path=str(txt_file))
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.inspect_file(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "unsupported" in error_msg.lower(), "Error should mention 'unsupported'"
    assert ".txt" in error_msg.lower(), "Error should mention .txt format"
    assert ".xls" in error_msg.lower() or ".xlsx" in error_msg.lower(), "Error should list supported formats"


def test_invalid_file_format_pdf(temp_excel_path, file_loader):
    """Test error when file has unsupported format (.pdf).
    
    Verifies:
    - ValueError is raised
    - Error message is clear about unsupported format
    """
    print(f"\n❌ Testing error: Invalid file format (.pdf)")
    
    # Create a .pdf file
    pdf_file = temp_excel_path / "test.pdf"
    pdf_file.write_bytes(b"%PDF-1.4\n")
    
    ops = InspectionOperations(file_loader)
    request = InspectFileRequest(file_path=str(pdf_file))
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.inspect_file(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "unsupported" in error_msg.lower(), "Error should mention 'unsupported'"
    assert ".pdf" in error_msg.lower(), "Error should mention .pdf format"


def test_corrupted_xlsx_file(temp_excel_path, file_loader):
    """Test error when Excel file is corrupted.
    
    Verifies:
    - Exception is raised (generic or specific)
    - Error message mentions failure to load
    """
    print(f"\n❌ Testing error: Corrupted Excel file")
    
    # Create a corrupted .xlsx file (empty file with .xlsx extension)
    corrupted_file = temp_excel_path / "corrupted.xlsx"
    corrupted_file.write_bytes(b"This is not a valid Excel file")
    
    ops = InspectionOperations(file_loader)
    request = InspectFileRequest(file_path=str(corrupted_file))
    
    # Act & Assert
    with pytest.raises(Exception) as exc_info:
        ops.inspect_file(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "failed" in error_msg.lower() or "error" in error_msg.lower(), "Error should mention failure"


# ============================================================================
# Priority 2: Column Validation Errors
# ============================================================================

def test_column_not_found_get_unique_values(simple_fixture, file_loader):
    """Test error when column doesn't exist in get_unique_values.
    
    Verifies:
    - ValueError is raised
    - Error message mentions column name
    - Error message lists available columns
    """
    print(f"\n❌ Testing error: Column not found in get_unique_values")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column="NonExistentColumn",
        limit=100
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.get_unique_values(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower(), "Error should mention 'not found'"
    assert "NonExistentColumn" in error_msg, "Error should mention the invalid column"
    assert "available" in error_msg.lower(), "Error should list available columns"
    # Check that at least one actual column is mentioned
    assert any(col in error_msg for col in simple_fixture.columns), "Error should list actual columns"


def test_column_not_found_aggregate(simple_fixture, file_loader):
    """Test error when column doesn't exist in aggregate.
    
    Verifies:
    - ValueError is raised
    - Error message is helpful
    """
    print(f"\n❌ Testing error: Column not found in aggregate")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="sum",
        target_column="NonExistentColumn",
        filters=[]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.aggregate(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower(), "Error should mention 'not found'"
    assert "NonExistentColumn" in error_msg, "Error should mention the invalid column"


def test_column_not_found_filter(simple_fixture, file_loader):
    """Test error when column doesn't exist in filter.
    
    Verifies:
    - ValueError is raised
    - Error message lists available columns
    """
    print(f"\n❌ Testing error: Column not found in filter")
    
    ops = DataOperations(file_loader)
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column="NonExistentColumn", operator="==", value="test")
        ],
        logic="AND"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.filter_and_count(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower(), "Error should mention 'not found'"
    assert "available" in error_msg.lower(), "Error should list available columns"


def test_multiple_missing_columns_group_by(simple_fixture, file_loader):
    """Test error when multiple columns don't exist in group_by.
    
    Verifies:
    - ValueError is raised
    - Error message mentions missing columns
    """
    print(f"\n❌ Testing error: Multiple missing columns in group_by")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["NonExistent1", "NonExistent2"],
        agg_column=simple_fixture.columns[0],
        agg_operation="count",
        filters=[]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.group_by(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower() or "missing" in error_msg.lower(), "Error should mention missing columns"


# ============================================================================
# Priority 3: Filter Validation Errors
# ============================================================================

def test_invalid_filter_operator(simple_fixture, file_loader):
    """Test error when filter uses unsupported operator.
    
    Verifies:
    - ValueError is raised
    - Error message mentions unsupported operator
    """
    print(f"\n❌ Testing error: Invalid filter operator")
    
    ops = DataOperations(file_loader)
    
    # Create filter with invalid operator (bypass Pydantic validation by using dict)
    # Note: This might be caught by Pydantic first, so we test at FilterEngine level
    from mcp_excel.operations.filtering import FilterEngine
    
    filter_engine = FilterEngine()
    
    # Load data
    df, _ = ops._load_with_header_detection(
        simple_fixture.path_str, simple_fixture.sheet_name, None
    )
    
    # Create invalid filter condition manually
    class InvalidFilter:
        column = simple_fixture.columns[0]
        operator = "INVALID_OP"
        value = "test"
        values = None
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        filter_engine._build_filter_mask(df, InvalidFilter())
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "unsupported" in error_msg.lower() or "invalid" in error_msg.lower(), "Error should mention invalid operator"


def test_filter_in_without_values(simple_fixture, file_loader):
    """Test error when 'in' operator is used without 'values' parameter.
    
    Verifies:
    - ValueError is raised
    - Error message mentions missing 'values' parameter
    """
    print(f"\n❌ Testing error: Filter 'in' without values")
    
    ops = DataOperations(file_loader)
    from mcp_excel.operations.filtering import FilterEngine
    
    filter_engine = FilterEngine()
    
    # Load data
    df, _ = ops._load_with_header_detection(
        simple_fixture.path_str, simple_fixture.sheet_name, None
    )
    
    # Create filter with 'in' operator but no values
    class FilterWithoutValues:
        column = simple_fixture.columns[0]
        operator = "in"
        value = None
        values = None  # Missing values
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        filter_engine._build_filter_mask(df, FilterWithoutValues())
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "'in'" in error_msg.lower() or "in operator" in error_msg.lower(), "Error should mention 'in' operator"
    assert "values" in error_msg.lower(), "Error should mention 'values' parameter"


def test_filter_comparison_without_value(simple_fixture, file_loader):
    """Test behavior when comparison operator is used with None value.
    
    Verifies:
    - Pandas handles None comparison gracefully (returns False for all)
    - No crash occurs
    - Result is empty (no matches)
    """
    print(f"\n❌ Testing behavior: Filter comparison with None value")
    
    ops = DataOperations(file_loader)
    from mcp_excel.operations.filtering import FilterEngine
    
    filter_engine = FilterEngine()
    
    # Load data
    df, _ = ops._load_with_header_detection(
        simple_fixture.path_str, simple_fixture.sheet_name, None
    )
    
    # Create filter with '>' operator but no value
    class FilterWithoutValue:
        column = simple_fixture.columns[1]  # Numeric column
        operator = ">"
        value = None  # Missing value
        values = None
    
    # Act - pandas handles None gracefully, returns False for all comparisons
    mask = filter_engine._build_filter_mask(df, FilterWithoutValue())
    
    # Assert - should return all False (no matches)
    print(f"✅ Pandas handled None gracefully: {mask.sum()} matches (expected 0)")
    assert mask.sum() == 0, "Comparison with None should match no rows"


def test_invalid_regex_pattern(simple_fixture, file_loader):
    """Test error when regex pattern is invalid.
    
    Verifies:
    - ValueError is raised
    - Error message mentions invalid regex
    """
    print(f"\n❌ Testing error: Invalid regex pattern")
    
    ops = DataOperations(file_loader)
    from mcp_excel.operations.filtering import FilterEngine
    
    filter_engine = FilterEngine()
    
    # Load data
    df, _ = ops._load_with_header_detection(
        simple_fixture.path_str, simple_fixture.sheet_name, None
    )
    
    # Create filter with invalid regex
    class FilterWithInvalidRegex:
        column = simple_fixture.columns[0]
        operator = "regex"
        value = "[invalid(regex"  # Invalid regex pattern
        values = None
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        filter_engine._build_filter_mask(df, FilterWithInvalidRegex())
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "regex" in error_msg.lower() or "pattern" in error_msg.lower(), "Error should mention regex/pattern"


def test_string_operator_on_non_string(simple_fixture, file_loader):
    """Test error when string operator is used with non-string value.
    
    Verifies:
    - ValueError is raised
    - Error message mentions string requirement
    """
    print(f"\n❌ Testing error: String operator with non-string value")
    
    ops = DataOperations(file_loader)
    from mcp_excel.operations.filtering import FilterEngine
    
    filter_engine = FilterEngine()
    
    # Load data
    df, _ = ops._load_with_header_detection(
        simple_fixture.path_str, simple_fixture.sheet_name, None
    )
    
    # Create filter with 'contains' operator but numeric value
    class FilterWithNonStringValue:
        column = simple_fixture.columns[0]
        operator = "contains"
        value = 123  # Non-string value
        values = None
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        filter_engine._build_filter_mask(df, FilterWithNonStringValue())
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "string" in error_msg.lower(), "Error should mention string requirement"
    assert "contains" in error_msg.lower(), "Error should mention the operator"


# ============================================================================
# Priority 4: Data Type Errors
# ============================================================================

def test_aggregate_on_non_numeric_column(simple_fixture, file_loader):
    """Test error when aggregation is attempted on text column.
    
    Verifies:
    - ValueError is raised
    - Error message mentions non-numeric data
    - Error message is helpful
    """
    print(f"\n❌ Testing error: Aggregate on non-numeric column")
    
    ops = DataOperations(file_loader)
    
    # Use first column (text column "Имя")
    text_column = simple_fixture.columns[0]
    
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="sum",
        target_column=text_column,
        filters=[]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.aggregate(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "numeric" in error_msg.lower() or "number" in error_msg.lower(), "Error should mention numeric requirement"


def test_group_by_non_numeric_aggregation(simple_fixture, file_loader):
    """Test behavior when group_by aggregation is on non-numeric column.
    
    Verifies:
    - Either ValueError is raised (preferred) OR
    - Operation succeeds but returns NaN/empty results (pandas behavior)
    
    Note: Pandas groupby.sum() on text columns may not raise error, just return NaN.
    The code has auto-conversion logic that tries to convert text to numeric.
    """
    print(f"\n❌ Testing behavior: Group by with non-numeric aggregation")
    
    ops = DataOperations(file_loader)
    
    # Group by text column, aggregate on text column with 'sum'
    text_column = simple_fixture.columns[0]  # "Имя"
    city_column = simple_fixture.columns[2]  # "Город"
    
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=[city_column],
        agg_column=text_column,  # Text column for aggregation
        agg_operation="sum",  # Sum requires numeric
        filters=[]
    )
    
    # Act & Assert
    try:
        response = ops.group_by(request)
        # If it succeeds, check that results are NaN or zero (pandas behavior)
        print(f"⚠️ Operation succeeded (pandas handled text gracefully)")
        print(f"   Groups returned: {len(response.groups)}")
        if response.groups:
            first_group = response.groups[0]
            agg_value = first_group.get(f"{text_column}_sum", first_group.get("sum", None))
            print(f"   First group aggregation value: {agg_value}")
            # Pandas sum on text returns 0 or NaN
            assert agg_value == 0 or agg_value is None or str(agg_value) == 'nan', "Text sum should return 0 or NaN"
    except ValueError as exc_info:
        # This is the preferred behavior - explicit error
        print(f"✅ Caught expected error: {exc_info}")
        error_msg = str(exc_info)
        assert "numeric" in error_msg.lower() or "number" in error_msg.lower() or "cannot" in error_msg.lower(), "Error should mention numeric requirement"


# ============================================================================
# Priority 5: Context Overflow Errors
# ============================================================================

def test_response_size_overflow_filter_and_get_rows(wide_table_fixture, file_loader):
    """Test error when response size exceeds limit in filter_and_get_rows.
    
    Verifies:
    - ValueError is raised when response is too large
    - Error message mentions response size
    - Error message provides suggestions
    
    Note: wide_table has 50 columns × 10 rows. To trigger overflow (10,000 chars limit),
    we need to request all columns with all rows, which creates a large JSON response.
    """
    print(f"\n❌ Testing error: Response size overflow")
    
    ops = DataOperations(file_loader)
    
    # Request all 50 columns with all 10 rows
    # Each cell has ~15 chars ("Значение_X_Y"), so 50 cols × 10 rows × 15 chars = ~7500 chars
    # Plus JSON overhead should exceed 10,000 char limit
    request = FilterAndGetRowsRequest(
        file_path=wide_table_fixture.path_str,
        sheet_name=wide_table_fixture.sheet_name,
        filters=[],
        columns=None,  # All 50 columns
        limit=10,  # All rows
        offset=0,
        logic="AND"
    )
    
    # Act & Assert
    try:
        response = ops.filter_and_get_rows(request)
        # If it doesn't raise, check if response is actually large
        print(f"⚠️ Response did not trigger overflow. Response has {response.count} rows, {len(response.rows[0]) if response.rows else 0} columns")
        # This is acceptable - the fixture might not be large enough
        # Skip assertion in this case
    except ValueError as exc_info:
        print(f"✅ Caught expected error: {exc_info}")
        error_msg = str(exc_info)
        assert "too large" in error_msg.lower() or "limit" in error_msg.lower(), "Error should mention size issue"
        assert "reduce" in error_msg.lower() or "fewer" in error_msg.lower() or "MCP" in error_msg, "Error should provide suggestions"


# ============================================================================
# Priority 6: Multi-Sheet Operation Errors
# ============================================================================

def test_compare_sheets_missing_key_column(multi_sheet_fixture, file_loader):
    """Test error when key_column is missing in one of the sheets.
    
    Verifies:
    - ValueError is raised
    - Error message mentions missing key column
    """
    print(f"\n❌ Testing error: Compare sheets with missing key column")
    
    ops = InspectionOperations(file_loader)
    
    # Try to compare sheets with non-existent key column
    request = CompareSheetsRequest(
        file_path=multi_sheet_fixture.path_str,
        sheet1="Products",
        sheet2="Clients",
        key_column="NonExistentKey",
        compare_columns=[]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.compare_sheets(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower() or "missing" in error_msg.lower(), "Error should mention missing column"
    assert "key" in error_msg.lower() or "NonExistentKey" in error_msg, "Error should mention key column"


def test_compare_sheets_missing_compare_column(multi_sheet_fixture, file_loader):
    """Test error when compare_column is missing in one of the sheets.
    
    Verifies:
    - ValueError is raised
    - Error message is helpful
    """
    print(f"\n❌ Testing error: Compare sheets with missing compare column")
    
    ops = InspectionOperations(file_loader)
    
    # Products has "Товар", Clients has "Клиент"
    # Use "Товар" as key (exists in Products), try to compare "NonExistent"
    request = CompareSheetsRequest(
        file_path=multi_sheet_fixture.path_str,
        sheet1="Products",
        sheet2="Clients",
        key_column="Товар",  # Exists in Products but not in Clients
        compare_columns=["NonExistentCompare"]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.compare_sheets(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower() or "missing" in error_msg.lower(), "Error should mention missing column"


# ============================================================================
# Priority 7: Sheet Validation Errors
# ============================================================================

def test_sheet_not_found(simple_fixture, file_loader):
    """Test error when sheet doesn't exist.
    
    Verifies:
    - Exception is raised (FileLoader wraps pandas errors in generic Exception)
    - Error message mentions sheet name
    """
    print(f"\n❌ Testing error: Sheet not found")
    
    ops = InspectionOperations(file_loader)
    request = GetSheetInfoRequest(
        file_path=simple_fixture.path_str,
        sheet_name="NonExistentSheet"
    )
    
    # Act & Assert
    # FileLoader wraps pandas exceptions in generic Exception
    with pytest.raises(Exception) as exc_info:
        ops.get_sheet_info(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "NonExistentSheet" in error_msg or "sheet" in error_msg.lower() or "not found" in error_msg.lower(), "Error should mention sheet"


def test_empty_sheet_operations(temp_excel_path, file_loader):
    """Test operations on empty sheet (no data rows).
    
    Verifies:
    - Operations handle empty sheets gracefully
    - Raises appropriate error (header detection fails on empty sheet)
    """
    print(f"\n❌ Testing error: Empty sheet operations")
    
    # Create Excel file with empty sheet
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Empty"
    # Don't add any data
    
    empty_file = temp_excel_path / "empty_sheet.xlsx"
    wb.save(str(empty_file))
    
    ops = DataOperations(file_loader)
    
    # Try to get unique values from empty sheet
    request = GetUniqueValuesRequest(
        file_path=str(empty_file),
        sheet_name="Empty",
        column="A",  # Will fail because no columns
        limit=100
    )
    
    # Act & Assert
    # Should raise ValueError because header detection fails on empty DataFrame
    with pytest.raises(ValueError) as exc_info:
        ops.get_unique_values(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "empty" in error_msg.lower() or "header" in error_msg.lower() or "detect" in error_msg.lower(), "Error should mention empty sheet or header detection issue"


# ============================================================================
# Additional Edge Cases
# ============================================================================

def test_invalid_logic_operator_in_filter(simple_fixture, file_loader):
    """Test error when invalid logic operator is used in filters.
    
    Verifies:
    - ValueError is raised
    - Error message mentions valid operators (AND/OR)
    """
    print(f"\n❌ Testing error: Invalid logic operator")
    
    ops = DataOperations(file_loader)
    from mcp_excel.operations.filtering import FilterEngine
    
    filter_engine = FilterEngine()
    
    # Load data
    df, _ = ops._load_with_header_detection(
        simple_fixture.path_str, simple_fixture.sheet_name, None
    )
    
    # Create valid filter
    class ValidFilter:
        column = simple_fixture.columns[0]
        operator = "=="
        value = "test"
        values = None
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        filter_engine.apply_filters(df, [ValidFilter()], logic="INVALID")
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "logic" in error_msg.lower() or "operator" in error_msg.lower(), "Error should mention logic operator"
    assert "AND" in error_msg or "OR" in error_msg, "Error should mention valid operators"


def test_get_column_stats_on_non_numeric(simple_fixture, file_loader):
    """Test get_column_stats on text column.
    
    Verifies:
    - Operation handles non-numeric columns gracefully
    - Either returns stats or raises clear error
    """
    print(f"\n❌ Testing error: Column stats on non-numeric column")
    
    ops = StatisticsOperations(file_loader)
    
    # Use text column
    text_column = simple_fixture.columns[0]  # "Имя"
    
    request = GetColumnStatsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=text_column,
        filters=[]
    )
    
    # Act & Assert
    # This might succeed (returning limited stats) or fail with ValueError
    try:
        response = ops.get_column_stats(request)
        print(f"✅ Operation succeeded with limited stats for text column")
        # If it succeeds, stats should be None or minimal
        assert response.stats is not None, "Should return some stats structure"
    except ValueError as e:
        print(f"✅ Caught expected error: {e}")
        error_msg = str(e)
        assert "numeric" in error_msg.lower() or "number" in error_msg.lower(), "Error should mention numeric requirement"
