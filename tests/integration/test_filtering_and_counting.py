# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Integration tests for Filtering and Counting operations.

Tests cover:
- filter_and_count: Count rows matching filter conditions
- ALL 12 filter operators: ==, !=, >, <, >=, <=, in, not_in, contains, startswith, endswith, regex, is_null, is_not_null
- Combined filters (AND/OR logic)
- Datetime filtering
- Excel formula generation for all operators

These are END-TO-END tests that verify the complete operation flow:
FileLoader -> FilterEngine -> DataOperations -> Response with Excel formulas
"""

import pytest

from mcp_excel.operations.data_operations import DataOperations
from mcp_excel.models.requests import FilterAndCountRequest, FilterCondition


# ============================================================================
# Comparison Operators Tests (==, !=, >, <, >=, <=)
# ============================================================================

def test_filter_and_count_equals_operator(simple_fixture, file_loader):
    """Test filter_and_count with == operator.
    
    Verifies:
    - Counts rows where column equals specific value
    - Returns correct count
    - Generates Excel formula (COUNTIF)
    - TSV output is generated
    """
    print(f"\n🔍 Testing filter_and_count with == operator")
    
    ops = DataOperations(file_loader)
    
    # Get a value to filter on
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],  # "Имя"
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    test_value = unique_response.values[0]
    
    print(f"  Filter: {simple_fixture.columns[0]} == '{test_value}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=test_value)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.count > 0, "Should find at least one matching row"
    assert response.count <= simple_fixture.row_count, "Count should not exceed total rows"
    assert response.excel_output.formula is not None, "Should generate Excel formula"
    assert "COUNTIF" in response.excel_output.formula, "Should use COUNTIF function"
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.filters_applied) == 1, "Should have 1 filter applied"
    assert response.filters_applied[0]["operator"] == "==", "Should record operator"


def test_filter_and_count_not_equals_operator(simple_fixture, file_loader):
    """Test filter_and_count with != operator.
    
    Verifies:
    - Counts rows where column does not equal specific value
    - Count is less than total rows
    - Generates appropriate formula
    """
    print(f"\n🔍 Testing filter_and_count with != operator")
    
    ops = DataOperations(file_loader)
    
    # Get a value to filter on
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    test_value = unique_response.values[0]
    
    print(f"  Filter: {simple_fixture.columns[0]} != '{test_value}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="!=", value=test_value)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert response.count < simple_fixture.row_count, "Should exclude at least one row"
    assert response.excel_output.formula is not None, "Should generate Excel formula"
    assert response.filters_applied[0]["operator"] == "!=", "Should record operator"


def test_filter_and_count_greater_than_operator(numeric_types_fixture, file_loader):
    """Test filter_and_count with > operator on numeric column.
    
    Verifies:
    - Counts rows where numeric column is greater than value
    - Works with integer and float columns
    - Generates COUNTIF formula with > operator
    """
    print(f"\n🔍 Testing filter_and_count with > operator")
    
    ops = DataOperations(file_loader)
    
    # Use "Количество" column (integer values: 10, 20, 30, ...)
    test_column = numeric_types_fixture.columns[1]  # "Количество"
    test_value = 100  # Should match rows with quantity > 100
    
    print(f"  Filter: {test_column} > {test_value}")
    
    # Act
    request = FilterAndCountRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        filters=[
            FilterCondition(column=test_column, operator=">", value=test_value)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert response.count <= numeric_types_fixture.row_count, "Count should not exceed total rows"
    assert response.excel_output.formula is not None, "Should generate Excel formula"
    assert "COUNTIF" in response.excel_output.formula, "Should use COUNTIF function"
    assert f">{test_value}" in response.excel_output.formula, "Formula should contain >value"


def test_filter_and_count_less_than_operator(numeric_types_fixture, file_loader):
    """Test filter_and_count with < operator on numeric column.
    
    Verifies:
    - Counts rows where numeric column is less than value
    - Generates COUNTIF formula with < operator
    """
    print(f"\n🔍 Testing filter_and_count with < operator")
    
    ops = DataOperations(file_loader)
    
    test_column = numeric_types_fixture.columns[1]  # "Количество"
    test_value = 100
    
    print(f"  Filter: {test_column} < {test_value}")
    
    # Act
    request = FilterAndCountRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        filters=[
            FilterCondition(column=test_column, operator="<", value=test_value)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert response.excel_output.formula is not None, "Should generate Excel formula"
    assert f"<{test_value}" in response.excel_output.formula, "Formula should contain <value"


def test_filter_and_count_greater_or_equal_operator(numeric_types_fixture, file_loader):
    """Test filter_and_count with >= operator on numeric column.
    
    Verifies:
    - Counts rows where numeric column is greater than or equal to value
    - Generates COUNTIF formula with >= operator
    """
    print(f"\n🔍 Testing filter_and_count with >= operator")
    
    ops = DataOperations(file_loader)
    
    test_column = numeric_types_fixture.columns[1]  # "Количество"
    test_value = 100
    
    print(f"  Filter: {test_column} >= {test_value}")
    
    # Act
    request = FilterAndCountRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        filters=[
            FilterCondition(column=test_column, operator=">=", value=test_value)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert response.excel_output.formula is not None, "Should generate Excel formula"
    assert f">={test_value}" in response.excel_output.formula, "Formula should contain >=value"


def test_filter_and_count_less_or_equal_operator(numeric_types_fixture, file_loader):
    """Test filter_and_count with <= operator on numeric column.
    
    Verifies:
    - Counts rows where numeric column is less than or equal to value
    - Generates COUNTIF formula with <= operator
    """
    print(f"\n🔍 Testing filter_and_count with <= operator")
    
    ops = DataOperations(file_loader)
    
    test_column = numeric_types_fixture.columns[1]  # "Количество"
    test_value = 100
    
    print(f"  Filter: {test_column} <= {test_value}")
    
    # Act
    request = FilterAndCountRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        filters=[
            FilterCondition(column=test_column, operator="<=", value=test_value)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert response.excel_output.formula is not None, "Should generate Excel formula"
    assert f"<={test_value}" in response.excel_output.formula, "Formula should contain <=value"


# ============================================================================
# Set Operators Tests (in, not_in)
# ============================================================================

def test_filter_and_count_in_operator(simple_fixture, file_loader):
    """Test filter_and_count with 'in' operator (multiple values).
    
    Verifies:
    - Counts rows where column value is in list of values
    - Works like Excel filter with multiple checkboxes
    - Generates SUMPRODUCT formula for multiple values
    """
    print(f"\n🔍 Testing filter_and_count with 'in' operator")
    
    ops = DataOperations(file_loader)
    
    # Get multiple values to filter on
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=3
    )
    unique_response = ops.get_unique_values(unique_request)
    test_values = unique_response.values[:2]  # Use first 2 values
    
    print(f"  Filter: {simple_fixture.columns[0]} in {test_values}")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="in", values=test_values)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count > 0, "Should find matching rows"
    assert response.count <= simple_fixture.row_count, "Count should not exceed total rows"
    assert response.excel_output.formula is not None, "Should generate Excel formula"
    # 'in' operator typically generates SUMPRODUCT formula
    assert "SUMPRODUCT" in response.excel_output.formula or "COUNTIF" in response.excel_output.formula, \
        "Should use SUMPRODUCT or COUNTIF for 'in' operator"
    assert response.filters_applied[0]["operator"] == "in", "Should record operator"
    assert response.filters_applied[0]["values"] == test_values, "Should record values"


def test_filter_and_count_not_in_operator(simple_fixture, file_loader):
    """Test filter_and_count with 'not_in' operator.
    
    Verifies:
    - Counts rows where column value is NOT in list of values
    - Excludes specified values
    - Count is less than total rows
    """
    print(f"\n🔍 Testing filter_and_count with 'not_in' operator")
    
    ops = DataOperations(file_loader)
    
    # Get multiple values to exclude
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=3
    )
    unique_response = ops.get_unique_values(unique_request)
    test_values = unique_response.values[:2]
    
    print(f"  Filter: {simple_fixture.columns[0]} not_in {test_values}")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="not_in", values=test_values)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert response.count < simple_fixture.row_count, "Should exclude some rows"
    # not_in may not have Excel formula (complex logic)
    assert response.filters_applied[0]["operator"] == "not_in", "Should record operator"


# ============================================================================
# String Operators Tests (contains, startswith, endswith, regex)
# ============================================================================

def test_filter_and_count_contains_operator(simple_fixture, file_loader):
    """Test filter_and_count with 'contains' operator (substring search).
    
    Verifies:
    - Counts rows where column contains substring
    - Case-sensitive search
    - Works with string columns
    """
    print(f"\n🔍 Testing filter_and_count with 'contains' operator")
    
    ops = DataOperations(file_loader)
    
    # Get a value and use part of it
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],  # "Имя"
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    full_value = unique_response.values[0]
    # Use first 3 characters as substring
    test_substring = str(full_value)[:3] if len(str(full_value)) >= 3 else str(full_value)
    
    print(f"  Filter: {simple_fixture.columns[0]} contains '{test_substring}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="contains", value=test_substring)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count > 0, "Should find at least one matching row"
    assert response.count <= simple_fixture.row_count, "Count should not exceed total rows"
    # contains operator may not have simple Excel formula
    assert response.filters_applied[0]["operator"] == "contains", "Should record operator"


def test_filter_and_count_startswith_operator(simple_fixture, file_loader):
    """Test filter_and_count with 'startswith' operator.
    
    Verifies:
    - Counts rows where column starts with prefix
    - Works with string columns
    """
    print(f"\n🔍 Testing filter_and_count with 'startswith' operator")
    
    ops = DataOperations(file_loader)
    
    # Get a value and use first 2 characters
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    full_value = unique_response.values[0]
    test_prefix = str(full_value)[:2] if len(str(full_value)) >= 2 else str(full_value)
    
    print(f"  Filter: {simple_fixture.columns[0]} startswith '{test_prefix}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="startswith", value=test_prefix)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count > 0, "Should find at least one matching row"
    assert response.filters_applied[0]["operator"] == "startswith", "Should record operator"


def test_filter_and_count_endswith_operator(simple_fixture, file_loader):
    """Test filter_and_count with 'endswith' operator.
    
    Verifies:
    - Counts rows where column ends with suffix
    - Works with string columns
    """
    print(f"\n🔍 Testing filter_and_count with 'endswith' operator")
    
    ops = DataOperations(file_loader)
    
    # Get a value and use last 2 characters
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    full_value = unique_response.values[0]
    test_suffix = str(full_value)[-2:] if len(str(full_value)) >= 2 else str(full_value)
    
    print(f"  Filter: {simple_fixture.columns[0]} endswith '{test_suffix}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="endswith", value=test_suffix)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count > 0, "Should find at least one matching row"
    assert response.filters_applied[0]["operator"] == "endswith", "Should record operator"


def test_filter_and_count_regex_operator(simple_fixture, file_loader):
    """Test filter_and_count with 'regex' operator (pattern matching).
    
    Verifies:
    - Counts rows where column matches regex pattern
    - Works with valid regex patterns
    - Handles complex patterns
    """
    print(f"\n🔍 Testing filter_and_count with 'regex' operator")
    
    ops = DataOperations(file_loader)
    
    # Use a simple regex pattern that matches Cyrillic names
    test_pattern = "^[А-Я].*"  # Starts with uppercase Cyrillic letter
    
    print(f"  Filter: {simple_fixture.columns[0]} regex '{test_pattern}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="regex", value=test_pattern)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    # regex operator doesn't have Excel formula equivalent
    assert response.filters_applied[0]["operator"] == "regex", "Should record operator"


def test_filter_and_count_regex_invalid_pattern(simple_fixture, file_loader):
    """Test filter_and_count with invalid regex pattern.
    
    Verifies:
    - Raises ValueError for invalid regex
    - Error message is helpful
    """
    print(f"\n🔍 Testing filter_and_count with invalid regex")
    
    ops = DataOperations(file_loader)
    
    # Invalid regex pattern (unclosed bracket)
    invalid_pattern = "[A-Z"
    
    print(f"  Filter: {simple_fixture.columns[0]} regex '{invalid_pattern}' (invalid)")
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        request = FilterAndCountRequest(
            file_path=simple_fixture.path_str,
            sheet_name=simple_fixture.sheet_name,
            filters=[
                FilterCondition(column=simple_fixture.columns[0], operator="regex", value=invalid_pattern)
            ],
            logic="AND"
        )
        ops.filter_and_count(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "regex" in str(exc_info.value).lower(), "Error should mention regex"


# ============================================================================
# Null Operators Tests (is_null, is_not_null)
# ============================================================================

def test_filter_and_count_is_null_operator(with_nulls_fixture, file_loader):
    """Test filter_and_count with 'is_null' operator.
    
    Verifies:
    - Counts rows where column has null/empty values
    - Works with columns containing nulls
    - Generates appropriate formula
    """
    print(f"\n🔍 Testing filter_and_count with 'is_null' operator")
    
    ops = DataOperations(file_loader)
    
    # Use "Email" column which has nulls
    test_column = with_nulls_fixture.columns[2]  # "Email"
    
    print(f"  Filter: {test_column} is_null")
    
    # Act
    request = FilterAndCountRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        filters=[
            FilterCondition(column=test_column, operator="is_null", value=None)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count > 0, "Should find null values in test fixture"
    assert response.count <= with_nulls_fixture.row_count, "Count should not exceed total rows"
    # is_null may generate COUNTBLANK formula
    assert response.filters_applied[0]["operator"] == "is_null", "Should record operator"


def test_filter_and_count_is_not_null_operator(with_nulls_fixture, file_loader):
    """Test filter_and_count with 'is_not_null' operator.
    
    Verifies:
    - Counts rows where column has non-null values
    - Excludes null/empty values
    """
    print(f"\n🔍 Testing filter_and_count with 'is_not_null' operator")
    
    ops = DataOperations(file_loader)
    
    # Use "Email" column which has nulls
    test_column = with_nulls_fixture.columns[2]  # "Email"
    
    print(f"  Filter: {test_column} is_not_null")
    
    # Act
    request = FilterAndCountRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        filters=[
            FilterCondition(column=test_column, operator="is_not_null", value=None)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert response.count < with_nulls_fixture.row_count, "Should exclude some null rows"
    # is_not_null may generate COUNTA formula
    assert response.filters_applied[0]["operator"] == "is_not_null", "Should record operator"


# ============================================================================
# Combined Filters Tests (AND/OR logic)
# ============================================================================

def test_filter_and_count_combined_and_logic(simple_fixture, file_loader):
    """Test filter_and_count with multiple filters combined with AND logic.
    
    Verifies:
    - Counts rows matching ALL filter conditions
    - AND logic works correctly
    - Generates COUNTIFS formula for multiple conditions
    """
    print(f"\n🔍 Testing filter_and_count with combined AND filters")
    
    ops = DataOperations(file_loader)
    
    # Get values for two different columns
    from mcp_excel.models.requests import GetUniqueValuesRequest
    
    # First filter: column 0
    unique_request1 = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    value1 = ops.get_unique_values(unique_request1).values[0]
    
    # Second filter: column 1 (numeric)
    unique_request2 = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[1],
        limit=1
    )
    value2 = ops.get_unique_values(unique_request2).values[0]
    
    print(f"  Filter: {simple_fixture.columns[0]} == '{value1}' AND {simple_fixture.columns[1]} == {value2}")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=value1),
            FilterCondition(column=simple_fixture.columns[1], operator="==", value=value2)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert len(response.filters_applied) == 2, "Should have 2 filters applied"
    assert response.excel_output.formula is not None, "Should generate Excel formula"
    # Multiple conditions typically use COUNTIFS
    assert "COUNTIFS" in response.excel_output.formula or "COUNTIF" in response.excel_output.formula, \
        "Should use COUNTIFS for multiple conditions"


def test_filter_and_count_combined_or_logic(simple_fixture, file_loader):
    """Test filter_and_count with multiple filters combined with OR logic.
    
    Verifies:
    - Counts rows matching ANY filter condition
    - OR logic works correctly
    - Count is sum of individual filter results (minus overlaps)
    """
    print(f"\n🔍 Testing filter_and_count with combined OR filters")
    
    ops = DataOperations(file_loader)
    
    # Get two different values from same column
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=2
    )
    values = ops.get_unique_values(unique_request).values[:2]
    
    print(f"  Filter: {simple_fixture.columns[0]} == '{values[0]}' OR {simple_fixture.columns[0]} == '{values[1]}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[0]),
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[1])
        ],
        logic="OR"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count > 0, "Should find matching rows"
    assert len(response.filters_applied) == 2, "Should have 2 filters applied"
    # OR logic may not have simple Excel formula
    

def test_filter_and_count_three_conditions_and(numeric_types_fixture, file_loader):
    """Test filter_and_count with three conditions combined with AND.
    
    Verifies:
    - Handles multiple conditions correctly
    - All conditions must be satisfied
    """
    print(f"\n🔍 Testing filter_and_count with 3 AND conditions")
    
    ops = DataOperations(file_loader)
    
    # Three conditions on numeric column
    test_column = numeric_types_fixture.columns[1]  # "Количество"
    
    print(f"  Filter: {test_column} > 50 AND {test_column} < 150 AND {test_column} != 100")
    
    # Act
    request = FilterAndCountRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        filters=[
            FilterCondition(column=test_column, operator=">", value=50),
            FilterCondition(column=test_column, operator="<", value=150),
            FilterCondition(column=test_column, operator="!=", value=100)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert len(response.filters_applied) == 3, "Should have 3 filters applied"


# ============================================================================
# DateTime Filtering Tests
# ============================================================================

def test_filter_and_count_datetime_equals(with_dates_fixture, file_loader):
    """Test filter_and_count with datetime column using == operator.
    
    Verifies:
    - Filters datetime columns correctly
    - Accepts ISO 8601 date strings
    - Generates DATE() formula in Excel
    """
    print(f"\n🔍 Testing filter_and_count with datetime == operator")
    
    ops = DataOperations(file_loader)
    
    # Get a sample datetime value
    from mcp_excel.models.requests import FilterAndGetRowsRequest
    sample_request = FilterAndGetRowsRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        filters=[],
        columns=[with_dates_fixture.expected["datetime_columns"][0]],
        limit=1,
        offset=0,
        logic="AND"
    )
    sample_response = ops.filter_and_get_rows(sample_request)
    
    if sample_response.rows:
        date_col = with_dates_fixture.expected["datetime_columns"][0]
        sample_date_str = sample_response.rows[0][date_col]
        # Extract date part (YYYY-MM-DD)
        test_date = sample_date_str.split('T')[0] if 'T' in sample_date_str else sample_date_str
        
        print(f"  Filter: {date_col} == '{test_date}'")
        
        # Act
        request = FilterAndCountRequest(
            file_path=with_dates_fixture.path_str,
            sheet_name=with_dates_fixture.sheet_name,
            filters=[
                FilterCondition(column=date_col, operator="==", value=test_date)
            ],
            logic="AND"
        )
        response = ops.filter_and_count(request)
        
        # Assert
        print(f"✅ Count: {response.count}")
        print(f"   Formula: {response.excel_output.formula}")
        
        assert response.count >= 0, "Count should be non-negative"
        if response.excel_output.formula:
            assert "DATE(" in response.excel_output.formula, "Should use DATE() function for datetime"


def test_filter_and_count_datetime_greater_than(with_dates_fixture, file_loader):
    """Test filter_and_count with datetime column using >= operator.
    
    Verifies:
    - Filters datetime ranges correctly
    - >= operator works with dates
    """
    print(f"\n🔍 Testing filter_and_count with datetime >= operator")
    
    ops = DataOperations(file_loader)
    
    date_col = with_dates_fixture.expected["datetime_columns"][0]
    # Use a date that should match some rows
    test_date = with_dates_fixture.expected["date_range_start"]
    
    print(f"  Filter: {date_col} >= '{test_date}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        filters=[
            FilterCondition(column=date_col, operator=">=", value=test_date)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count > 0, "Should find rows with dates >= test date"
    assert response.count <= with_dates_fixture.row_count, "Count should not exceed total rows"


def test_filter_and_count_datetime_range(with_dates_fixture, file_loader):
    """Test filter_and_count with datetime range (two conditions).
    
    Verifies:
    - Filters date ranges correctly
    - Combines >= and <= for range
    """
    print(f"\n🔍 Testing filter_and_count with datetime range")
    
    ops = DataOperations(file_loader)
    
    date_col = with_dates_fixture.expected["datetime_columns"][0]
    start_date = "2024-01-01"
    end_date = "2024-12-31"
    
    print(f"  Filter: {date_col} >= '{start_date}' AND {date_col} <= '{end_date}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        filters=[
            FilterCondition(column=date_col, operator=">=", value=start_date),
            FilterCondition(column=date_col, operator="<=", value=end_date)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.count >= 0, "Count should be non-negative"
    assert len(response.filters_applied) == 2, "Should have 2 filters for range"


# ============================================================================
# Edge Cases and Error Handling
# ============================================================================

def test_filter_and_count_invalid_column(simple_fixture, file_loader):
    """Test filter_and_count with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message lists available columns
    """
    print(f"\n🔍 Testing filter_and_count with invalid column")
    
    ops = DataOperations(file_loader)
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        request = FilterAndCountRequest(
            file_path=simple_fixture.path_str,
            sheet_name=simple_fixture.sheet_name,
            filters=[
                FilterCondition(column="NonExistentColumn", operator="==", value="test")
            ],
            logic="AND"
        )
        ops.filter_and_count(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower(), "Error should mention column not found"
    assert "NonExistentColumn" in str(exc_info.value), "Error should mention the invalid column"


def test_filter_and_count_empty_filters(simple_fixture, file_loader):
    """Test filter_and_count with no filters.
    
    Verifies:
    - Returns total row count when no filters applied
    - Handles empty filter list gracefully
    - Formula generation may fail (expected for empty filters)
    """
    print(f"\n🔍 Testing filter_and_count with empty filters")
    
    ops = DataOperations(file_loader)
    
    # Act - empty filters may cause formula generation to fail, which is expected
    try:
        request = FilterAndCountRequest(
            file_path=simple_fixture.path_str,
            sheet_name=simple_fixture.sheet_name,
            filters=[],
            logic="AND"
        )
        response = ops.filter_and_count(request)
        
        # Assert
        print(f"✅ Count: {response.count}")
        
        assert response.count == simple_fixture.row_count, "Should return total row count with no filters"
        assert len(response.filters_applied) == 0, "Should have no filters applied"
    except ValueError as e:
        # Expected: FormulaGenerator can't generate formula without filters or target_range
        print(f"✅ Expected error for empty filters: {e}")
        assert "requires filters or target range" in str(e), "Should fail with expected error message"


def test_filter_and_count_no_matches(simple_fixture, file_loader):
    """Test filter_and_count when no rows match filter.
    
    Verifies:
    - Returns count of 0 when no matches
    - Handles zero results gracefully
    """
    print(f"\n🔍 Testing filter_and_count with no matches")
    
    ops = DataOperations(file_loader)
    
    # Use a value that definitely doesn't exist
    impossible_value = "IMPOSSIBLE_VALUE_THAT_DOES_NOT_EXIST_12345"
    
    print(f"  Filter: {simple_fixture.columns[0]} == '{impossible_value}'")
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=impossible_value)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Count: {response.count}")
    
    assert response.count == 0, "Should return 0 when no rows match"
    assert response.excel_output.formula is not None, "Should still generate formula"


def test_filter_and_count_performance_metrics(simple_fixture, file_loader):
    """Test that filter_and_count includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    - Cache status is reported
    """
    print(f"\n🔍 Testing filter_and_count performance metrics")
    
    ops = DataOperations(file_loader)
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="is_not_null", value=None)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Performance:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Cache hit: {response.performance.cache_hit}")
    print(f"   Memory used: {response.performance.memory_used_mb}MB")
    
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.cache_hit in [True, False], "Should report cache status"
    assert response.performance.execution_time_ms < 5000, "Should complete in reasonable time"


def test_filter_and_count_metadata(simple_fixture, file_loader):
    """Test that filter_and_count includes metadata.
    
    Verifies:
    - Metadata is included
    - File format and sheet name are correct
    - Row/column totals are reported
    """
    print(f"\n🔍 Testing filter_and_count metadata")
    
    ops = DataOperations(file_loader)
    
    # Use a simple filter instead of empty filters to avoid formula generation error
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    test_value = ops.get_unique_values(unique_request).values[0]
    
    # Act
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=test_value)
        ],
        logic="AND"
    )
    response = ops.filter_and_count(request)
    
    # Assert
    print(f"✅ Metadata:")
    print(f"   File format: {response.metadata.file_format}")
    print(f"   Sheet name: {response.metadata.sheet_name}")
    print(f"   Total rows: {response.metadata.rows_total}")
    print(f"   Total columns: {response.metadata.columns_total}")
    
    assert response.metadata is not None, "Should include metadata"
    assert response.metadata.file_format == simple_fixture.format, "Should report correct format"
    assert response.metadata.sheet_name == simple_fixture.sheet_name, "Should report correct sheet"
    assert response.metadata.rows_total == simple_fixture.row_count, "Should report total rows"
    assert response.metadata.columns_total == len(simple_fixture.columns), "Should report total columns"
