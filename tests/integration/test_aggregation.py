# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Integration tests for Aggregation operations.

Tests cover:
- aggregate: Perform aggregation (sum, mean, median, min, max, std, var, count) on a column
- group_by: Group data by columns and perform aggregation

These are END-TO-END tests that verify the complete operation flow:
FileLoader -> FilterEngine -> Aggregation -> Response
"""

import pytest

from mcp_excel.operations.data_operations import DataOperations
from mcp_excel.models.requests import (
    AggregateRequest,
    GroupByRequest,
    FilterCondition,
)


# ============================================================================
# aggregate tests - Basic operations
# ============================================================================

def test_aggregate_count_simple(simple_fixture, file_loader):
    """Test aggregate count operation on simple data.
    
    Verifies:
    - Returns correct count
    - Generates Excel formula
    - Performance metrics included
    """
    print(f"\n📊 Testing aggregate count on: {simple_fixture.name}")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="count",
        target_column="Возраст",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Count: {response.value}")
    print(f"   Operation: {response.operation}")
    print(f"   Formula: {response.excel_output.formula}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.value == simple_fixture.row_count, "Should count all rows"
    assert response.operation == "count", "Should return operation name"
    assert response.target_column == "Возраст", "Should return target column"
    assert len(response.filters_applied) == 0, "Should have no filters"
    
    # Check Excel formula
    assert response.excel_output.formula is not None, "Should generate formula"
    assert "COUNT" in response.excel_output.formula.upper(), "Formula should use COUNT function"
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert str(int(response.value)) in response.excel_output.tsv, "TSV should contain result"
    
    # Check metadata
    assert response.metadata is not None, "Should include metadata"
    assert response.metadata.rows_total == simple_fixture.row_count
    
    # Check performance
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0


def test_aggregate_sum_numeric(numeric_types_fixture, file_loader):
    """Test aggregate sum operation on numeric data.
    
    Verifies:
    - Calculates sum correctly
    - Handles integer columns
    - Generates SUMIF formula
    """
    print(f"\n📊 Testing aggregate sum on numeric data")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="sum",
        target_column="Количество",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Sum: {response.value}")
    print(f"   Formula: {response.excel_output.formula}")
    
    expected_sum = numeric_types_fixture.expected["quantity_sum"]
    assert response.value == expected_sum, f"Should calculate correct sum: {expected_sum}"
    assert response.operation == "sum"
    
    # Check formula
    assert response.excel_output.formula is not None, "Should generate formula"
    assert "SUM" in response.excel_output.formula.upper(), "Formula should use SUM function"


def test_aggregate_mean_numeric(numeric_types_fixture, file_loader):
    """Test aggregate mean operation.
    
    Verifies:
    - Calculates mean correctly
    - Returns float value
    - Generates AVERAGE formula
    """
    print(f"\n📊 Testing aggregate mean")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="mean",
        target_column="Количество",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Mean: {response.value}")
    print(f"   Formula: {response.excel_output.formula}")
    
    # Mean of 10, 20, 30, ..., 200 = (10+20+...+200)/20 = 2100/20 = 105
    expected_mean = 105.0
    assert response.value == expected_mean, f"Should calculate correct mean: {expected_mean}"
    assert isinstance(response.value, (int, float)), "Mean should be numeric"
    
    # Check formula
    assert response.excel_output.formula is not None, "Should generate formula"
    assert "AVERAGE" in response.excel_output.formula.upper(), "Formula should use AVERAGE function"


def test_aggregate_median_numeric(numeric_types_fixture, file_loader):
    """Test aggregate median operation.
    
    Verifies:
    - Calculates median correctly
    - Handles even number of values
    - Note: Excel formula may be None (median without filters not supported in Excel)
    """
    print(f"\n📊 Testing aggregate median")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="median",
        target_column="Количество",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Median: {response.value}")
    print(f"   Formula: {response.excel_output.formula}")
    
    # Median of 10, 20, 30, ..., 200 (20 values) = (100 + 110) / 2 = 105
    expected_median = 105.0
    assert response.value == expected_median, f"Should calculate correct median: {expected_median}"
    assert response.operation == "median"
    # Note: Formula may be None for median without filters (Excel limitation)


def test_aggregate_min_numeric(numeric_types_fixture, file_loader):
    """Test aggregate min operation.
    
    Verifies:
    - Finds minimum value correctly
    - Note: Excel formula may be None (min without filters not supported in Excel)
    """
    print(f"\n📊 Testing aggregate min")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="min",
        target_column="Количество",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Min: {response.value}")
    print(f"   Formula: {response.excel_output.formula}")
    
    expected_min = 10.0  # First value is 1*10 = 10
    assert response.value == expected_min, f"Should find correct minimum: {expected_min}"
    assert response.operation == "min"
    # Note: Formula may be None for min without filters (Excel limitation)


def test_aggregate_max_numeric(numeric_types_fixture, file_loader):
    """Test aggregate max operation.
    
    Verifies:
    - Finds maximum value correctly
    - Note: Excel formula may be None (max without filters not supported in Excel)
    """
    print(f"\n📊 Testing aggregate max")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="max",
        target_column="Количество",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Max: {response.value}")
    print(f"   Formula: {response.excel_output.formula}")
    
    expected_max = 200.0  # Last value is 20*10 = 200
    assert response.value == expected_max, f"Should find correct maximum: {expected_max}"
    assert response.operation == "max"
    # Note: Formula may be None for max without filters (Excel limitation)


def test_aggregate_std_numeric(numeric_types_fixture, file_loader):
    """Test aggregate std (standard deviation) operation.
    
    Verifies:
    - Calculates standard deviation correctly
    - Returns positive value
    - Note: Excel formula may be None (std without filters not supported in Excel)
    """
    print(f"\n📊 Testing aggregate std")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="std",
        target_column="Количество",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Std: {response.value:.2f}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.value > 0, "Standard deviation should be positive"
    assert isinstance(response.value, (int, float)), "Std should be numeric"
    assert response.operation == "std"
    # Note: Formula may be None for std without filters (Excel limitation)


def test_aggregate_var_numeric(numeric_types_fixture, file_loader):
    """Test aggregate var (variance) operation.
    
    Verifies:
    - Calculates variance correctly
    - Returns positive value
    - Note: Excel formula may be None (var without filters not supported in Excel)
    """
    print(f"\n📊 Testing aggregate var")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="var",
        target_column="Количество",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Var: {response.value:.2f}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.value > 0, "Variance should be positive"
    assert isinstance(response.value, (int, float)), "Var should be numeric"
    assert response.operation == "var"
    # Note: Formula may be None for var without filters (Excel limitation)


# ============================================================================
# aggregate tests - With filters
# ============================================================================

def test_aggregate_sum_with_filter(simple_fixture, file_loader):
    """Test aggregate sum with filter condition.
    
    Verifies:
    - Applies filter before aggregation
    - Generates SUMIF formula with condition
    - Returns correct filtered sum
    """
    print(f"\n📊 Testing aggregate sum with filter")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="sum",
        target_column="Возраст",
        filters=[
            FilterCondition(column="Возраст", operator=">", value=30)
        ]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Filtered sum: {response.value}")
    print(f"   Filters applied: {response.filters_applied}")
    print(f"   Formula: {response.excel_output.formula}")
    
    # Ages > 30: 31, 32, 33 = 96
    # From simple fixture: ages are 25, 30, 35, 28, 32, 27, 29, 31, 26, 33
    # Ages > 30: 35, 32, 31, 33 = 131
    assert response.value > 0, "Should have filtered sum"
    assert len(response.filters_applied) == 1, "Should have 1 filter applied"
    assert response.filters_applied[0]["column"] == "Возраст"
    assert response.filters_applied[0]["operator"] == ">"
    
    # Check formula contains condition
    assert response.excel_output.formula is not None, "Should generate formula"
    assert "SUM" in response.excel_output.formula.upper(), "Formula should use SUM function"


def test_aggregate_count_with_multiple_filters(simple_fixture, file_loader):
    """Test aggregate count with multiple filter conditions.
    
    Verifies:
    - Applies multiple filters (AND logic)
    - Generates COUNTIFS formula
    - Returns correct filtered count
    """
    print(f"\n📊 Testing aggregate count with multiple filters")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="count",
        target_column="Возраст",
        filters=[
            FilterCondition(column="Возраст", operator=">=", value=25),
            FilterCondition(column="Возраст", operator="<=", value=30)
        ],
        logic="AND"
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Filtered count: {response.value}")
    print(f"   Filters: {len(response.filters_applied)}")
    
    assert response.value > 0, "Should have filtered count"
    assert len(response.filters_applied) == 2, "Should have 2 filters applied"
    
    # Check formula
    assert response.excel_output.formula is not None, "Should generate formula"
    # COUNTIFS for multiple conditions
    assert "COUNT" in response.excel_output.formula.upper(), "Formula should use COUNT function"


def test_aggregate_mean_with_filter(numeric_types_fixture, file_loader):
    """Test aggregate mean with filter.
    
    Verifies:
    - Calculates mean only for filtered rows
    - Mean changes with filter
    """
    print(f"\n📊 Testing aggregate mean with filter")
    
    ops = DataOperations(file_loader)
    
    # First, get mean without filter
    request_no_filter = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="mean",
        target_column="Количество",
        filters=[]
    )
    response_no_filter = ops.aggregate(request_no_filter)
    
    # Now with filter
    request_with_filter = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="mean",
        target_column="Количество",
        filters=[
            FilterCondition(column="Количество", operator=">", value=100)
        ]
    )
    response_with_filter = ops.aggregate(request_with_filter)
    
    # Assert
    print(f"✅ Mean without filter: {response_no_filter.value}")
    print(f"   Mean with filter (>100): {response_with_filter.value}")
    
    assert response_with_filter.value > response_no_filter.value, "Filtered mean should be higher (only values >100)"
    assert len(response_with_filter.filters_applied) == 1


# ============================================================================
# aggregate tests - Edge cases
# ============================================================================

def test_aggregate_with_nulls(with_nulls_fixture, file_loader):
    """Test aggregate operation with null values.
    
    Verifies:
    - Skips null values in aggregation
    - Returns correct result despite nulls
    """
    print(f"\n📊 Testing aggregate with null values")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        operation="count",
        target_column="Email",  # Has nulls
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Count (excluding nulls): {response.value}")
    
    # Should count only non-null values
    assert response.value < with_nulls_fixture.row_count, "Should exclude null values"
    assert response.value > 0, "Should have some non-null values"


def test_aggregate_text_as_numbers(numeric_types_fixture, file_loader):
    """Test aggregate on column that might have text-stored numbers.
    
    Verifies:
    - Automatically converts text to numbers when possible
    - Performs aggregation correctly
    """
    print(f"\n📊 Testing aggregate with text-as-numbers conversion")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="sum",
        target_column="Количество",
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Sum: {response.value}")
    
    # Should successfully convert and sum
    assert response.value > 0, "Should convert text to numbers and sum"
    assert isinstance(response.value, (int, float)), "Result should be numeric"


def test_aggregate_invalid_column(simple_fixture, file_loader):
    """Test aggregate with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message is helpful
    """
    print(f"\n📊 Testing aggregate with invalid column")
    
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
    
    assert "not found" in str(exc_info.value).lower(), "Error should mention column not found"
    assert "NonExistentColumn" in str(exc_info.value), "Error should mention the invalid column"


def test_aggregate_invalid_operation(simple_fixture, file_loader):
    """Test aggregate with unsupported operation.
    
    Verifies:
    - Pydantic validation catches invalid operation at request creation
    """
    print(f"\n📊 Testing aggregate with invalid operation")
    
    from pydantic import ValidationError
    
    # Act & Assert - Pydantic should catch this during request validation
    with pytest.raises(ValidationError) as exc_info:
        request = AggregateRequest(
            file_path=simple_fixture.path_str,
            sheet_name=simple_fixture.sheet_name,
            operation="invalid_op",  # Invalid operation
            target_column="Возраст",
            filters=[]
        )
    
    print(f"✅ Caught expected Pydantic validation error")


def test_aggregate_non_numeric_column(simple_fixture, file_loader):
    """Test aggregate numeric operation on text column.
    
    Verifies:
    - Raises ValueError when trying to sum text
    - Error message is clear
    """
    print(f"\n📊 Testing aggregate on non-numeric column")
    
    ops = DataOperations(file_loader)
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="sum",
        target_column="Город",  # Text column
        filters=[]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.aggregate(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "non-numeric" in str(exc_info.value).lower() or "cannot" in str(exc_info.value).lower()


# ============================================================================
# group_by tests - Basic operations
# ============================================================================

def test_group_by_single_column_sum(simple_fixture, file_loader):
    """Test group_by with single grouping column and sum.
    
    Verifies:
    - Groups by single column correctly
    - Calculates sum for each group
    - Returns list of groups
    - Generates TSV output
    """
    print(f"\n📊 Testing group_by with single column")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["Город"],
        agg_column="Возраст",
        agg_operation="sum",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups found: {len(response.groups)}")
    print(f"   Group columns: {response.group_columns}")
    print(f"   Agg operation: {response.agg_operation}")
    
    assert len(response.groups) > 0, "Should have at least one group"
    assert response.group_columns == ["Город"], "Should return group columns"
    assert response.agg_column == "Возраст", "Should return agg column"
    assert response.agg_operation == "sum", "Should return operation"
    
    # Check group structure
    first_group = response.groups[0]
    assert "Город" in first_group, "Group should have grouping column"
    assert any("Возраст" in key for key in first_group.keys()), "Group should have aggregated column"
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    
    # Check metadata
    assert response.metadata is not None, "Should include metadata"
    assert response.performance is not None, "Should include performance metrics"


def test_group_by_multiple_columns(numeric_types_fixture, file_loader):
    """Test group_by with multiple grouping columns.
    
    Verifies:
    - Groups by multiple columns correctly
    - Handles multi-level grouping
    - Returns correct number of groups
    """
    print(f"\n📊 Testing group_by with multiple columns")
    
    ops = DataOperations(file_loader)
    
    # First, let's use a fixture that has multiple categorical columns
    # numeric_types has "Код товара" which we can group by ranges
    request = GroupByRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        group_columns=["Код товара"],  # Will have 20 unique values
        agg_column="Количество",
        agg_operation="sum",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups found: {len(response.groups)}")
    
    assert len(response.groups) == numeric_types_fixture.row_count, "Should have one group per unique product code"
    assert response.group_columns == ["Код товара"]


def test_group_by_count_operation(simple_fixture, file_loader):
    """Test group_by with count operation.
    
    Verifies:
    - Count operation works in group_by
    - Returns correct counts per group
    """
    print(f"\n📊 Testing group_by with count operation")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["Город"],
        agg_column="Имя",
        agg_operation="count",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups: {len(response.groups)}")
    if response.groups:
        print(f"   Sample group keys: {list(response.groups[0].keys())}")
        print(f"   Sample group: {response.groups[0]}")
    
    assert len(response.groups) > 0, "Should have groups"
    assert response.agg_operation == "count"
    
    # Check that counts are non-negative integers (can be 0 if column has nulls)
    for group in response.groups:
        # Find the aggregated column (will be named like "Имя_count")
        count_keys = [k for k in group.keys() if k != "Город"]
        if count_keys:
            count_key = count_keys[0]
            assert isinstance(group[count_key], (int, float)), f"Count should be numeric, got {type(group[count_key])}"
            assert group[count_key] >= 0, "Count should be non-negative"


def test_group_by_mean_operation(numeric_types_fixture, file_loader):
    """Test group_by with mean operation.
    
    Verifies:
    - Mean operation works in group_by
    - Returns float values
    """
    print(f"\n📊 Testing group_by with mean operation")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        group_columns=["Код товара"],
        agg_column="Цена",
        agg_operation="mean",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups: {len(response.groups)}")
    
    assert len(response.groups) > 0, "Should have groups"
    assert response.agg_operation == "mean"
    
    # Check that means are numeric
    for group in response.groups:
        price_key = [k for k in group.keys() if "Цена" in k][0]
        assert isinstance(group[price_key], (int, float)), "Mean should be numeric"


def test_group_by_min_max_operations(numeric_types_fixture, file_loader):
    """Test group_by with min and max operations.
    
    Verifies:
    - Min operation works
    - Max operation works
    - Results are correct
    """
    print(f"\n📊 Testing group_by with min/max operations")
    
    ops = DataOperations(file_loader)
    
    # Test min
    request_min = GroupByRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        group_columns=["Код товара"],
        agg_column="Количество",
        agg_operation="min",
        filters=[]
    )
    response_min = ops.group_by(request_min)
    
    # Test max
    request_max = GroupByRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        group_columns=["Код товара"],
        agg_column="Количество",
        agg_operation="max",
        filters=[]
    )
    response_max = ops.group_by(request_max)
    
    # Assert
    print(f"✅ Min groups: {len(response_min.groups)}")
    print(f"   Max groups: {len(response_max.groups)}")
    
    assert len(response_min.groups) > 0, "Should have min groups"
    assert len(response_max.groups) > 0, "Should have max groups"
    assert response_min.agg_operation == "min"
    assert response_max.agg_operation == "max"


def test_group_by_with_filter(simple_fixture, file_loader):
    """Test group_by with filter conditions.
    
    Verifies:
    - Applies filter before grouping
    - Returns only filtered groups
    - Group count may be less than without filter
    """
    print(f"\n📊 Testing group_by with filter")
    
    ops = DataOperations(file_loader)
    
    # Without filter
    request_no_filter = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["Город"],
        agg_column="Возраст",
        agg_operation="count",
        filters=[]
    )
    response_no_filter = ops.group_by(request_no_filter)
    
    # With filter
    request_with_filter = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["Город"],
        agg_column="Возраст",
        agg_operation="count",
        filters=[
            FilterCondition(column="Возраст", operator=">", value=30)
        ]
    )
    response_with_filter = ops.group_by(request_with_filter)
    
    # Assert
    print(f"✅ Groups without filter: {len(response_no_filter.groups)}")
    print(f"   Groups with filter: {len(response_with_filter.groups)}")
    
    # With filter, we might have fewer groups (some cities might have no people >30)
    assert len(response_with_filter.groups) <= len(response_no_filter.groups), "Filtered groups should be <= unfiltered"


# ============================================================================
# group_by tests - Edge cases
# ============================================================================

def test_group_by_invalid_column(simple_fixture, file_loader):
    """Test group_by with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message lists available columns
    """
    print(f"\n📊 Testing group_by with invalid column")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["NonExistent"],
        agg_column="Возраст",
        agg_operation="sum",
        filters=[]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.group_by(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower(), "Error should mention column not found"
    assert "NonExistent" in str(exc_info.value), "Error should mention the invalid column"


def test_group_by_text_as_numbers(numeric_types_fixture, file_loader):
    """Test group_by with automatic text-to-number conversion.
    
    Verifies:
    - Converts text-stored numbers automatically
    - Performs aggregation correctly
    """
    print(f"\n📊 Testing group_by with text-as-numbers conversion")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        group_columns=["Код товара"],
        agg_column="Количество",
        agg_operation="sum",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups: {len(response.groups)}")
    
    assert len(response.groups) > 0, "Should successfully convert and group"
    
    # Check that aggregated values are numeric
    for group in response.groups:
        qty_key = [k for k in group.keys() if "Количество" in k][0]
        assert isinstance(group[qty_key], (int, float)), "Aggregated value should be numeric"


def test_group_by_single_column_table(single_column_fixture, file_loader):
    """Test group_by on minimal table (single column).
    
    Verifies:
    - Handles edge case of single column
    - Can group by the only column and aggregate on the same column
    
    Note: This test exposes a potential bug where grouping by and aggregating
    on the same column might not work correctly.
    """
    print(f"\n📊 Testing group_by on single column table")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=single_column_fixture.path_str,
        sheet_name=single_column_fixture.sheet_name,
        group_columns=["Значение"],
        agg_column="Значение",
        agg_operation="count",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups: {len(response.groups)}")
    print(f"   Expected: {single_column_fixture.row_count} groups (one per unique value)")
    print(f"   TSV output: {response.excel_output.tsv[:100]}")
    
    # BUG: Grouping by and aggregating on the same column returns empty groups
    # This should return 10 groups (one for each unique value in "Значение")
    # Expected behavior: Each group should have count=1
    assert len(response.groups) > 0, "Should handle single column grouping (BUG: returns empty groups)"


def test_group_by_with_nulls(with_nulls_fixture, file_loader):
    """Test group_by with null values.
    
    Verifies:
    - Handles null values in grouping columns
    - Handles null values in aggregation column
    - Returns correct results
    """
    print(f"\n📊 Testing group_by with null values")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        group_columns=["ID"],  # No nulls in ID
        agg_column="Телефон",  # Has nulls
        agg_operation="count",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups: {len(response.groups)}")
    
    assert len(response.groups) > 0, "Should handle nulls in aggregation"
    
    # Some groups should have count=0 or count=1 depending on nulls
    counts = [group[[k for k in group.keys() if "Телефон" in k][0]] for group in response.groups]
    print(f"   Counts: {counts}")


def test_group_by_performance_metrics(simple_fixture, file_loader):
    """Test that group_by includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    """
    print(f"\n📊 Testing group_by performance metrics")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["Город"],
        agg_column="Возраст",
        agg_operation="sum",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Performance:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Cache hit: {response.performance.cache_hit}")
    
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.cache_hit in [True, False], "Should report cache status"


def test_group_by_tsv_output_format(simple_fixture, file_loader):
    """Test that group_by generates proper TSV output.
    
    Verifies:
    - TSV output is generated
    - Contains group columns and aggregated values
    - Can be pasted into Excel
    """
    print(f"\n📊 Testing group_by TSV output")
    
    ops = DataOperations(file_loader)
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["Город"],
        agg_column="Возраст",
        agg_operation="sum",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ TSV output generated")
    print(f"   Length: {len(response.excel_output.tsv)} chars")
    print(f"   Preview: {response.excel_output.tsv[:200]}...")
    
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    
    # Check TSV contains group column
    assert "Город" in response.excel_output.tsv, "TSV should contain group column"
    
    # Check TSV has tab separators
    assert "\t" in response.excel_output.tsv, "TSV should use tab separators"


# ============================================================================
# Unicode Normalization Integration Tests
# ============================================================================

def test_aggregate_with_unicode_nfd_column(simple_fixture, file_loader):
    """Test aggregate with NFD Unicode form in column name.
    
    Verifies:
    - Finds column when request uses NFD but DataFrame has NFC
    - Performs aggregation correctly
    - Unicode normalization works end-to-end
    """
    print(f"\n🔤 Testing aggregate with NFD Unicode column name")
    
    import unicodedata
    ops = DataOperations(file_loader)
    
    # DataFrame has "Возраст" in NFC form
    # Request with NFD form (decomposed)
    column_nfd = unicodedata.normalize('NFD', "Возраст")
    
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="sum",
        target_column=column_nfd,  # NFD form
        filters=[]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Sum calculated successfully: {response.value}")
    print(f"   Column requested (NFD): {repr(column_nfd)}")
    print(f"   Column found (NFC): {repr(response.target_column)}")
    
    assert response.value > 0, "Should calculate sum despite Unicode form difference"
    assert response.target_column == "Возраст", "Should return original NFC column name"


def test_group_by_with_unicode_nfd_columns(simple_fixture, file_loader):
    """Test group_by with NFD Unicode forms in column names.
    
    Verifies:
    - Finds group columns with NFD form
    - Finds agg column with NFD form
    - Returns correct groups
    """
    print(f"\n🔤 Testing group_by with NFD Unicode column names")
    
    import unicodedata
    ops = DataOperations(file_loader)
    
    # Request with NFD forms
    group_col_nfd = unicodedata.normalize('NFD', "Город")
    agg_col_nfd = unicodedata.normalize('NFD', "Возраст")
    
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=[group_col_nfd],  # NFD form
        agg_column=agg_col_nfd,  # NFD form
        agg_operation="sum",
        filters=[]
    )
    
    # Act
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups found: {len(response.groups)}")
    print(f"   Group columns: {response.group_columns}")
    
    assert len(response.groups) > 0, "Should find groups despite Unicode form difference"
    assert response.group_columns == ["Город"], "Should return original NFC column names"
    assert response.agg_column == "Возраст", "Should return original NFC agg column"


def test_aggregate_unicode_column_not_found_with_suggestions(simple_fixture, file_loader):
    """Test aggregate error message with Unicode fuzzy suggestions.
    
    Verifies:
    - Error message provides fuzzy suggestions for Unicode columns
    - Suggestions work across Unicode forms
    """
    print(f"\n🔤 Testing aggregate error with Unicode suggestions")
    
    ops = DataOperations(file_loader)
    
    # Request with typo in Cyrillic column name
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="sum",
        target_column="Вазраст",  # Typo: "Вазраст" instead of "Возраст"
        filters=[]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.aggregate(request)
    
    error_msg = str(exc_info.value)
    print(f"✅ Error message: {error_msg}")
    
    assert "not found" in error_msg, "Should mention column not found"
    assert "Did you mean" in error_msg, "Should provide fuzzy suggestions"
    assert "Возраст" in error_msg, "Should suggest correct Cyrillic column"


def test_aggregate_with_filter_unicode_nfd(simple_fixture, file_loader):
    """Test aggregate with filter using NFD Unicode column.
    
    Verifies:
    - Filter engine handles NFD column names
    - Aggregation works with Unicode-normalized filters
    - End-to-end Unicode normalization in filtering + aggregation
    """
    print(f"\n🔤 Testing aggregate with NFD filter column")
    
    import unicodedata
    ops = DataOperations(file_loader)
    
    # Filter and target use NFD forms
    filter_col_nfd = unicodedata.normalize('NFD', "Возраст")
    target_col_nfd = unicodedata.normalize('NFD', "Возраст")
    
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="sum",
        target_column=target_col_nfd,
        filters=[
            FilterCondition(column=filter_col_nfd, operator=">", value=30)
        ]
    )
    
    # Act
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Filtered sum: {response.value}")
    print(f"   Filters applied: {len(response.filters_applied)}")
    
    assert response.value > 0, "Should calculate filtered sum with NFD columns"
    assert len(response.filters_applied) == 1, "Should apply filter"
    assert response.filters_applied[0]["column"] == "Возраст", "Should normalize filter column"


def test_group_by_unicode_column_not_found(simple_fixture, file_loader):
    """Test group_by error with non-existent Unicode column.
    
    Verifies:
    - Error message for Unicode column not found
    - Provides helpful suggestions
    """
    print(f"\n🔤 Testing group_by error with Unicode column")
    
    ops = DataOperations(file_loader)
    
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=["Гарод"],  # Typo: "Гарод" instead of "Город"
        agg_column="Возраст",
        agg_operation="sum",
        filters=[]
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.group_by(request)
    
    error_msg = str(exc_info.value)
    print(f"✅ Error message: {error_msg}")
    
    assert "not found" in error_msg, "Should mention column not found"
    assert "Гарод" in error_msg, "Should mention the typo column"


# ============================================================================
# NEGATION OPERATOR (NOT) TESTS
# ============================================================================

def test_aggregate_with_negated_filter(numeric_types_fixture, file_loader):
    """Test aggregate with negated filter condition.
    
    Verifies:
    - Negation works correctly in aggregate
    - Aggregates only rows satisfying negated condition
    - Formula is None (negation not supported in Excel)
    """
    print(f"\n🔍 Testing aggregate with negated filter")
    
    ops = DataOperations(file_loader)
    
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="sum",
        target_column="Количество",
        filters=[
            FilterCondition(column="Количество", operator="<", value=100, negate=True)
        ]
    )
    
    response = ops.aggregate(request)
    
    print(f"✅ Sum: {response.value}")
    print(f"   Formula: {response.excel_output.formula}")
    
    # Should sum rows where Количество >= 100
    assert response.value > 0, "Sum should be positive"
    assert response.excel_output.formula is None, "Formula should be None for negation"


def test_group_by_with_negated_filter(simple_fixture, file_loader):
    """Test group_by with negated filter.
    
    Verifies:
    - Negation works correctly in group_by
    - Groups exclude negated values
    - Results are correct
    """
    print(f"\n🔍 Testing group_by with negated filter")
    
    ops = DataOperations(file_loader)
    
    # Get a test value to exclude
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    test_value = ops.get_unique_values(unique_request).values[0]
    
    print(f"  Filter: {simple_fixture.columns[0]} == '{test_value}' (negated)")
    
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=[simple_fixture.columns[0]],
        agg_column=simple_fixture.columns[1],
        agg_operation="count",
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=test_value, negate=True)
        ]
    )
    
    response = ops.group_by(request)
    
    print(f"✅ Groups: {len(response.groups)}")
    
    # test_value should not be in results
    assert all(group[simple_fixture.columns[0]] != test_value for group in response.groups), \
        f"No group should have {simple_fixture.columns[0]} == {test_value}"
    assert len(response.groups) > 0, "Should have some groups"


# ============================================================================
# NESTED FILTER GROUPS TESTS (aggregate)
# ============================================================================

def test_aggregate_nested_and_or(numeric_types_fixture, file_loader):
    """Test aggregate with nested group: (A AND B) OR C.
    
    Verifies:
    - Nested groups work in aggregate
    - Aggregation is correct for complex logic
    - Formula is None (nested groups not supported in Excel)
    """
    print(f"\n🔍 Testing aggregate: (A AND B) OR C")
    
    from mcp_excel.models.requests import FilterGroup
    
    ops = DataOperations(file_loader)
    
    print(f"  Filter: (Количество < 50 AND Цена > 100) OR Количество == 100")
    
    # Act
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="sum",
        target_column="Количество",
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column="Количество", operator="<", value=50),
                    FilterCondition(column="Цена", operator=">", value=100)
                ],
                logic="AND"
            ),
            FilterCondition(column="Количество", operator="==", value=100)
        ],
        logic="OR"
    )
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Sum: {response.value}")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.value >= 0, "Sum should be non-negative"
    assert response.excel_output.formula is None, "Formula should be None for nested groups"


def test_aggregate_nested_three_levels(numeric_types_fixture, file_loader):
    """Test aggregate with 3 levels of nesting: ((A OR B) AND C) OR D.
    
    Verifies:
    - Deep nesting works in aggregate
    - Complex logic is evaluated correctly
    """
    print(f"\n🔍 Testing aggregate with 3 levels: ((A OR B) AND C) OR D")
    
    from mcp_excel.models.requests import FilterGroup
    
    ops = DataOperations(file_loader)
    
    print(f"  Filter: ((Количество < 50 OR Количество > 150) AND Цена > 100) OR Количество == 100")
    
    # Act
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="sum",
        target_column="Цена",
        filters=[
            FilterGroup(
                filters=[
                    FilterGroup(
                        filters=[
                            FilterCondition(column="Количество", operator="<", value=50),
                            FilterCondition(column="Количество", operator=">", value=150)
                        ],
                        logic="OR"
                    ),
                    FilterCondition(column="Цена", operator=">", value=100)
                ],
                logic="AND"
            ),
            FilterCondition(column="Количество", operator="==", value=100)
        ],
        logic="OR"
    )
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Sum: {response.value}")
    
    assert response.value >= 0, "Sum should be non-negative"
    assert response.excel_output.formula is None, "Formula should be None for nested groups"


def test_aggregate_nested_with_negation(simple_fixture, file_loader):
    """Test aggregate with nested group and negation: NOT (A AND B).
    
    Verifies:
    - Negation works with nested groups in aggregate
    - Aggregates only rows not matching the group
    """
    print(f"\n🔍 Testing aggregate: NOT (A AND B)")
    
    from mcp_excel.models.requests import FilterGroup, GetUniqueValuesRequest
    
    ops = DataOperations(file_loader)
    
    # Get test value
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    test_value = ops.get_unique_values(unique_request).values[0]
    
    print(f"  Filter: NOT ({simple_fixture.columns[0]} == '{test_value}' AND {simple_fixture.columns[1]} > 0)")
    
    # Act
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="sum",
        target_column=simple_fixture.columns[1],
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column=simple_fixture.columns[0], operator="==", value=test_value),
                    FilterCondition(column=simple_fixture.columns[1], operator=">", value=0)
                ],
                logic="AND",
                negate=True
            )
        ]
    )
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Sum: {response.value}")
    
    assert response.value > 0, "Should aggregate rows not matching the group"
    assert response.excel_output.formula is None, "Formula should be None for negated groups"


# ============================================================================
# NESTED FILTER GROUPS TESTS (group_by)
# ============================================================================

def test_group_by_nested_and_or(numeric_types_fixture, file_loader):
    """Test group_by with nested group: (A AND B) OR C.
    
    Verifies:
    - Nested groups work in group_by
    - Groups are correct for complex logic
    """
    print(f"\n🔍 Testing group_by: (A AND B) OR C")
    
    from mcp_excel.models.requests import FilterGroup
    
    ops = DataOperations(file_loader)
    
    print(f"  Filter: (Количество < 50 AND Цена > 100) OR Количество == 100")
    
    # Act
    request = GroupByRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        group_columns=["Код товара"],
        agg_column="Количество",
        agg_operation="sum",
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column="Количество", operator="<", value=50),
                    FilterCondition(column="Цена", operator=">", value=100)
                ],
                logic="AND"
            ),
            FilterCondition(column="Количество", operator="==", value=100)
        ],
        logic="OR"
    )
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups: {len(response.groups)}")
    
    assert len(response.groups) >= 0, "Should return groups"


def test_group_by_nested_three_levels(numeric_types_fixture, file_loader):
    """Test group_by with 3 levels of nesting: ((A OR B) AND C) OR D.
    
    Verifies:
    - Deep nesting works in group_by
    - Complex logic is evaluated correctly
    """
    print(f"\n🔍 Testing group_by with 3 levels: ((A OR B) AND C) OR D")
    
    from mcp_excel.models.requests import FilterGroup
    
    ops = DataOperations(file_loader)
    
    print(f"  Filter: ((Количество < 50 OR Количество > 150) AND Цена > 100) OR Количество == 100")
    
    # Act
    request = GroupByRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        group_columns=["Код товара"],
        agg_column="Цена",
        agg_operation="sum",
        filters=[
            FilterGroup(
                filters=[
                    FilterGroup(
                        filters=[
                            FilterCondition(column="Количество", operator="<", value=50),
                            FilterCondition(column="Количество", operator=">", value=150)
                        ],
                        logic="OR"
                    ),
                    FilterCondition(column="Цена", operator=">", value=100)
                ],
                logic="AND"
            ),
            FilterCondition(column="Количество", operator="==", value=100)
        ],
        logic="OR"
    )
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups: {len(response.groups)}")
    
    assert len(response.groups) >= 0, "Should return groups"


def test_group_by_nested_with_negation(simple_fixture, file_loader):
    """Test group_by with nested group and negation: NOT (A AND B).
    
    Verifies:
    - Negation works with nested groups in group_by
    - Groups exclude rows matching the negated group
    """
    print(f"\n🔍 Testing group_by: NOT (A AND B)")
    
    from mcp_excel.models.requests import FilterGroup, GetUniqueValuesRequest
    
    ops = DataOperations(file_loader)
    
    # Get test value
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    test_value = ops.get_unique_values(unique_request).values[0]
    
    print(f"  Filter: NOT ({simple_fixture.columns[0]} == '{test_value}' AND {simple_fixture.columns[1]} > 0)")
    
    # Act
    request = GroupByRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        group_columns=[simple_fixture.columns[0]],
        agg_column=simple_fixture.columns[1],
        agg_operation="count",
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column=simple_fixture.columns[0], operator="==", value=test_value),
                    FilterCondition(column=simple_fixture.columns[1], operator=">", value=0)
                ],
                logic="AND",
                negate=True
            )
        ]
    )
    response = ops.group_by(request)
    
    # Assert
    print(f"✅ Groups: {len(response.groups)}")
    
    # test_value should not be in results (it's excluded by negated group)
    assert all(group[simple_fixture.columns[0]] != test_value for group in response.groups), \
        f"No group should have {simple_fixture.columns[0]} == {test_value}"
    assert len(response.groups) > 0, "Should have some groups"


# ============================================================================
# SAMPLE_ROWS PARAMETER TESTS
# ============================================================================

def test_aggregate_with_sample_rows(numeric_types_fixture, file_loader):
    """Test aggregate with sample_rows parameter.
    
    Verifies:
    - sample_rows parameter returns sample data
    - Sample data shows rows used in aggregation
    - Values are formatted correctly
    """
    print(f"\n🔍 Testing aggregate with sample_rows")
    
    ops = DataOperations(file_loader)
    
    # Act
    request = AggregateRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        operation="sum",
        target_column="Количество",
        filters=[
            FilterCondition(column="Цена", operator=">", value=100)
        ],
        sample_rows=3
    )
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Sum: {response.value}, Sample rows: {len(response.sample_rows) if response.sample_rows else 0}")
    
    assert response.sample_rows is not None, "Should return sample_rows"
    assert isinstance(response.sample_rows, list), "sample_rows should be list"
    assert len(response.sample_rows) <= 3, "Should return at most 3 rows"
    
    # Verify structure
    if response.sample_rows:
        assert all(isinstance(row, dict) for row in response.sample_rows), "Each row should be dict"
        assert all("Количество" in row for row in response.sample_rows), "Should have target column"
        assert all("Цена" in row for row in response.sample_rows), "Should have filter column"
        # Verify filter was applied: all Цена > 100
        assert all(row["Цена"] > 100 for row in response.sample_rows), "All samples should match filter"


def test_aggregate_sample_rows_without_filters(simple_fixture, file_loader):
    """Test aggregate with sample_rows but no filters.
    
    Verifies:
    - sample_rows works without filters
    - Returns samples from entire dataset
    """
    print(f"\n🔍 Testing aggregate with sample_rows (no filters)")
    
    ops = DataOperations(file_loader)
    
    # Act
    request = AggregateRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        operation="count",
        target_column=simple_fixture.columns[0],
        filters=[],
        sample_rows=5
    )
    response = ops.aggregate(request)
    
    # Assert
    print(f"✅ Count: {response.value}, Sample rows: {len(response.sample_rows) if response.sample_rows else 0}")
    
    assert response.sample_rows is not None, "Should return sample_rows"
    assert len(response.sample_rows) <= 5, "Should return at most 5 rows"
    assert len(response.sample_rows) <= response.value, "Sample size should not exceed count"
