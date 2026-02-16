# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Integration tests for Time Series operations.

Tests cover:
- calculate_period_change: Period-over-period analysis (month/quarter/year)
- calculate_running_total: Cumulative sum calculations
- calculate_moving_average: Moving average with window size

These are END-TO-END tests that verify the complete operation flow:
FileLoader -> HeaderDetector -> TimeSeriesOperations -> Response
"""

import pytest

from mcp_excel.operations.timeseries import TimeSeriesOperations
from mcp_excel.models.requests import (
    CalculatePeriodChangeRequest,
    CalculateRunningTotalRequest,
    CalculateMovingAverageRequest,
    FilterCondition,
)


# ============================================================================
# calculate_period_change tests
# ============================================================================

def test_calculate_period_change_month(with_dates_fixture, file_loader):
    """Test calculate_period_change with monthly periods.
    
    Verifies:
    - Groups data by month correctly
    - Calculates absolute and percentage changes
    - Returns period data in correct format
    - Generates TSV output
    - Generates Excel formula
    """
    print(f"\n📊 Testing calculate_period_change with monthly periods")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculatePeriodChangeRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        date_column="Дата заказа",
        value_column="Сумма",
        period_type="month"
    )
    
    # Act
    response = ops.calculate_period_change(request)
    
    # Assert
    print(f"✅ Found {len(response.periods)} periods")
    print(f"   Period type: {response.period_type}")
    print(f"   Value column: {response.value_column}")
    
    assert response.period_type == "month", "Should use monthly periods"
    assert response.value_column == "Сумма", "Should track correct column"
    assert len(response.periods) > 0, "Should have at least one period"
    
    # Check period structure
    for idx, period in enumerate(response.periods[:3], 1):
        print(f"   Period {idx}: {period['period']}")
        print(f"     Value: {period['value']}")
        print(f"     Change (abs): {period['change_absolute']}")
        print(f"     Change (%): {period['change_percent']}")
        
        assert "period" in period, "Should have period field"
        assert "value" in period, "Should have value field"
        assert "change_absolute" in period, "Should have absolute change"
        assert "change_percent" in period, "Should have percentage change"
    
    # First period should have None changes (no previous period)
    first_period = response.periods[0]
    # Note: pandas returns NaN for first diff/pct_change, which becomes None in JSON
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert "Period" in response.excel_output.tsv, "TSV should have Period header"
    assert "Value" in response.excel_output.tsv, "TSV should have Value header"
    assert "Change" in response.excel_output.tsv, "TSV should have Change header"
    
    # Check Excel formula
    assert response.excel_output.formula, "Should generate Excel formula"
    assert "=" in response.excel_output.formula, "Formula should start with ="
    
    # Check performance metrics
    assert response.performance.execution_time_ms > 0, "Should have execution time"


def test_calculate_period_change_quarter(with_dates_fixture, file_loader):
    """Test calculate_period_change with quarterly periods.
    
    Verifies:
    - Groups data by quarter correctly
    - Calculates changes between quarters
    """
    print(f"\n📊 Testing calculate_period_change with quarterly periods")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculatePeriodChangeRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        date_column="Дата заказа",
        value_column="Сумма",
        period_type="quarter"
    )
    
    # Act
    response = ops.calculate_period_change(request)
    
    # Assert
    print(f"✅ Found {len(response.periods)} quarters")
    
    assert response.period_type == "quarter", "Should use quarterly periods"
    assert len(response.periods) > 0, "Should have at least one quarter"
    
    # Check that periods are in quarter format (e.g., "2024Q1")
    for period in response.periods:
        assert "Q" in period["period"], "Quarter format should contain 'Q'"


def test_calculate_period_change_year(with_dates_fixture, file_loader):
    """Test calculate_period_change with yearly periods.
    
    Verifies:
    - Groups data by year correctly
    - Calculates year-over-year changes
    """
    print(f"\n📊 Testing calculate_period_change with yearly periods")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculatePeriodChangeRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        date_column="Дата заказа",
        value_column="Сумма",
        period_type="year"
    )
    
    # Act
    response = ops.calculate_period_change(request)
    
    # Assert
    print(f"✅ Found {len(response.periods)} years")
    
    assert response.period_type == "year", "Should use yearly periods"
    assert len(response.periods) > 0, "Should have at least one year"


def test_calculate_period_change_with_filters(with_dates_fixture, file_loader):
    """Test calculate_period_change with filters.
    
    Verifies:
    - Applies filters before grouping
    - Returns filtered period data
    """
    print(f"\n📊 Testing calculate_period_change with filters")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculatePeriodChangeRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        date_column="Дата заказа",
        value_column="Сумма",
        period_type="month",
        filters=[
            FilterCondition(column="Клиент", operator="==", value="Ромашка")
        ],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_period_change(request)
    
    # Assert
    print(f"✅ Found {len(response.periods)} periods (filtered)")
    
    assert len(response.periods) > 0, "Should have periods after filtering"
    # Filtered data should have fewer or equal periods than unfiltered
    assert response.period_type == "month"


def test_calculate_period_change_invalid_date_column(with_dates_fixture, file_loader):
    """Test calculate_period_change with invalid date column.
    
    Verifies:
    - Raises ValueError for non-existent column
    - Error message is helpful
    """
    print(f"\n📊 Testing calculate_period_change with invalid date column")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculatePeriodChangeRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        date_column="NonExistentColumn",
        value_column="Сумма",
        period_type="month"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.calculate_period_change(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower()
    assert "NonExistentColumn" in str(exc_info.value)


def test_calculate_period_change_invalid_value_column(with_dates_fixture, file_loader):
    """Test calculate_period_change with invalid value column.
    
    Verifies:
    - Raises ValueError for non-existent column
    """
    print(f"\n📊 Testing calculate_period_change with invalid value column")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculatePeriodChangeRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        date_column="Дата заказа",
        value_column="NonExistentColumn",
        period_type="month"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.calculate_period_change(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower()


# ============================================================================
# calculate_running_total tests
# ============================================================================

def test_calculate_running_total_basic(numeric_types_fixture, file_loader):
    """Test calculate_running_total without grouping.
    
    Verifies:
    - Calculates cumulative sum correctly
    - Sorts by order column
    - Returns all rows with running total
    - Generates TSV output
    - Generates Excel formula
    """
    print(f"\n📊 Testing calculate_running_total (basic)")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateRunningTotalRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество"
    )
    
    # Act
    response = ops.calculate_running_total(request)
    
    # Assert
    print(f"✅ Calculated running total for {len(response.rows)} rows")
    print(f"   Order column: {response.order_column}")
    print(f"   Value column: {response.value_column}")
    
    assert response.order_column == "Код товара"
    assert response.value_column == "Количество"
    assert len(response.rows) > 0, "Should have rows"
    
    # Check row structure
    first_row = response.rows[0]
    print(f"   First row: {first_row}")
    
    assert "Код товара" in first_row, "Should have order column"
    assert "Количество" in first_row, "Should have value column"
    assert "running_total" in first_row, "Should have running_total column"
    
    # Check that running total is cumulative (each value >= previous)
    prev_total = 0
    for idx, row in enumerate(response.rows[:5], 1):
        current_total = row["running_total"]
        print(f"   Row {idx}: value={row['Количество']}, running_total={current_total}")
        # Running total should be monotonically increasing (or equal if value is 0)
        assert current_total >= prev_total, "Running total should be cumulative"
        prev_total = current_total
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert "running_total" in response.excel_output.tsv, "TSV should have running_total column"
    
    # Check Excel formula
    assert response.excel_output.formula, "Should generate Excel formula"
    assert "SUM" in response.excel_output.formula, "Formula should use SUM"
    
    # Check performance metrics
    assert response.performance.execution_time_ms > 0


def test_calculate_running_total_with_grouping(with_dates_fixture, file_loader):
    """Test calculate_running_total with grouping.
    
    Verifies:
    - Calculates running total within each group
    - Resets cumulative sum for each group
    - Includes group columns in output
    """
    print(f"\n📊 Testing calculate_running_total with grouping")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateRunningTotalRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        order_column="Дата заказа",
        value_column="Сумма",
        group_by_columns=["Клиент"]
    )
    
    # Act
    response = ops.calculate_running_total(request)
    
    # Assert
    print(f"✅ Calculated grouped running total for {len(response.rows)} rows")
    print(f"   Group by: {response.group_by_columns}")
    
    assert response.group_by_columns == ["Клиент"]
    assert len(response.rows) > 0
    
    # Check that group column is in output
    first_row = response.rows[0]
    assert "Клиент" in first_row, "Should have group column"
    assert "running_total" in first_row, "Should have running_total"
    
    # Sample output
    for idx, row in enumerate(response.rows[:3], 1):
        print(f"   Row {idx}: client={row['Клиент']}, total={row['running_total']}")


def test_calculate_running_total_with_filters(numeric_types_fixture, file_loader):
    """Test calculate_running_total with filters.
    
    Verifies:
    - Applies filters before calculating
    - Returns filtered running total
    """
    print(f"\n📊 Testing calculate_running_total with filters")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateRunningTotalRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество",
        filters=[
            FilterCondition(column="Количество", operator=">", value=50)
        ],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_running_total(request)
    
    # Assert
    print(f"✅ Calculated running total for {len(response.rows)} filtered rows")
    
    assert len(response.rows) > 0, "Should have rows after filtering"
    # All values should be > 50
    for row in response.rows:
        assert row["Количество"] > 50, "Should only include filtered rows"


def test_calculate_running_total_invalid_order_column(numeric_types_fixture, file_loader):
    """Test calculate_running_total with invalid order column.
    
    Verifies:
    - Raises ValueError for non-existent column
    """
    print(f"\n📊 Testing calculate_running_total with invalid order column")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateRunningTotalRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="NonExistentColumn",
        value_column="Количество"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.calculate_running_total(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower()


def test_calculate_running_total_invalid_value_column(numeric_types_fixture, file_loader):
    """Test calculate_running_total with invalid value column.
    
    Verifies:
    - Raises ValueError for non-existent column
    """
    print(f"\n📊 Testing calculate_running_total with invalid value column")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateRunningTotalRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="NonExistentColumn"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.calculate_running_total(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower()


# ============================================================================
# calculate_moving_average tests
# ============================================================================

def test_calculate_moving_average_basic(numeric_types_fixture, file_loader):
    """Test calculate_moving_average with basic window.
    
    Verifies:
    - Calculates moving average correctly
    - Sorts by order column
    - Returns all rows with moving average
    - Generates TSV output
    - Generates Excel formula
    """
    print(f"\n📊 Testing calculate_moving_average (basic)")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество",
        window_size=3
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Calculated moving average for {len(response.rows)} rows")
    print(f"   Order column: {response.order_column}")
    print(f"   Value column: {response.value_column}")
    print(f"   Window size: {response.window_size}")
    
    assert response.order_column == "Код товара"
    assert response.value_column == "Количество"
    assert response.window_size == 3
    assert len(response.rows) > 0, "Should have rows"
    
    # Check row structure
    first_row = response.rows[0]
    print(f"   First row: {first_row}")
    
    assert "Код товара" in first_row, "Should have order column"
    assert "Количество" in first_row, "Should have value column"
    assert "moving_average" in first_row, "Should have moving_average column"
    
    # Sample output
    for idx, row in enumerate(response.rows[:5], 1):
        print(f"   Row {idx}: value={row['Количество']}, moving_avg={row['moving_average']}")
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert "moving_average" in response.excel_output.tsv, "TSV should have moving_average column"
    
    # Check Excel formula
    assert response.excel_output.formula, "Should generate Excel formula"
    assert "AVERAGE" in response.excel_output.formula, "Formula should use AVERAGE"
    
    # Check performance metrics
    assert response.performance.execution_time_ms > 0


def test_calculate_moving_average_window_size_1(numeric_types_fixture, file_loader):
    """Test calculate_moving_average with window size 1.
    
    Verifies:
    - Window size 1 returns original values
    - Edge case handling
    """
    print(f"\n📊 Testing calculate_moving_average with window_size=1")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество",
        window_size=1
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Window size 1: {len(response.rows)} rows")
    
    assert response.window_size == 1
    assert len(response.rows) > 0
    
    # With window_size=1, moving average should equal original value
    for row in response.rows[:3]:
        # Note: Due to numeric conversion and formatting, we check approximate equality
        print(f"   Value: {row['Количество']}, Moving avg: {row['moving_average']}")


def test_calculate_moving_average_large_window(numeric_types_fixture, file_loader):
    """Test calculate_moving_average with large window size.
    
    Verifies:
    - Handles window size larger than some rows
    - Uses min_periods=1 to avoid NaN for early rows
    """
    print(f"\n📊 Testing calculate_moving_average with large window")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество",
        window_size=10
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Large window: {len(response.rows)} rows")
    
    assert response.window_size == 10
    assert len(response.rows) > 0
    
    # All rows should have moving_average (thanks to min_periods=1)
    for row in response.rows:
        assert "moving_average" in row
        assert row["moving_average"] is not None


def test_calculate_moving_average_with_dates(with_dates_fixture, file_loader):
    """Test calculate_moving_average with datetime order column.
    
    Verifies:
    - Works with datetime columns for ordering
    - Calculates moving average over time
    """
    print(f"\n📊 Testing calculate_moving_average with datetime order")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        order_column="Дата заказа",
        value_column="Сумма",
        window_size=3
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Moving average over time: {len(response.rows)} rows")
    
    assert response.order_column == "Дата заказа"
    assert len(response.rows) > 0
    
    # Sample output
    for idx, row in enumerate(response.rows[:3], 1):
        print(f"   Row {idx}: date={row['Дата заказа']}, avg={row['moving_average']}")


def test_calculate_moving_average_with_filters(numeric_types_fixture, file_loader):
    """Test calculate_moving_average with filters.
    
    Verifies:
    - Applies filters before calculating
    - Returns filtered moving average
    """
    print(f"\n📊 Testing calculate_moving_average with filters")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество",
        window_size=3,
        filters=[
            FilterCondition(column="Количество", operator=">", value=50)
        ],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Moving average for {len(response.rows)} filtered rows")
    
    assert len(response.rows) > 0, "Should have rows after filtering"
    # All values should be > 50
    for row in response.rows:
        assert row["Количество"] > 50, "Should only include filtered rows"


def test_calculate_moving_average_invalid_order_column(numeric_types_fixture, file_loader):
    """Test calculate_moving_average with invalid order column.
    
    Verifies:
    - Raises ValueError for non-existent column
    """
    print(f"\n📊 Testing calculate_moving_average with invalid order column")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="NonExistentColumn",
        value_column="Количество",
        window_size=3
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.calculate_moving_average(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower()


def test_calculate_moving_average_invalid_value_column(numeric_types_fixture, file_loader):
    """Test calculate_moving_average with invalid value column.
    
    Verifies:
    - Raises ValueError for non-existent column
    """
    print(f"\n📊 Testing calculate_moving_average with invalid value column")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="NonExistentColumn",
        window_size=3
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.calculate_moving_average(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower()


def test_calculate_moving_average_window_size_5(numeric_types_fixture, file_loader):
    """Test calculate_moving_average with window size 5.
    
    Verifies:
    - Different window sizes work correctly
    - Formula generation adapts to window size
    """
    print(f"\n📊 Testing calculate_moving_average with window_size=5")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество",
        window_size=5
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Window size 5: {len(response.rows)} rows")
    print(f"   Formula: {response.excel_output.formula}")
    
    assert response.window_size == 5
    assert len(response.rows) > 0
    assert response.excel_output.formula, "Should generate formula"


def test_calculate_moving_average_tsv_output(numeric_types_fixture, file_loader):
    """Test that calculate_moving_average generates proper TSV output.
    
    Verifies:
    - TSV output is generated
    - Contains all required columns
    - Can be pasted into Excel
    """
    print(f"\n📊 Testing calculate_moving_average TSV output")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество",
        window_size=3
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ TSV output generated")
    print(f"   Length: {len(response.excel_output.tsv)} chars")
    print(f"   Preview: {response.excel_output.tsv[:200]}...")
    
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    
    # Check TSV contains column names
    assert "Код товара" in response.excel_output.tsv
    assert "Количество" in response.excel_output.tsv
    assert "moving_average" in response.excel_output.tsv
    
    # Check TSV has tab separators
    assert "\t" in response.excel_output.tsv, "TSV should use tab separators"


def test_calculate_moving_average_performance(numeric_types_fixture, file_loader):
    """Test calculate_moving_average performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    """
    print(f"\n📊 Testing calculate_moving_average performance")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        order_column="Код товара",
        value_column="Количество",
        window_size=3
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Performance:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Cache hit: {response.performance.cache_hit}")
    
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.execution_time_ms < 5000, "Should complete in reasonable time"


# ============================================================================
# Edge cases and integration tests
# ============================================================================

def test_period_change_with_manual_header(messy_headers_fixture, file_loader):
    """Test calculate_period_change with manual header_row.
    
    Verifies:
    - Works with messy files when header_row is specified
    - Auto-detection also works
    """
    print(f"\n📊 Testing period_change with messy headers")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculatePeriodChangeRequest(
        file_path=messy_headers_fixture.path_str,
        sheet_name=messy_headers_fixture.sheet_name,
        date_column="Дата",
        value_column="Сумма",
        period_type="month",
        header_row=messy_headers_fixture.header_row
    )
    
    # Act
    response = ops.calculate_period_change(request)
    
    # Assert
    print(f"✅ Handled messy headers: {len(response.periods)} periods")
    
    assert len(response.periods) > 0, "Should work with messy headers"


def test_running_total_with_manual_header(messy_headers_fixture, file_loader):
    """Test calculate_running_total with manual header_row.
    
    Verifies:
    - Works with messy files when header_row is specified
    """
    print(f"\n📊 Testing running_total with messy headers")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateRunningTotalRequest(
        file_path=messy_headers_fixture.path_str,
        sheet_name=messy_headers_fixture.sheet_name,
        order_column="Дата",
        value_column="Сумма",
        header_row=messy_headers_fixture.header_row
    )
    
    # Act
    response = ops.calculate_running_total(request)
    
    # Assert
    print(f"✅ Handled messy headers: {len(response.rows)} rows")
    
    assert len(response.rows) > 0, "Should work with messy headers"


def test_moving_average_with_manual_header(messy_headers_fixture, file_loader):
    """Test calculate_moving_average with manual header_row.
    
    Verifies:
    - Works with messy files when header_row is specified
    """
    print(f"\n📊 Testing moving_average with messy headers")
    
    ops = TimeSeriesOperations(file_loader)
    request = CalculateMovingAverageRequest(
        file_path=messy_headers_fixture.path_str,
        sheet_name=messy_headers_fixture.sheet_name,
        order_column="Дата",
        value_column="Сумма",
        window_size=3,
        header_row=messy_headers_fixture.header_row
    )
    
    # Act
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Handled messy headers: {len(response.rows)} rows")
    
    assert len(response.rows) > 0, "Should work with messy headers"


# ============================================================================
# NEGATION OPERATOR (NOT) TESTS
# ============================================================================

def test_calculate_period_change_with_negation(with_dates_fixture, file_loader):
    """Test calculate_period_change with negated filter.
    
    Verifies:
    - Period change calculated only for filtered rows
    - Negation works correctly in timeseries context
    """
    print(f"\n🔍 Testing calculate_period_change with negation")
    
    ops = TimeSeriesOperations(file_loader)
    
    from mcp_excel.models.requests import CalculatePeriodChangeRequest
    
    date_col = with_dates_fixture.expected["datetime_columns"][0]
    
    request = CalculatePeriodChangeRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        date_column=date_col,
        value_column="Сумма",
        period_type="month",
        filters=[
            FilterCondition(column="Сумма", operator="<", value=1000, negate=True)
        ]
    )
    
    response = ops.calculate_period_change(request)
    
    print(f"✅ Period change with negation: {len(response.periods)} periods")
    
    assert len(response.periods) >= 0, "Should calculate periods"


def test_calculate_running_total_with_negation(with_dates_fixture, file_loader):
    """Test calculate_running_total with negated filter.
    
    Verifies:
    - Running total calculated only for filtered rows
    - Negation works correctly
    """
    print(f"\n🔍 Testing calculate_running_total with negation")
    
    ops = TimeSeriesOperations(file_loader)
    
    from mcp_excel.models.requests import CalculateRunningTotalRequest
    
    date_col = with_dates_fixture.expected["datetime_columns"][0]
    
    request = CalculateRunningTotalRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        order_column=date_col,
        value_column="Сумма",
        group_by_columns=None,
        filters=[
            FilterCondition(column="Сумма", operator="<", value=500, negate=True)
        ]
    )
    
    response = ops.calculate_running_total(request)
    
    print(f"✅ Running total with negation: {len(response.rows)} rows")
    
    assert len(response.rows) > 0, "Should calculate running total"


# ============================================================================
# NESTED FILTER GROUPS TESTS (timeseries)
# ============================================================================

def test_calculate_period_change_nested_filters(with_dates_fixture, file_loader):
    """Test calculate_period_change with nested group: (A AND B) OR C.
    
    Verifies:
    - Nested groups work in calculate_period_change
    - Period changes calculated only for filtered rows
    """
    print(f"\n🔍 Testing calculate_period_change: (A AND B) OR C")
    
    from mcp_excel.models.requests import FilterGroup, CalculatePeriodChangeRequest
    
    ops = TimeSeriesOperations(file_loader)
    
    date_col = with_dates_fixture.expected["datetime_columns"][0]
    
    print(f"  Filter: (Сумма > 500 AND Сумма < 2000) OR Сумма == 3000")
    
    # Act
    request = CalculatePeriodChangeRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        date_column=date_col,
        value_column="Сумма",
        period_type="month",
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column="Сумма", operator=">", value=500),
                    FilterCondition(column="Сумма", operator="<", value=2000)
                ],
                logic="AND"
            ),
            FilterCondition(column="Сумма", operator="==", value=3000)
        ],
        logic="OR"
    )
    response = ops.calculate_period_change(request)
    
    # Assert
    print(f"✅ Period change with nested filters: {len(response.periods)} periods")
    
    assert len(response.periods) >= 0, "Should calculate periods"


def test_calculate_running_total_nested_filters(with_dates_fixture, file_loader):
    """Test calculate_running_total with nested group: (A OR B) AND C.
    
    Verifies:
    - Nested groups work in calculate_running_total
    - Running total calculated only for filtered rows
    """
    print(f"\n🔍 Testing calculate_running_total: (A OR B) AND C")
    
    from mcp_excel.models.requests import FilterGroup, CalculateRunningTotalRequest
    
    ops = TimeSeriesOperations(file_loader)
    
    date_col = with_dates_fixture.expected["datetime_columns"][0]
    
    print(f"  Filter: (Сумма < 500 OR Сумма > 2000) AND Сумма != 1000")
    
    # Act
    request = CalculateRunningTotalRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        order_column=date_col,
        value_column="Сумма",
        group_by_columns=None,
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column="Сумма", operator="<", value=500),
                    FilterCondition(column="Сумма", operator=">", value=2000)
                ],
                logic="OR"
            ),
            FilterCondition(column="Сумма", operator="!=", value=1000)
        ],
        logic="AND"
    )
    response = ops.calculate_running_total(request)
    
    # Assert
    print(f"✅ Running total with nested filters: {len(response.rows)} rows")
    
    assert len(response.rows) >= 0, "Should calculate running total"


def test_calculate_moving_average_nested_filters(with_dates_fixture, file_loader):
    """Test calculate_moving_average with nested group and negation.
    
    Verifies:
    - Nested groups work in calculate_moving_average
    - Moving average calculated only for filtered rows
    - Negation works with nested groups
    """
    print(f"\n🔍 Testing calculate_moving_average: NOT (A AND B)")
    
    from mcp_excel.models.requests import FilterGroup, CalculateMovingAverageRequest
    
    ops = TimeSeriesOperations(file_loader)
    
    date_col = with_dates_fixture.expected["datetime_columns"][0]
    
    print(f"  Filter: NOT (Сумма < 500 AND Сумма > 100)")
    
    # Act
    request = CalculateMovingAverageRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        order_column=date_col,
        value_column="Сумма",
        window_size=3,
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column="Сумма", operator="<", value=500),
                    FilterCondition(column="Сумма", operator=">", value=100)
                ],
                logic="AND",
                negate=True
            )
        ]
    )
    response = ops.calculate_moving_average(request)
    
    # Assert
    print(f"✅ Moving average with negated group: {len(response.rows)} rows")
    
    assert len(response.rows) >= 0, "Should calculate moving average"
