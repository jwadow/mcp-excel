# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Smoke tests for time series tools.

Tests for 3 time series tools:
- calculate_period_change
- calculate_running_total
- calculate_moving_average

Each tool is tested with:
- Basic functionality (happy path)
- Different parameter combinations
- With filters
- Full response structure validation
- Excel output validation
"""

import pytest


# ============================================================================
# CALCULATE_PERIOD_CHANGE TESTS
# ============================================================================

def test_calculate_period_change_month(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_period_change with month period type."""
    print(f"\n📈 Testing calculate_period_change (month)...")
    
    result = mcp_call_tool("calculate_period_change", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "date_column": "Дата заказа",
        "value_column": "Сумма",
        "period_type": "month"
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from CalculatePeriodChangeResponse
    assert "periods" in result, "Missing 'periods'"
    assert "period_type" in result, "Missing 'period_type'"
    assert "value_column" in result, "Missing 'value_column'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify period_type
    assert result["period_type"] == "month", f"Expected period_type='month', got {result['period_type']}"
    
    # Verify value_column
    assert result["value_column"] == "Сумма", f"Expected value_column='Сумма', got {result['value_column']}"
    
    # Verify periods structure
    periods = result["periods"]
    assert isinstance(periods, list), "periods should be list"
    
    if len(periods) > 0:
        print(f"  Found {len(periods)} periods")
        
        # Verify each period structure (based on integration test)
        for i, period in enumerate(periods):
            assert isinstance(period, dict), f"Period {i} should be dict"
            assert "period" in period, f"Period {i} missing 'period'"
            assert "value" in period, f"Period {i} missing 'value'"
            assert "change_absolute" in period, f"Period {i} missing 'change_absolute'"
            assert "change_percent" in period, f"Period {i} missing 'change_percent'"
    else:
        print(f"  ℹ️  No periods found (file might not have enough data)")
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    
    # Verify metadata
    metadata = result["metadata"]
    assert metadata["sheet_name"] == with_dates_fixture.sheet_name, "Metadata sheet_name mismatch"
    
    # Verify performance
    performance = result["performance"]
    assert performance["execution_time_ms"] >= 0, "Execution time should be non-negative"
    
    print(f"  ✅ Period change calculated successfully")


def test_calculate_period_change_quarter(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_period_change with quarter period type."""
    print(f"\n📈 Testing calculate_period_change (quarter)...")
    
    result = mcp_call_tool("calculate_period_change", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "date_column": "Дата заказа",
        "value_column": "Сумма",
        "period_type": "quarter"
    })
    
    # Verify period_type
    assert result["period_type"] == "quarter", f"Expected period_type='quarter', got {result['period_type']}"
    
    # Verify response structure
    assert "periods" in result
    assert "excel_output" in result
    assert "metadata" in result
    assert "performance" in result
    
    print(f"  ✅ Quarter period change works")


def test_calculate_period_change_year(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_period_change with year period type."""
    print(f"\n📈 Testing calculate_period_change (year)...")
    
    result = mcp_call_tool("calculate_period_change", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "date_column": "Дата заказа",
        "value_column": "Сумма",
        "period_type": "year"
    })
    
    # Verify period_type
    assert result["period_type"] == "year", f"Expected period_type='year', got {result['period_type']}"
    
    # Verify response structure
    assert "periods" in result
    assert "excel_output" in result
    
    print(f"  ✅ Year period change works")


def test_calculate_period_change_with_filter(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_period_change with filters applied."""
    print(f"\n📈 Testing calculate_period_change with filter...")
    
    result = mcp_call_tool("calculate_period_change", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "date_column": "Дата заказа",
        "value_column": "Сумма",
        "period_type": "month",
        "filters": [
            {"column": "Сумма", "operator": ">", "value": 0}
        ]
    })
    
    # Should work with filters
    assert "periods" in result
    assert "period_type" in result
    assert result["period_type"] == "month"
    
    print(f"  ✅ Period change with filter works")


# ============================================================================
# CALCULATE_RUNNING_TOTAL TESTS
# ============================================================================

def test_calculate_running_total_basic(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_running_total returns complete running total data."""
    print(f"\n📊 Testing calculate_running_total (basic)...")
    
    result = mcp_call_tool("calculate_running_total", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "order_column": "Дата заказа",
        "value_column": "Сумма"
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from CalculateRunningTotalResponse
    assert "rows" in result, "Missing 'rows'"
    assert "order_column" in result, "Missing 'order_column'"
    assert "value_column" in result, "Missing 'value_column'"
    assert "group_by_columns" in result, "Missing 'group_by_columns'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify order_column and value_column
    assert result["order_column"] == "Дата заказа", f"Expected order_column='Дата заказа', got {result['order_column']}"
    assert result["value_column"] == "Сумма", f"Expected value_column='Сумма', got {result['value_column']}"
    
    # Verify group_by_columns (should be None for basic test)
    assert result["group_by_columns"] is None or result["group_by_columns"] == [], "group_by_columns should be None or empty for basic test"
    
    # Verify rows structure
    rows = result["rows"]
    assert isinstance(rows, list), "rows should be list"
    
    if len(rows) > 0:
        print(f"  Found {len(rows)} rows with running totals")
        
        # Verify each row has running_total
        for i, row in enumerate(rows[:3]):  # Check first 3 rows
            assert isinstance(row, dict), f"Row {i} should be dict"
            assert "running_total" in row, f"Row {i} missing 'running_total'"
            assert isinstance(row["running_total"], (int, float)), f"Row {i} running_total should be numeric"
    else:
        print(f"  ℹ️  No rows returned (file might be empty)")
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    
    # Verify metadata
    metadata = result["metadata"]
    assert metadata["sheet_name"] == with_dates_fixture.sheet_name, "Metadata sheet_name mismatch"
    
    print(f"  ✅ Running total calculated successfully")


def test_calculate_running_total_with_grouping(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_running_total with grouping (running total per group)."""
    print(f"\n📊 Testing calculate_running_total with grouping...")
    
    result = mcp_call_tool("calculate_running_total", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "order_column": "Дата заказа",
        "value_column": "Сумма",
        "group_by_columns": ["Клиент"]
    })
    
    # Verify group_by_columns
    assert result["group_by_columns"] is not None, "group_by_columns should not be None"
    assert isinstance(result["group_by_columns"], list), "group_by_columns should be list"
    assert "Клиент" in result["group_by_columns"], "group_by_columns should contain 'Клиент'"
    
    # Verify rows structure
    assert "rows" in result
    assert isinstance(result["rows"], list), "rows should be list"
    
    # Verify excel_output
    assert "excel_output" in result
    
    print(f"  ✅ Running total with grouping works")


def test_calculate_running_total_with_filter(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_running_total with filters applied."""
    print(f"\n📊 Testing calculate_running_total with filter...")
    
    result = mcp_call_tool("calculate_running_total", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "order_column": "Дата заказа",
        "value_column": "Сумма",
        "filters": [
            {"column": "Сумма", "operator": ">", "value": 0}
        ]
    })
    
    # Should work with filters
    assert "rows" in result
    assert "order_column" in result
    assert result["order_column"] == "Дата заказа"
    
    print(f"  ✅ Running total with filter works")


# ============================================================================
# CALCULATE_MOVING_AVERAGE TESTS
# ============================================================================

def test_calculate_moving_average_basic(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_moving_average returns complete moving average data."""
    print(f"\n📉 Testing calculate_moving_average (window=3)...")
    
    result = mcp_call_tool("calculate_moving_average", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "order_column": "Дата заказа",
        "value_column": "Сумма",
        "window_size": 3
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from CalculateMovingAverageResponse
    assert "rows" in result, "Missing 'rows'"
    assert "order_column" in result, "Missing 'order_column'"
    assert "value_column" in result, "Missing 'value_column'"
    assert "window_size" in result, "Missing 'window_size'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify order_column, value_column, window_size
    assert result["order_column"] == "Дата заказа", f"Expected order_column='Дата заказа', got {result['order_column']}"
    assert result["value_column"] == "Сумма", f"Expected value_column='Сумма', got {result['value_column']}"
    assert result["window_size"] == 3, f"Expected window_size=3, got {result['window_size']}"
    
    # Verify rows structure
    rows = result["rows"]
    assert isinstance(rows, list), "rows should be list"
    
    if len(rows) > 0:
        print(f"  Found {len(rows)} rows with moving averages")
        
        # Verify each row has moving_average
        for i, row in enumerate(rows[:3]):  # Check first 3 rows
            assert isinstance(row, dict), f"Row {i} should be dict"
            assert "moving_average" in row, f"Row {i} missing 'moving_average'"
            
            # moving_average can be None for first few rows (not enough data for window)
            if row["moving_average"] is not None:
                assert isinstance(row["moving_average"], (int, float)), f"Row {i} moving_average should be numeric or None"
    else:
        print(f"  ℹ️  No rows returned (file might be empty)")
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    
    # Verify metadata
    metadata = result["metadata"]
    assert metadata["sheet_name"] == with_dates_fixture.sheet_name, "Metadata sheet_name mismatch"
    
    print(f"  ✅ Moving average calculated successfully")


def test_calculate_moving_average_window_7(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_moving_average with window_size=7 (7-day average)."""
    print(f"\n📉 Testing calculate_moving_average (window=7)...")
    
    result = mcp_call_tool("calculate_moving_average", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "order_column": "Дата заказа",
        "value_column": "Сумма",
        "window_size": 7
    })
    
    # Verify window_size
    assert result["window_size"] == 7, f"Expected window_size=7, got {result['window_size']}"
    
    # Verify response structure
    assert "rows" in result
    assert "excel_output" in result
    
    print(f"  ✅ 7-day moving average works")


def test_calculate_moving_average_with_filter(mcp_call_tool, with_dates_fixture):
    """Smoke: calculate_moving_average with filters applied."""
    print(f"\n📉 Testing calculate_moving_average with filter...")
    
    result = mcp_call_tool("calculate_moving_average", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name,
        "order_column": "Дата заказа",
        "value_column": "Сумма",
        "window_size": 3,
        "filters": [
            {"column": "Сумма", "operator": ">", "value": 0}
        ]
    })
    
    # Should work with filters
    assert "rows" in result
    assert "window_size" in result
    assert result["window_size"] == 3
    
    print(f"  ✅ Moving average with filter works")
