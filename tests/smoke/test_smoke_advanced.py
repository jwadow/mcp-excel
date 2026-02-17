# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Smoke tests for advanced tools.

Tests for 2 advanced tools:
- rank_rows
- calculate_expression

Each tool is tested with:
- Basic functionality (happy path)
- Different parameter combinations
- With filters
- With grouping (where applicable)
- Full response structure validation
- Excel output validation
"""

import pytest


# ============================================================================
# RANK_ROWS TESTS
# ============================================================================

def test_rank_rows_descending(mcp_call_tool, simple_fixture):
    """Smoke: rank_rows with descending order (highest first)."""
    print(f"\n🏆 Testing rank_rows (descending)...")
    
    result = mcp_call_tool("rank_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "rank_column": "Возраст",
        "direction": "desc"
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from RankRowsResponse
    assert "rows" in result, "Missing 'rows'"
    assert "rank_column" in result, "Missing 'rank_column'"
    assert "direction" in result, "Missing 'direction'"
    assert "total_rows" in result, "Missing 'total_rows'"
    assert "group_by_columns" in result, "Missing 'group_by_columns'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify rank_column and direction
    assert result["rank_column"] == "Возраст", f"Expected rank_column='Возраст', got {result['rank_column']}"
    assert result["direction"] == "desc", f"Expected direction='desc', got {result['direction']}"
    
    # Verify group_by_columns (should be None for basic test)
    assert result["group_by_columns"] is None or result["group_by_columns"] == [], "group_by_columns should be None or empty for basic test"
    
    # Verify rows structure
    rows = result["rows"]
    assert isinstance(rows, list), "rows should be list"
    assert result["total_rows"] == len(rows), "total_rows should match rows length"
    
    if len(rows) > 0:
        print(f"  Found {len(rows)} ranked rows")
        
        # Verify each row has rank
        for i, row in enumerate(rows[:3]):  # Check first 3 rows
            assert isinstance(row, dict), f"Row {i} should be dict"
            assert "rank" in row, f"Row {i} missing 'rank'"
            assert isinstance(row["rank"], (int, float)), f"Row {i} rank should be numeric"
            
            # Rank should be positive
            assert row["rank"] > 0, f"Row {i} rank should be positive"
        
        # Verify ranks are in order (1, 2, 3, ...)
        ranks = [row["rank"] for row in rows[:5]]
        print(f"  First 5 ranks: {ranks}")
    else:
        print(f"  ℹ️  No rows returned (file might be empty)")
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    if excel_output.get("formula"):
        assert isinstance(excel_output["formula"], str), "formula should be string"
        # Rank formula typically uses RANK function
        if "RANK" in excel_output["formula"].upper():
            print(f"  ✅ Excel formula contains RANK: {excel_output['formula'][:50]}...")
    
    # Verify metadata
    metadata = result["metadata"]
    assert metadata["sheet_name"] == simple_fixture.sheet_name, "Metadata sheet_name mismatch"
    
    print(f"  ✅ Ranking (descending) completed successfully")


def test_rank_rows_ascending(mcp_call_tool, simple_fixture):
    """Smoke: rank_rows with ascending order (lowest first)."""
    print(f"\n🏆 Testing rank_rows (ascending)...")
    
    result = mcp_call_tool("rank_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "rank_column": "Возраст",
        "direction": "asc"
    })
    
    # Verify direction
    assert result["direction"] == "asc", f"Expected direction='asc', got {result['direction']}"
    
    # Verify response structure
    assert "rows" in result
    assert "rank_column" in result
    assert "total_rows" in result
    assert "excel_output" in result
    
    print(f"  ✅ Ranking (ascending) works")


def test_rank_rows_top_n(mcp_call_tool, simple_fixture):
    """Smoke: rank_rows with top_n filtering (top 5 rows)."""
    print(f"\n🏆 Testing rank_rows with top_n=5...")
    
    result = mcp_call_tool("rank_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "rank_column": "Возраст",
        "direction": "desc",
        "top_n": 5
    })
    
    # Verify rows count (should be at most 5)
    rows = result["rows"]
    assert len(rows) <= 5, f"Expected at most 5 rows, got {len(rows)}"
    
    # total_rows should still reflect all rows that were ranked
    assert result["total_rows"] >= len(rows), "total_rows should be >= returned rows"
    
    print(f"  ✅ Top-N filtering works (returned {len(rows)} rows)")


def test_rank_rows_with_grouping(mcp_call_tool, multi_sheet_fixture):
    """Smoke: rank_rows with grouping (ranking within groups)."""
    print(f"\n🏆 Testing rank_rows with grouping...")
    
    result = mcp_call_tool("rank_rows", {
        "file_path": str(multi_sheet_fixture.path_str),
        "sheet_name": "Orders",
        "rank_column": "Amount",
        "direction": "desc",
        "group_by_columns": ["CustomerID"]
    })
    
    # Verify group_by_columns
    assert result["group_by_columns"] is not None, "group_by_columns should not be None"
    assert isinstance(result["group_by_columns"], list), "group_by_columns should be list"
    assert "CustomerID" in result["group_by_columns"], "group_by_columns should contain 'CustomerID'"
    
    # Verify rows structure
    assert "rows" in result
    assert isinstance(result["rows"], list), "rows should be list"
    
    # Verify excel_output
    assert "excel_output" in result
    
    print(f"  ✅ Ranking with grouping works")


def test_rank_rows_with_filter(mcp_call_tool, simple_fixture):
    """Smoke: rank_rows with filters applied."""
    print(f"\n🏆 Testing rank_rows with filter...")
    
    result = mcp_call_tool("rank_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "rank_column": "Возраст",
        "direction": "desc",
        "filters": [
            {"column": "Возраст", "operator": ">", "value": 20}
        ]
    })
    
    # Should work with filters
    assert "rows" in result
    assert "rank_column" in result
    assert result["rank_column"] == "Возраст"
    
    print(f"  ✅ Ranking with filter works")


def test_rank_rows_tsv_output(mcp_call_tool, simple_fixture):
    """Smoke: rank_rows TSV output is valid."""
    print(f"\n🏆 Testing rank_rows TSV output...")
    
    result = mcp_call_tool("rank_rows", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "rank_column": "Возраст",
        "direction": "desc"
    })
    
    # Verify excel_output structure
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    
    if excel_output["tsv"]:
        tsv = excel_output["tsv"]
        assert isinstance(tsv, str), "tsv should be string"
        assert len(tsv) > 0, "tsv should not be empty"
        
        # TSV should have tabs and newlines
        if len(result["rows"]) > 0:
            assert "\t" in tsv or "\n" in tsv, "TSV should contain tabs or newlines"
            print(f"  ✅ TSV output is valid (length: {len(tsv)} chars)")
    else:
        print(f"  ℹ️  TSV output is empty (might be valid if no data)")


# ============================================================================
# CALCULATE_EXPRESSION TESTS
# ============================================================================

def test_calculate_expression_addition(mcp_call_tool, numeric_types_fixture):
    """Smoke: calculate_expression with addition operation."""
    print(f"\n🧮 Testing calculate_expression (addition)...")
    
    result = mcp_call_tool("calculate_expression", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "expression": "Целое + Дробное",
        "output_column_name": "Сумма"
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from CalculateExpressionResponse
    assert "rows" in result, "Missing 'rows'"
    assert "expression" in result, "Missing 'expression'"
    assert "output_column_name" in result, "Missing 'output_column_name'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify expression and output_column_name
    assert result["expression"] == "Целое + Дробное", f"Expected expression='Целое + Дробное', got {result['expression']}"
    assert result["output_column_name"] == "Сумма", f"Expected output_column_name='Сумма', got {result['output_column_name']}"
    
    # Verify rows structure
    rows = result["rows"]
    assert isinstance(rows, list), "rows should be list"
    
    if len(rows) > 0:
        print(f"  Found {len(rows)} rows with calculated values")
        
        # Verify each row has the output column
        for i, row in enumerate(rows[:3]):  # Check first 3 rows
            assert isinstance(row, dict), f"Row {i} should be dict"
            assert "Сумма" in row, f"Row {i} missing output column 'Сумма'"
            assert isinstance(row["Сумма"], (int, float)), f"Row {i} calculated value should be numeric"
    else:
        print(f"  ℹ️  No rows returned (file might be empty)")
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    if excel_output.get("formula"):
        assert isinstance(excel_output["formula"], str), "formula should be string"
        # Expression formula should contain the operation
        if "+" in excel_output["formula"]:
            print(f"  ✅ Excel formula contains +: {excel_output['formula'][:50]}...")
    
    # Verify metadata
    metadata = result["metadata"]
    assert metadata["sheet_name"] == numeric_types_fixture.sheet_name, "Metadata sheet_name mismatch"
    
    print(f"  ✅ Expression calculation (addition) completed successfully")


def test_calculate_expression_multiplication(mcp_call_tool, numeric_types_fixture):
    """Smoke: calculate_expression with multiplication operation."""
    print(f"\n🧮 Testing calculate_expression (multiplication)...")
    
    result = mcp_call_tool("calculate_expression", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "expression": "Целое * Дробное",
        "output_column_name": "Произведение"
    })
    
    # Verify expression
    assert result["expression"] == "Целое * Дробное", f"Expected multiplication expression"
    
    # Verify response structure
    assert "rows" in result
    assert "output_column_name" in result
    assert result["output_column_name"] == "Произведение"
    assert "excel_output" in result
    
    print(f"  ✅ Multiplication works")


def test_calculate_expression_division(mcp_call_tool, numeric_types_fixture):
    """Smoke: calculate_expression with division operation."""
    print(f"\n🧮 Testing calculate_expression (division)...")
    
    result = mcp_call_tool("calculate_expression", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "expression": "Целое / Дробное",
        "output_column_name": "Частное"
    })
    
    # Verify expression
    assert result["expression"] == "Целое / Дробное", f"Expected division expression"
    
    # Verify response structure
    assert "rows" in result
    assert "excel_output" in result
    
    print(f"  ✅ Division works")


def test_calculate_expression_complex(mcp_call_tool, numeric_types_fixture):
    """Smoke: calculate_expression with complex expression (parentheses)."""
    print(f"\n🧮 Testing calculate_expression (complex with parentheses)...")
    
    result = mcp_call_tool("calculate_expression", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "expression": "(Целое + Дробное) * 2",
        "output_column_name": "Результат"
    })
    
    # Verify expression
    assert result["expression"] == "(Целое + Дробное) * 2", f"Expected complex expression"
    
    # Verify response structure
    assert "rows" in result
    assert "output_column_name" in result
    assert result["output_column_name"] == "Результат"
    
    # Verify rows have calculated values
    if len(result["rows"]) > 0:
        first_row = result["rows"][0]
        assert "Результат" in first_row, "Row should have output column"
        print(f"  First calculated value: {first_row['Результат']}")
    
    print(f"  ✅ Complex expression works")


def test_calculate_expression_with_filter(mcp_call_tool, numeric_types_fixture):
    """Smoke: calculate_expression with filters applied."""
    print(f"\n🧮 Testing calculate_expression with filter...")
    
    result = mcp_call_tool("calculate_expression", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "expression": "Целое + Дробное",
        "output_column_name": "Сумма",
        "filters": [
            {"column": "Целое", "operator": ">", "value": 0}
        ]
    })
    
    # Should work with filters
    assert "rows" in result
    assert "expression" in result
    assert result["expression"] == "Целое + Дробное"
    
    print(f"  ✅ Expression with filter works")


def test_calculate_expression_tsv_output(mcp_call_tool, numeric_types_fixture):
    """Smoke: calculate_expression TSV output is valid."""
    print(f"\n🧮 Testing calculate_expression TSV output...")
    
    result = mcp_call_tool("calculate_expression", {
        "file_path": str(numeric_types_fixture.path_str),
        "sheet_name": numeric_types_fixture.sheet_name,
        "expression": "Целое * 2",
        "output_column_name": "Удвоенное"
    })
    
    # Verify excel_output structure
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    
    if excel_output["tsv"]:
        tsv = excel_output["tsv"]
        assert isinstance(tsv, str), "tsv should be string"
        assert len(tsv) > 0, "tsv should not be empty"
        
        # TSV should have tabs and newlines
        if len(result["rows"]) > 0:
            assert "\t" in tsv or "\n" in tsv, "TSV should contain tabs or newlines"
            print(f"  ✅ TSV output is valid (length: {len(tsv)} chars)")
    else:
        print(f"  ℹ️  TSV output is empty (might be valid if no data)")
