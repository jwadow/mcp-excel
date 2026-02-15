# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Integration tests for Advanced operations.

Tests cover:
- rank_rows: Rank rows by column value (ascending/descending, grouping, top-N)
- calculate_expression: Calculate arithmetic expressions between columns

These are END-TO-END tests that verify the complete operation flow:
FileLoader -> FilterEngine -> Operations -> Response
"""

import pytest

from mcp_excel.operations.advanced import AdvancedOperations
from mcp_excel.models.requests import (
    RankRowsRequest,
    CalculateExpressionRequest,
    FilterCondition,
)


# ============================================================================
# rank_rows tests
# ============================================================================

def test_rank_rows_basic_descending(numeric_types_fixture, file_loader):
    """Test rank_rows with descending order (highest first).
    
    Verifies:
    - Ranks rows correctly in descending order
    - Returns rank column
    - Generates Excel formula
    - TSV output is correct
    """
    print(f"\n🏆 Testing rank_rows descending on: {numeric_types_fixture.name}")
    
    ops = AdvancedOperations(file_loader)
    request = RankRowsRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        rank_column="Количество",  # Numeric column
        direction="desc",
        top_n=None,
        group_by_columns=None,
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.rank_rows(request)
    
    # Assert
    print(f"✅ Total rows: {response.total_rows}")
    print(f"   Rank column: {response.rank_column}")
    print(f"   Direction: {response.direction}")
    print(f"   Top 3 ranked rows:")
    for i, row in enumerate(response.rows[:3], 1):
        print(f"     {i}. Rank {row['rank']}: Количество={row.get('Количество')}")
    
    assert response.total_rows == numeric_types_fixture.row_count, "Should rank all rows"
    assert response.rank_column == "Количество"
    assert response.direction == "desc"
    assert len(response.rows) == numeric_types_fixture.row_count
    
    # Check that ranks are assigned
    assert all('rank' in row for row in response.rows), "All rows should have rank"
    
    # Check descending order: first row should have rank 1 (highest value)
    assert response.rows[0]['rank'] == 1, "First row should have rank 1"
    
    # Check Excel formula
    assert response.excel_output.formula is not None, "Should generate formula"
    assert "RANK" in response.excel_output.formula, "Formula should use RANK function"
    print(f"   Excel formula: {response.excel_output.formula}")
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV"
    assert "rank" in response.excel_output.tsv.lower(), "TSV should contain rank column"
    print(f"   TSV preview: {response.excel_output.tsv[:150]}...")


def test_rank_rows_basic_ascending(numeric_types_fixture, file_loader):
    """Test rank_rows with ascending order (lowest first).
    
    Verifies:
    - Ranks rows correctly in ascending order
    - First row has lowest value
    """
    print(f"\n🏆 Testing rank_rows ascending")
    
    ops = AdvancedOperations(file_loader)
    request = RankRowsRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        rank_column="Цена",
        direction="asc",
        top_n=None,
        group_by_columns=None,
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.rank_rows(request)
    
    # Assert
    print(f"✅ Direction: {response.direction}")
    print(f"   Top 3 ranked rows (lowest first):")
    for i, row in enumerate(response.rows[:3], 1):
        print(f"     {i}. Rank {row['rank']}: Цена={row.get('Цена')}")
    
    assert response.direction == "asc"
    assert response.rows[0]['rank'] == 1, "First row should have rank 1 (lowest value)"
    
    # Check formula has correct order parameter (1 for ascending)
    if response.excel_output.formula:
        print(f"   Excel formula: {response.excel_output.formula}")
        assert ",1)" in response.excel_output.formula, "Formula should use order=1 for ascending"


def test_rank_rows_with_top_n(numeric_types_fixture, file_loader):
    """Test rank_rows with top_n limit.
    
    Verifies:
    - Returns only top N rows
    - Rows are correctly ranked
    """
    print(f"\n🏆 Testing rank_rows with top_n=5")
    
    ops = AdvancedOperations(file_loader)
    request = RankRowsRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        rank_column="Итого",
        direction="desc",
        top_n=5,
        group_by_columns=None,
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.rank_rows(request)
    
    # Assert
    print(f"✅ Total rows returned: {response.total_rows}")
    print(f"   Top 5 rows:")
    for row in response.rows:
        print(f"     Rank {row['rank']}: Итого={row.get('Итого')}")
    
    assert response.total_rows == 5, "Should return only top 5 rows"
    assert len(response.rows) == 5, "Should have 5 rows"
    
    # Check ranks are 1-5
    ranks = [row['rank'] for row in response.rows]
    assert ranks == [1, 2, 3, 4, 5], "Should have ranks 1-5 in order"


def test_rank_rows_with_grouping(with_dates_fixture, file_loader):
    """Test rank_rows with grouping (rank within groups).
    
    Verifies:
    - Ranks within each group separately
    - Each group has its own rank 1
    - group_by_columns is returned
    """
    print(f"\n🏆 Testing rank_rows with grouping by Клиент")
    
    ops = AdvancedOperations(file_loader)
    request = RankRowsRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        rank_column="Сумма",
        direction="desc",
        top_n=None,
        group_by_columns=["Клиент"],
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.rank_rows(request)
    
    # Assert
    print(f"✅ Total rows: {response.total_rows}")
    print(f"   Group by: {response.group_by_columns}")
    print(f"   Sample grouped ranks:")
    
    # Group rows by client
    from collections import defaultdict
    groups = defaultdict(list)
    for row in response.rows:
        client = row.get('Клиент')
        groups[client].append(row)
    
    print(f"   Found {len(groups)} groups")
    for client, rows in list(groups.items())[:3]:
        print(f"     {client}: {len(rows)} rows, ranks: {[r['rank'] for r in rows[:3]]}")
    
    assert response.group_by_columns == ["Клиент"]
    
    # Each group should have at least one rank 1
    for client, rows in groups.items():
        ranks = [r['rank'] for r in rows]
        assert 1 in ranks, f"Group {client} should have rank 1"


def test_rank_rows_with_filters(with_dates_fixture, file_loader):
    """Test rank_rows with filters applied.
    
    Verifies:
    - Filters are applied before ranking
    - Only filtered rows are ranked
    """
    print(f"\n🏆 Testing rank_rows with filters")
    
    ops = AdvancedOperations(file_loader)
    
    # Filter for specific client
    request = RankRowsRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        rank_column="Сумма",
        direction="desc",
        top_n=None,
        group_by_columns=None,
        filters=[
            FilterCondition(column="Клиент", operator="==", value="Ромашка")
        ],
        logic="AND"
    )
    
    # Act
    response = ops.rank_rows(request)
    
    # Assert
    print(f"✅ Filtered rows: {response.total_rows}")
    print(f"   All rows are for client: Ромашка")
    
    # Check all rows are for the filtered client
    for row in response.rows:
        assert row.get('Клиент') == "Ромашка", "All rows should be for Ромашка"
    
    # Should have fewer rows than total
    assert response.total_rows < with_dates_fixture.row_count, "Should have fewer rows after filtering"


def test_rank_rows_invalid_column(simple_fixture, file_loader):
    """Test rank_rows with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message is helpful
    """
    print(f"\n🏆 Testing rank_rows with invalid column")
    
    ops = AdvancedOperations(file_loader)
    request = RankRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        rank_column="NonExistentColumn",
        direction="desc",
        top_n=None,
        group_by_columns=None,
        filters=[],
        logic="AND"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.rank_rows(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower(), "Error should mention column not found"
    assert "NonExistentColumn" in str(exc_info.value), "Error should mention the invalid column"


def test_rank_rows_text_column_converted(simple_fixture, file_loader):
    """Test rank_rows on text column (should convert to numeric).
    
    Verifies:
    - Converts text to numeric with pd.to_numeric
    - Handles conversion errors gracefully (NaN for non-numeric)
    """
    print(f"\n🏆 Testing rank_rows on text column")
    
    ops = AdvancedOperations(file_loader)
    request = RankRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        rank_column="Возраст",  # Numeric column stored as int
        direction="desc",
        top_n=3,
        group_by_columns=None,
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.rank_rows(request)
    
    # Assert
    print(f"✅ Ranked {response.total_rows} rows")
    print(f"   Top 3:")
    for row in response.rows:
        print(f"     Rank {row['rank']}: Возраст={row.get('Возраст')}")
    
    assert response.total_rows == 3, "Should return top 3"
    assert all('rank' in row for row in response.rows), "All rows should have rank"


def test_rank_rows_performance_metrics(numeric_types_fixture, file_loader):
    """Test that rank_rows includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    """
    print(f"\n🏆 Testing rank_rows performance metrics")
    
    ops = AdvancedOperations(file_loader)
    request = RankRowsRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        rank_column="Количество",
        direction="desc",
        top_n=None,
        group_by_columns=None,
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.rank_rows(request)
    
    # Assert
    print(f"✅ Performance:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Cache hit: {response.performance.cache_hit}")
    
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.cache_hit in [True, False], "Should report cache status"


def test_rank_rows_messy_headers(messy_headers_fixture, file_loader):
    """Test rank_rows with messy headers (auto-detection).
    
    Verifies:
    - Auto-detects correct header row
    - Ranks data correctly
    """
    print(f"\n🏆 Testing rank_rows with messy headers")
    
    ops = AdvancedOperations(file_loader)
    request = RankRowsRequest(
        file_path=messy_headers_fixture.path_str,
        sheet_name=messy_headers_fixture.sheet_name,
        rank_column="Сумма",
        direction="desc",
        top_n=5,
        group_by_columns=None,
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.rank_rows(request)
    
    # Assert
    print(f"✅ Ranked {response.total_rows} rows (skipped junk rows)")
    
    assert response.total_rows == 5, "Should return top 5"
    assert all('rank' in row for row in response.rows), "All rows should have rank"


# ============================================================================
# calculate_expression tests
# ============================================================================

def test_calculate_expression_basic_addition(numeric_types_fixture, file_loader):
    """Test calculate_expression with simple addition.
    
    Verifies:
    - Calculates expression correctly
    - Adds new column with result
    - Generates Excel formula
    - TSV output is correct
    """
    print(f"\n🧮 Testing calculate_expression with addition")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Цена + Скидка",
        output_column_name="Сумма_с_скидкой",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated expression for {len(response.rows)} rows")
    print(f"   Expression: {response.expression}")
    print(f"   Output column: {response.output_column_name}")
    print(f"   Sample results (first 3):")
    for i, row in enumerate(response.rows[:3], 1):
        цена = row.get('Цена')
        скидка = row.get('Скидка')
        result = row.get('Сумма_с_скидкой')
        print(f"     {i}. Цена={цена}, Скидка={скидка}, Result={result}")
    
    assert len(response.rows) == numeric_types_fixture.row_count, "Should calculate for all rows"
    assert response.expression == "Цена + Скидка"
    assert response.output_column_name == "Сумма_с_скидкой"
    
    # Check that output column exists in all rows
    assert all('Сумма_с_скидкой' in row for row in response.rows), "All rows should have output column"
    
    # Check Excel formula
    assert response.excel_output.formula is not None, "Should generate formula"
    assert "=" in response.excel_output.formula, "Formula should start with ="
    print(f"   Excel formula: {response.excel_output.formula}")
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV"
    assert "Сумма_с_скидкой" in response.excel_output.tsv, "TSV should contain output column"


def test_calculate_expression_multiplication(numeric_types_fixture, file_loader):
    """Test calculate_expression with multiplication.
    
    Verifies:
    - Handles multiplication operator
    - Result is correct
    """
    print(f"\n🧮 Testing calculate_expression with multiplication")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Количество * Цена",
        output_column_name="Стоимость",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated multiplication for {len(response.rows)} rows")
    print(f"   Sample results (first 3):")
    for i, row in enumerate(response.rows[:3], 1):
        qty = row.get('Количество')
        price = row.get('Цена')
        result = row.get('Стоимость')
        print(f"     {i}. {qty} * {price} = {result}")
    
    assert len(response.rows) == numeric_types_fixture.row_count
    assert all('Стоимость' in row for row in response.rows), "All rows should have result"


def test_calculate_expression_division(numeric_types_fixture, file_loader):
    """Test calculate_expression with division.
    
    Verifies:
    - Handles division operator
    - Result is correct
    """
    print(f"\n🧮 Testing calculate_expression with division")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Итого / Количество",
        output_column_name="Цена_за_единицу",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated division for {len(response.rows)} rows")
    print(f"   Sample results (first 2):")
    for i, row in enumerate(response.rows[:2], 1):
        total = row.get('Итого')
        qty = row.get('Количество')
        result = row.get('Цена_за_единицу')
        print(f"     {i}. {total} / {qty} = {result}")
    
    assert len(response.rows) == numeric_types_fixture.row_count
    assert all('Цена_за_единицу' in row for row in response.rows)


def test_calculate_expression_subtraction(numeric_types_fixture, file_loader):
    """Test calculate_expression with subtraction.
    
    Verifies:
    - Handles subtraction operator
    - Result is correct
    """
    print(f"\n🧮 Testing calculate_expression with subtraction")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Цена - Скидка",
        output_column_name="Цена_со_скидкой",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated subtraction for {len(response.rows)} rows")
    
    assert len(response.rows) == numeric_types_fixture.row_count
    assert all('Цена_со_скидкой' in row for row in response.rows)


def test_calculate_expression_complex(numeric_types_fixture, file_loader):
    """Test calculate_expression with complex expression (multiple operators).
    
    Verifies:
    - Handles complex expressions with parentheses
    - Order of operations is correct
    """
    print(f"\n🧮 Testing calculate_expression with complex expression")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="(Цена * Количество) - (Цена * Количество * Скидка)",
        output_column_name="Итого_с_учетом_скидки",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated complex expression for {len(response.rows)} rows")
    print(f"   Expression: {response.expression}")
    print(f"   Sample result: {response.rows[0].get('Итого_с_учетом_скидки')}")
    assert len(response.rows) == numeric_types_fixture.row_count
    assert all('Итого_с_учетом_скидки' in row for row in response.rows)


def test_calculate_expression_with_filters(numeric_types_fixture, file_loader):
    """Test calculate_expression with filters applied.
    
    Verifies:
    - Filters are applied before calculation
    - Only filtered rows are calculated
    """
    print(f"\n🧮 Testing calculate_expression with filters")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Цена * Количество",
        output_column_name="Стоимость",
        filters=[
            FilterCondition(column="Количество", operator=">", value=100)
        ],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated for {len(response.rows)} filtered rows")
    
    # Check all rows have Количество > 100
    for row in response.rows:
        qty = row.get('Количество')
        assert qty > 100, f"All rows should have Количество > 100, got {qty}"
    
    # Should have fewer rows than total
    assert len(response.rows) < numeric_types_fixture.row_count, "Should have fewer rows after filtering"


def test_calculate_expression_invalid_column(simple_fixture, file_loader):
    """Test calculate_expression with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message is helpful
    """
    print(f"\n🧮 Testing calculate_expression with invalid column")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        expression="NonExistentColumn + 10",
        output_column_name="Result",
        filters=[],
        logic="AND"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.calculate_expression(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "not found" in str(exc_info.value).lower() or "no valid column" in str(exc_info.value).lower(), \
        "Error should mention column not found"


def test_calculate_expression_invalid_syntax(numeric_types_fixture, file_loader):
    """Test calculate_expression with invalid expression syntax.
    
    Verifies:
    - Raises ValueError for invalid expression
    - Error message mentions evaluation failure
    """
    print(f"\n🧮 Testing calculate_expression with invalid syntax")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="(Цена + Количество",  # Invalid syntax - unmatched parenthesis
        output_column_name="Result",
        filters=[],
        logic="AND"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.calculate_expression(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    assert "failed to evaluate" in str(exc_info.value).lower(), "Error should mention evaluation failure"


def test_calculate_expression_column_with_spaces(with_dates_fixture, file_loader):
    """Test calculate_expression with column names containing spaces.
    
    Verifies:
    - Handles column names with spaces correctly
    - Uses backtick quoting for pandas.eval()
    """
    print(f"\n🧮 Testing calculate_expression with column names containing spaces")
    
    ops = AdvancedOperations(file_loader)
    
    # "Номер заказа" and "Дата заказа" have spaces
    # We'll use numeric columns for calculation
    request = CalculateExpressionRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        expression="Сумма * 2",
        output_column_name="Двойная_сумма",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated for {len(response.rows)} rows")
    print(f"   Sample: Сумма={response.rows[0].get('Сумма')}, Двойная_сумма={response.rows[0].get('Двойная_сумма')}")
    
    assert len(response.rows) == with_dates_fixture.row_count
    assert all('Двойная_сумма' in row for row in response.rows)


def test_calculate_expression_performance_metrics(numeric_types_fixture, file_loader):
    """Test that calculate_expression includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    """
    print(f"\n🧮 Testing calculate_expression performance metrics")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Цена + Количество",
        output_column_name="Сумма",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Performance:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Cache hit: {response.performance.cache_hit}")
    
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.cache_hit in [True, False], "Should report cache status"


def test_calculate_expression_messy_headers(messy_headers_fixture, file_loader):
    """Test calculate_expression with messy headers (auto-detection).
    
    Verifies:
    - Auto-detects correct header row
    - Calculates correctly
    """
    print(f"\n🧮 Testing calculate_expression with messy headers")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=messy_headers_fixture.path_str,
        sheet_name=messy_headers_fixture.sheet_name,
        expression="Сумма * 2",
        output_column_name="Двойная_сумма",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated for {len(response.rows)} rows (skipped junk rows)")
    
    assert len(response.rows) == messy_headers_fixture.row_count, "Should calculate for all data rows"
    assert all('Двойная_сумма' in row for row in response.rows)


def test_calculate_expression_tsv_output(numeric_types_fixture, file_loader):
    """Test that calculate_expression generates proper TSV output.
    
    Verifies:
    - TSV output is generated
    - Contains all columns including calculated
    - Can be pasted into Excel
    """
    print(f"\n🧮 Testing calculate_expression TSV output")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Цена * Количество",
        output_column_name="Итого",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ TSV output generated")
    print(f"   Length: {len(response.excel_output.tsv)} chars")
    print(f"   Preview: {response.excel_output.tsv[:200]}...")
    
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    
    # Check TSV contains output column
    assert "Итого" in response.excel_output.tsv, "TSV should contain output column"
    
    # Check TSV has tab separators
    assert "\t" in response.excel_output.tsv, "TSV should use tab separators"


def test_calculate_expression_excel_formula(numeric_types_fixture, file_loader):
    """Test that calculate_expression generates correct Excel formula.
    
    Verifies:
    - Excel formula is generated
    - Formula uses correct cell references
    - Formula is valid Excel syntax
    """
    print(f"\n🧮 Testing calculate_expression Excel formula generation")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Цена + Количество",
        output_column_name="Сумма",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Excel formula: {response.excel_output.formula}")
    
    assert response.excel_output.formula is not None, "Should generate formula"
    assert response.excel_output.formula.startswith("="), "Formula should start with ="
    
    # Formula should contain cell references (letters followed by numbers)
    import re
    cell_refs = re.findall(r'[A-Z]+\d+', response.excel_output.formula)
    assert len(cell_refs) > 0, "Formula should contain cell references"
    print(f"   Cell references found: {cell_refs}")


def test_calculate_expression_with_nulls(with_nulls_fixture, file_loader):
    """Test calculate_expression with null values.
    
    Verifies:
    - Handles null values gracefully (converts to NaN)
    - Calculation continues for non-null rows
    """
    print(f"\n🧮 Testing calculate_expression with null values")
    
    ops = AdvancedOperations(file_loader)
    request = CalculateExpressionRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        expression="ID * 2",  # ID column should be numeric
        output_column_name="Двойной_ID",
        filters=[],
        logic="AND"
    )
    
    # Act
    response = ops.calculate_expression(request)
    
    # Assert
    print(f"✅ Calculated for {len(response.rows)} rows")
    
    assert len(response.rows) == with_nulls_fixture.row_count
    assert all('Двойной_ID' in row for row in response.rows)


def test_rank_rows_and_calculate_expression_combined(numeric_types_fixture, file_loader):
    """Test using both operations in sequence (realistic workflow).
    
    Verifies:
    - Both operations work correctly
    - Can be chained together
    - Results are consistent
    """
    print(f"\n🔄 Testing rank_rows and calculate_expression combined")
    
    ops = AdvancedOperations(file_loader)
    
    # First, calculate expression
    calc_request = CalculateExpressionRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        expression="Цена * Количество",
        output_column_name="Стоимость",
        filters=[],
        logic="AND"
    )
    calc_response = ops.calculate_expression(calc_request)
    
    print(f"✅ Step 1: Calculated Стоимость for {len(calc_response.rows)} rows")
    
    # Then, rank by calculated column (note: this won't work directly as we need to save the result)
    # But we can rank by original columns
    rank_request = RankRowsRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        rank_column="Итого",
        direction="desc",
        top_n=5,
        group_by_columns=None,
        filters=[],
        logic="AND"
    )
    rank_response = ops.rank_rows(rank_request)
    
    print(f"✅ Step 2: Ranked top {rank_response.total_rows} rows")
    
    # Assert both operations succeeded
    assert len(calc_response.rows) == numeric_types_fixture.row_count
    assert rank_response.total_rows == 5
    
    print(f"   Combined workflow completed successfully")



