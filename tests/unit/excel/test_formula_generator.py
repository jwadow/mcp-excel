# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Unit tests for FormulaGenerator component.

Tests cover:
- Sheet name escaping
- Value escaping and formula injection protection
- Column letter generation (A-Z, AA-ZZ, etc.)
- Basic formula generation (COUNTIF, SUMIF, AVERAGEIF, etc.)
- Complex formula generation with filters (all 12 operators)
- DateTime handling in formulas
- Edge cases and error handling
"""

import pytest
import pandas as pd

from mcp_excel.excel.formula_generator import FormulaGenerator
from mcp_excel.models.requests import FilterCondition


# ============================================================================
# BASIC METHODS TESTS
# ============================================================================

def test_escape_sheet_name_simple():
    """Test sheet name escaping for simple names."""
    print("\n📂 Testing sheet name escaping (simple)")
    
    gen = FormulaGenerator("Sheet1")
    assert gen._sheet_name == "Sheet1", "Simple name should not be quoted"
    
    gen = FormulaGenerator("Data")
    assert gen._sheet_name == "Data", "Simple name should not be quoted"
    
    print("✅ Simple sheet names handled correctly")


def test_escape_sheet_name_with_spaces():
    """Test sheet name escaping for names with spaces."""
    print("\n📂 Testing sheet name escaping (spaces)")
    
    gen = FormulaGenerator("Sales Data")
    assert gen._sheet_name == "'Sales Data'", "Name with spaces should be quoted"
    
    gen = FormulaGenerator("Q1 Report")
    assert gen._sheet_name == "'Q1 Report'", "Name with spaces should be quoted"
    
    print("✅ Sheet names with spaces quoted correctly")


def test_escape_sheet_name_with_special_chars():
    """Test sheet name escaping for names with special characters."""
    print("\n📂 Testing sheet name escaping (special chars)")
    
    gen = FormulaGenerator("Sales-2024")
    assert gen._sheet_name == "'Sales-2024'", "Name with dash should be quoted"
    
    gen = FormulaGenerator("Q1'Report")
    assert gen._sheet_name == "'Q1'Report'", "Name with apostrophe should be quoted"
    
    print("✅ Sheet names with special chars quoted correctly")


def test_escape_value_string():
    """Test value escaping for strings."""
    print("\n📂 Testing value escaping (strings)")
    
    gen = FormulaGenerator("Sheet1")
    
    # Simple string
    assert gen._escape_value("test") == '"test"', "String should be quoted"
    
    # String with quotes
    assert gen._escape_value('test"value') == '"test""value"', "Quotes should be doubled"
    
    # Cyrillic
    assert gen._escape_value("Москва") == '"Москва"', "Cyrillic should work"
    
    print("✅ String values escaped correctly")


def test_escape_value_formula_injection():
    """Test formula injection protection."""
    print("\n📂 Testing formula injection protection")
    
    gen = FormulaGenerator("Sheet1")
    
    # Values starting with dangerous chars should be prefixed with apostrophe
    assert gen._escape_value("=1+1") == '"\'=1+1"', "= prefix should be escaped"
    assert gen._escape_value("+7999") == '"\'+7999"', "+ prefix should be escaped"
    assert gen._escape_value("-100") == '"\'-100"', "- prefix should be escaped"
    assert gen._escape_value("@user") == '"\'@user"', "@ prefix should be escaped"
    
    print("✅ Formula injection protection works")


def test_escape_value_numbers():
    """Test value escaping for numbers."""
    print("\n📂 Testing value escaping (numbers)")
    
    gen = FormulaGenerator("Sheet1")
    
    assert gen._escape_value(42) == "42", "Integer should not be quoted"
    assert gen._escape_value(3.14) == "3.14", "Float should not be quoted"
    assert gen._escape_value(0) == "0", "Zero should work"
    
    print("✅ Number values handled correctly")


def test_escape_value_none():
    """Test value escaping for None."""
    print("\n📂 Testing value escaping (None)")
    
    gen = FormulaGenerator("Sheet1")
    
    assert gen._escape_value(None) == '""', "None should become empty string"
    
    print("✅ None value handled correctly")


def test_column_letter_basic():
    """Test column letter generation for A-Z."""
    print("\n📂 Testing column letter generation (A-Z)")
    
    gen = FormulaGenerator("Sheet1")
    
    assert gen._column_letter(0) == "A", "Index 0 should be A"
    assert gen._column_letter(1) == "B", "Index 1 should be B"
    assert gen._column_letter(25) == "Z", "Index 25 should be Z"
    
    print("✅ Basic column letters (A-Z) generated correctly")


def test_column_letter_extended():
    """Test column letter generation for AA-ZZ and beyond."""
    print("\n📂 Testing column letter generation (AA-ZZ)")
    
    gen = FormulaGenerator("Sheet1")
    
    assert gen._column_letter(26) == "AA", "Index 26 should be AA"
    assert gen._column_letter(27) == "AB", "Index 27 should be AB"
    assert gen._column_letter(51) == "AZ", "Index 51 should be AZ"
    assert gen._column_letter(52) == "BA", "Index 52 should be BA"
    assert gen._column_letter(701) == "ZZ", "Index 701 should be ZZ"
    assert gen._column_letter(702) == "AAA", "Index 702 should be AAA"
    
    print("✅ Extended column letters (AA-ZZ, AAA) generated correctly")


def test_get_column_range():
    """Test Excel range generation."""
    print("\n📂 Testing column range generation")
    
    gen = FormulaGenerator("Sheet1")
    
    range_a = gen._get_column_range("Name", 0)
    assert range_a == "Sheet1!$A:$A", "Range for column A should be correct"
    
    range_b = gen._get_column_range("Age", 1)
    assert range_b == "Sheet1!$B:$B", "Range for column B should be correct"
    
    # Test with quoted sheet name
    gen_quoted = FormulaGenerator("Sales Data")
    range_quoted = gen_quoted._get_column_range("Amount", 2)
    assert range_quoted == "'Sales Data'!$C:$C", "Range with quoted sheet should work"
    
    print("✅ Column ranges generated correctly")


# ============================================================================
# BASIC FORMULA GENERATION TESTS
# ============================================================================

def test_generate_countif():
    """Test COUNTIF formula generation."""
    print("\n📂 Testing COUNTIF formula generation")
    
    gen = FormulaGenerator("Sheet1")
    
    formula = gen.generate_countif("Sheet1!$A:$A", "Moscow")
    assert formula == '=COUNTIF(Sheet1!$A:$A,"Moscow")', "COUNTIF formula should be correct"
    
    formula_num = gen.generate_countif("Sheet1!$B:$B", 25)
    assert formula_num == "=COUNTIF(Sheet1!$B:$B,25)", "COUNTIF with number should work"
    
    print("✅ COUNTIF formulas generated correctly")


def test_generate_sumif():
    """Test SUMIF formula generation."""
    print("\n📂 Testing SUMIF formula generation")
    
    gen = FormulaGenerator("Sheet1")
    
    formula = gen.generate_sumif("Sheet1!$A:$A", "Moscow", "Sheet1!$B:$B")
    assert formula == '=SUMIF(Sheet1!$A:$A,"Moscow",Sheet1!$B:$B)', "SUMIF formula should be correct"
    
    print("✅ SUMIF formulas generated correctly")


def test_generate_averageif():
    """Test AVERAGEIF formula generation."""
    print("\n📂 Testing AVERAGEIF formula generation")
    
    gen = FormulaGenerator("Sheet1")
    
    formula = gen.generate_averageif("Sheet1!$A:$A", "Moscow", "Sheet1!$B:$B")
    assert formula == '=AVERAGEIF(Sheet1!$A:$A,"Moscow",Sheet1!$B:$B)', "AVERAGEIF formula should be correct"
    
    print("✅ AVERAGEIF formulas generated correctly")


def test_generate_countifs():
    """Test COUNTIFS formula generation (multiple criteria)."""
    print("\n📂 Testing COUNTIFS formula generation")
    
    gen = FormulaGenerator("Sheet1")
    
    formula = gen.generate_countifs(
        ["Sheet1!$A:$A", "Sheet1!$B:$B"],
        ["Moscow", 25]
    )
    assert formula == '=COUNTIFS(Sheet1!$A:$A,"Moscow",Sheet1!$B:$B,25)', "COUNTIFS formula should be correct"
    
    print("✅ COUNTIFS formulas generated correctly")


def test_generate_sumifs():
    """Test SUMIFS formula generation (multiple criteria)."""
    print("\n📂 Testing SUMIFS formula generation")
    
    gen = FormulaGenerator("Sheet1")
    
    formula = gen.generate_sumifs(
        "Sheet1!$C:$C",
        ["Sheet1!$A:$A", "Sheet1!$B:$B"],
        ["Moscow", 25]
    )
    assert formula == '=SUMIFS(Sheet1!$C:$C,Sheet1!$A:$A,"Moscow",Sheet1!$B:$B,25)', "SUMIFS formula should be correct"
    
    print("✅ SUMIFS formulas generated correctly")


# ============================================================================
# FILTER OPERATORS TESTS (12 operators)
# ============================================================================

def test_operator_equal():
    """Test == operator formula generation."""
    print("\n📂 Testing == operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"City": "Sheet1!$A:$A"}
    
    filters = [FilterCondition(column="City", operator="==", value="Moscow")]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    assert formula == '=COUNTIF(Sheet1!$A:$A,"Moscow")', "== operator should generate COUNTIF"
    print(f"   Formula: {formula}")
    print("✅ == operator works correctly")


def test_operator_not_equal():
    """Test != operator formula generation."""
    print("\n📂 Testing != operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"City": "Sheet1!$A:$A"}
    
    filters = [FilterCondition(column="City", operator="!=", value="Moscow")]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    assert formula == '=COUNTIF(Sheet1!$A:$A,"<>Moscow")', "!= operator should use <> in Excel"
    print(f"   Formula: {formula}")
    print("✅ != operator works correctly")


def test_operator_greater():
    """Test > operator formula generation."""
    print("\n📂 Testing > operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Age": "Sheet1!$B:$B"}
    
    filters = [FilterCondition(column="Age", operator=">", value=30)]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    assert formula == '=COUNTIF(Sheet1!$B:$B,">30")', "> operator should work"
    print(f"   Formula: {formula}")
    print("✅ > operator works correctly")


def test_operator_less():
    """Test < operator formula generation."""
    print("\n📂 Testing < operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Age": "Sheet1!$B:$B"}
    
    filters = [FilterCondition(column="Age", operator="<", value=30)]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    assert formula == '=COUNTIF(Sheet1!$B:$B,"<30")', "< operator should work"
    print(f"   Formula: {formula}")
    print("✅ < operator works correctly")


def test_operator_greater_equal():
    """Test >= operator formula generation."""
    print("\n📂 Testing >= operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Age": "Sheet1!$B:$B"}
    
    filters = [FilterCondition(column="Age", operator=">=", value=30)]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    assert formula == '=COUNTIF(Sheet1!$B:$B,">=30")', ">= operator should work"
    print(f"   Formula: {formula}")
    print("✅ >= operator works correctly")


def test_operator_less_equal():
    """Test <= operator formula generation."""
    print("\n📂 Testing <= operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Age": "Sheet1!$B:$B"}
    
    filters = [FilterCondition(column="Age", operator="<=", value=30)]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    assert formula == '=COUNTIF(Sheet1!$B:$B,"<=30")', "<= operator should work"
    print(f"   Formula: {formula}")
    print("✅ <= operator works correctly")


def test_operator_in():
    """Test 'in' operator formula generation."""
    print("\n📂 Testing 'in' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$C:$C"}
    
    filters = [FilterCondition(column="Status", operator="in", values=["Active", "Pending"])]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should use SUMPRODUCT for multiple values
    assert "SUMPRODUCT" in formula, "'in' operator should use SUMPRODUCT"
    assert "Active" in formula and "Pending" in formula, "Both values should be in formula"
    print(f"   Formula: {formula}")
    print("✅ 'in' operator works correctly")


def test_operator_not_in():
    """Test 'not_in' operator formula generation."""
    print("\n📂 Testing 'not_in' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$C:$C"}
    
    filters = [FilterCondition(column="Status", operator="not_in", values=["Cancelled", "Deleted"])]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should use COUNTA - SUMPRODUCT
    assert "COUNTA" in formula and "SUMPRODUCT" in formula, "'not_in' should use COUNTA-SUMPRODUCT"
    print(f"   Formula: {formula}")
    print("✅ 'not_in' operator works correctly")


def test_operator_contains():
    """Test 'contains' operator formula generation."""
    print("\n📂 Testing 'contains' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    
    filters = [FilterCondition(column="Name", operator="contains", value="John")]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should use wildcards: *John*
    assert formula == '=COUNTIF(Sheet1!$A:$A,"*John*")', "'contains' should use wildcards"
    print(f"   Formula: {formula}")
    print("✅ 'contains' operator works correctly")


def test_operator_startswith():
    """Test 'startswith' operator formula generation."""
    print("\n📂 Testing 'startswith' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    
    filters = [FilterCondition(column="Name", operator="startswith", value="John")]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should use wildcards: John*
    assert formula == '=COUNTIF(Sheet1!$A:$A,"John*")', "'startswith' should use wildcard at end"
    print(f"   Formula: {formula}")
    print("✅ 'startswith' operator works correctly")


def test_operator_endswith():
    """Test 'endswith' operator formula generation."""
    print("\n📂 Testing 'endswith' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    
    filters = [FilterCondition(column="Name", operator="endswith", value="son")]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should use wildcards: *son
    assert formula == '=COUNTIF(Sheet1!$A:$A,"*son")', "'endswith' should use wildcard at start"
    print(f"   Formula: {formula}")
    print("✅ 'endswith' operator works correctly")


def test_operator_is_null():
    """Test 'is_null' operator (no formula generated)."""
    print("\n📂 Testing 'is_null' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Email": "Sheet1!$D:$D"}
    
    filters = [FilterCondition(column="Email", operator="is_null", value=None)]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should return None (formula not supported due to full column reference issue)
    assert formula is None, "'is_null' should return None (not supported in formulas)"
    print("   ℹ️ Formula: None (by design - see ARCHITECTURE.md)")
    print("✅ 'is_null' operator handled correctly")


def test_operator_is_not_null():
    """Test 'is_not_null' operator formula generation."""
    print("\n📂 Testing 'is_not_null' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Email": "Sheet1!$D:$D"}
    
    filters = [FilterCondition(column="Email", operator="is_not_null", value=None)]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should use COUNTA
    assert formula == "=COUNTA(Sheet1!$D:$D)", "'is_not_null' should use COUNTA"
    print(f"   Formula: {formula}")
    print("✅ 'is_not_null' operator works correctly")


def test_operator_regex():
    """Test 'regex' operator (no formula generated)."""
    print("\n📂 Testing 'regex' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Email": "Sheet1!$D:$D"}
    
    filters = [FilterCondition(column="Email", operator="regex", value=r".*@gmail\.com")]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should return None (regex not supported in Excel formulas)
    assert formula is None, "'regex' should return None (not supported in Excel)"
    print("   ℹ️ Formula: None (regex not supported in Excel)")
    print("✅ 'regex' operator handled correctly")


# ============================================================================
# DATETIME TESTS
# ============================================================================

def test_format_date_for_excel():
    """Test datetime formatting as DATE() function."""
    print("\n📂 Testing datetime formatting")
    
    gen = FormulaGenerator("Sheet1")
    
    dt = pd.Timestamp("2024-03-15 14:30:00")
    date_func = gen._format_date_for_excel(dt)
    
    assert date_func == "DATE(2024,3,15)", "Should format as DATE(year,month,day)"
    print(f"   Timestamp: {dt}")
    print(f"   Excel DATE(): {date_func}")
    print("✅ DateTime formatted correctly")


def test_format_date_for_excel_null():
    """Test datetime formatting for null values."""
    print("\n📂 Testing datetime formatting (null)")
    
    gen = FormulaGenerator("Sheet1")
    
    dt = pd.NaT
    date_func = gen._format_date_for_excel(dt)
    
    assert date_func == '""', "NaT should become empty string"
    print("✅ Null datetime handled correctly")


def test_datetime_filter_conversion():
    """Test datetime filter value conversion."""
    print("\n📂 Testing datetime filter conversion")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Date": "Sheet1!$E:$E"}
    column_types = {"Date": "datetime"}
    
    # String datetime value should be converted to pd.Timestamp
    filters = [FilterCondition(column="Date", operator=">=", value="2024-01-01")]
    formula = gen.generate_from_filter("count", filters, column_ranges, column_types=column_types)
    
    # Should contain DATE() function
    assert "DATE(" in formula, "DateTime filter should use DATE() function"
    assert "2024" in formula, "Year should be in formula"
    print(f"   Formula: {formula}")
    print("✅ DateTime filter conversion works")


def test_datetime_comparison_operators():
    """Test datetime with comparison operators."""
    print("\n📂 Testing datetime comparison operators")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Date": "Sheet1!$E:$E"}
    column_types = {"Date": "datetime"}
    
    # Test >= operator with datetime
    filters = [FilterCondition(column="Date", operator=">=", value="2024-01-01")]
    formula = gen.generate_from_filter("count", filters, column_ranges, column_types=column_types)
    
    # Should use ">="&DATE(...)
    assert ">=" in formula and "DATE(" in formula, "Should use comparison with DATE()"
    print(f"   >= Formula: {formula}")
    
    # Test == operator with datetime
    filters_eq = [FilterCondition(column="Date", operator="==", value="2024-01-01")]
    formula_eq = gen.generate_from_filter("count", filters_eq, column_ranges, column_types=column_types)
    
    assert "DATE(" in formula_eq, "== should also use DATE()"
    print(f"   == Formula: {formula_eq}")
    print("✅ DateTime comparison operators work")


# ============================================================================
# MULTIPLE FILTERS TESTS
# ============================================================================

def test_multiple_filters_count():
    """Test formula generation with multiple filters (COUNTIFS)."""
    print("\n📂 Testing multiple filters (count)")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {
        "City": "Sheet1!$A:$A",
        "Age": "Sheet1!$B:$B"
    }
    
    filters = [
        FilterCondition(column="City", operator="==", value="Moscow"),
        FilterCondition(column="Age", operator=">", value=25)
    ]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    assert "COUNTIFS" in formula, "Multiple filters should use COUNTIFS"
    assert "Moscow" in formula and ">25" in formula, "Both criteria should be in formula"
    print(f"   Formula: {formula}")
    print("✅ Multiple filters (count) work correctly")


def test_multiple_filters_sum():
    """Test formula generation with multiple filters (SUMIFS)."""
    print("\n📂 Testing multiple filters (sum)")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {
        "City": "Sheet1!$A:$A",
        "Age": "Sheet1!$B:$B"
    }
    target_range = "Sheet1!$C:$C"
    
    filters = [
        FilterCondition(column="City", operator="==", value="Moscow"),
        FilterCondition(column="Age", operator=">", value=25)
    ]
    formula = gen.generate_from_filter("sum", filters, column_ranges, target_range=target_range)
    
    assert "SUMIFS" in formula, "Multiple filters with sum should use SUMIFS"
    assert target_range in formula, "Target range should be in formula"
    print(f"   Formula: {formula}")
    print("✅ Multiple filters (sum) work correctly")


def test_multiple_filters_with_python_only_operators():
    """Test that Python-only operators in multiple filters return None."""
    print("\n📂 Testing multiple filters with Python-only operators")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {
        "City": "Sheet1!$A:$A",
        "Status": "Sheet1!$B:$B"
    }
    
    # Mix of Excel-supported and Python-only operators
    filters = [
        FilterCondition(column="City", operator="==", value="Moscow"),
        FilterCondition(column="Status", operator="in", values=["Active", "Pending"])
    ]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should return None because 'in' operator in multiple filters is not supported
    assert formula is None, "Python-only operators in multiple filters should return None"
    print("   ℹ️ Formula: None (Python-only operator in multiple filters)")
    print("✅ Python-only operators handled correctly")


# ============================================================================
# EDGE CASES AND ERROR HANDLING
# ============================================================================

def test_no_filters_with_target_range():
    """Test formula generation without filters but with target range."""
    print("\n📂 Testing no filters with target range")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {}
    target_range = "Sheet1!$C:$C"
    
    # Count without filters
    formula_count = gen.generate_from_filter("count", [], column_ranges, target_range=target_range)
    assert formula_count == "=COUNTA(Sheet1!$C:$C)", "Count without filters should use COUNTA"
    
    # Sum without filters
    formula_sum = gen.generate_from_filter("sum", [], column_ranges, target_range=target_range)
    assert formula_sum == "=SUM(Sheet1!$C:$C)", "Sum without filters should use SUM"
    
    # Mean without filters
    formula_mean = gen.generate_from_filter("mean", [], column_ranges, target_range=target_range)
    assert formula_mean == "=AVERAGE(Sheet1!$C:$C)", "Mean without filters should use AVERAGE"
    
    print("✅ No filters with target range works correctly")


def test_missing_column_in_ranges():
    """Test error handling when column is not in ranges."""
    print("\n📂 Testing missing column error")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"City": "Sheet1!$A:$A"}
    
    filters = [FilterCondition(column="NonExistent", operator="==", value="test")]
    
    with pytest.raises(ValueError) as exc_info:
        gen.generate_from_filter("count", filters, column_ranges)
    
    assert "not found" in str(exc_info.value).lower(), "Should raise error for missing column"
    print(f"   ✅ Error caught: {exc_info.value}")
    print("✅ Missing column error handled correctly")


def test_get_references():
    """Test cell references generation."""
    print("\n📂 Testing cell references generation")
    
    gen = FormulaGenerator("Sheet1")
    column_names = ["Name", "Age", "City"]
    column_indices = {"Name": 0, "Age": 1, "City": 2}
    
    refs = gen.get_references(column_names, column_indices)
    
    assert refs["sheet"] == "Sheet1", "Sheet name should be in references"
    assert "Name" in refs["columns"], "Name column should be in references"
    assert refs["columns"]["Name"]["column"] == "A", "Name should be column A"
    assert refs["columns"]["Name"]["range"] == "$A:$A", "Range should be correct"
    assert refs["columns"]["Age"]["column"] == "B", "Age should be column B"
    
    print(f"   References: {refs}")
    print("✅ Cell references generated correctly")


def test_cyrillic_in_formulas():
    """Test Cyrillic values in formulas."""
    print("\n📂 Testing Cyrillic in formulas")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Город": "Sheet1!$A:$A"}
    
    filters = [FilterCondition(column="Город", operator="==", value="Москва")]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    assert "Москва" in formula, "Cyrillic should be preserved in formula"
    print(f"   Formula: {formula}")
    print("✅ Cyrillic handled correctly in formulas")


# ============================================================================
# ADDITIONAL COVERAGE TESTS
# ============================================================================

def test_operator_in_with_sum():
    """Test 'in' operator with sum operation."""
    print("\n📂 Testing 'in' operator with sum")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$C:$C"}
    target_range = "Sheet1!$D:$D"
    
    filters = [FilterCondition(column="Status", operator="in", values=["Active", "Pending"])]
    formula = gen.generate_from_filter("sum", filters, column_ranges, target_range=target_range)
    
    # Should use SUMPRODUCT with multiplication
    assert "SUMPRODUCT" in formula, "'in' with sum should use SUMPRODUCT"
    assert target_range in formula, "Target range should be in formula"
    print(f"   Formula: {formula}")
    print("✅ 'in' operator with sum works")


def test_operator_in_with_mean():
    """Test 'in' operator with mean operation."""
    print("\n📂 Testing 'in' operator with mean")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$C:$C"}
    target_range = "Sheet1!$D:$D"
    
    filters = [FilterCondition(column="Status", operator="in", values=["Active", "Pending"])]
    formula = gen.generate_from_filter("mean", filters, column_ranges, target_range=target_range)
    
    # Should use SUMPRODUCT division for average
    assert "SUMPRODUCT" in formula, "'in' with mean should use SUMPRODUCT"
    assert "/" in formula, "Mean should use division"
    print(f"   Formula: {formula}")
    print("✅ 'in' operator with mean works")


def test_operator_not_in_with_sum():
    """Test 'not_in' operator with sum operation."""
    print("\n📂 Testing 'not_in' operator with sum")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$C:$C"}
    target_range = "Sheet1!$D:$D"
    
    filters = [FilterCondition(column="Status", operator="not_in", values=["Cancelled"])]
    formula = gen.generate_from_filter("sum", filters, column_ranges, target_range=target_range)
    
    # Should return NA() - not supported
    assert "NA()" in formula, "'not_in' with sum should return NA()"
    print(f"   Formula: {formula}")
    print("✅ 'not_in' operator with sum handled correctly")


def test_operator_contains_with_sum():
    """Test 'contains' operator with sum operation."""
    print("\n📂 Testing 'contains' operator with sum")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    target_range = "Sheet1!$B:$B"
    
    filters = [FilterCondition(column="Name", operator="contains", value="John")]
    formula = gen.generate_from_filter("sum", filters, column_ranges, target_range=target_range)
    
    # Should use SUMIF with wildcards
    assert "SUMIF" in formula, "'contains' with sum should use SUMIF"
    assert "*John*" in formula, "Should use wildcards"
    print(f"   Formula: {formula}")
    print("✅ 'contains' operator with sum works")


def test_operator_contains_with_mean():
    """Test 'contains' operator with mean operation."""
    print("\n📂 Testing 'contains' operator with mean")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    target_range = "Sheet1!$B:$B"
    
    filters = [FilterCondition(column="Name", operator="contains", value="John")]
    formula = gen.generate_from_filter("mean", filters, column_ranges, target_range=target_range)
    
    # Should use AVERAGEIF with wildcards
    assert "AVERAGEIF" in formula, "'contains' with mean should use AVERAGEIF"
    assert "*John*" in formula, "Should use wildcards"
    print(f"   Formula: {formula}")
    print("✅ 'contains' operator with mean works")


def test_operator_startswith_with_sum():
    """Test 'startswith' operator with sum operation."""
    print("\n📂 Testing 'startswith' operator with sum")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    target_range = "Sheet1!$B:$B"
    
    filters = [FilterCondition(column="Name", operator="startswith", value="John")]
    formula = gen.generate_from_filter("sum", filters, column_ranges, target_range=target_range)
    
    # Should use SUMIF with wildcard at end
    assert "SUMIF" in formula, "'startswith' with sum should use SUMIF"
    assert "John*" in formula, "Should use wildcard at end"
    print(f"   Formula: {formula}")
    print("✅ 'startswith' operator with sum works")


def test_operator_startswith_with_mean():
    """Test 'startswith' operator with mean operation."""
    print("\n📂 Testing 'startswith' operator with mean")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    target_range = "Sheet1!$B:$B"
    
    filters = [FilterCondition(column="Name", operator="startswith", value="John")]
    formula = gen.generate_from_filter("mean", filters, column_ranges, target_range=target_range)
    
    # Should use AVERAGEIF with wildcard at end
    assert "AVERAGEIF" in formula, "'startswith' with mean should use AVERAGEIF"
    assert "John*" in formula, "Should use wildcard at end"
    print(f"   Formula: {formula}")
    print("✅ 'startswith' operator with mean works")


def test_operator_endswith_with_sum():
    """Test 'endswith' operator with sum operation."""
    print("\n📂 Testing 'endswith' operator with sum")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    target_range = "Sheet1!$B:$B"
    
    filters = [FilterCondition(column="Name", operator="endswith", value="son")]
    formula = gen.generate_from_filter("sum", filters, column_ranges, target_range=target_range)
    
    # Should use SUMIF with wildcard at start
    assert "SUMIF" in formula, "'endswith' with sum should use SUMIF"
    assert "*son" in formula, "Should use wildcard at start"
    print(f"   Formula: {formula}")
    print("✅ 'endswith' operator with sum works")


def test_operator_endswith_with_mean():
    """Test 'endswith' operator with mean operation."""
    print("\n📂 Testing 'endswith' operator with mean")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Name": "Sheet1!$A:$A"}
    target_range = "Sheet1!$B:$B"
    
    filters = [FilterCondition(column="Name", operator="endswith", value="son")]
    formula = gen.generate_from_filter("mean", filters, column_ranges, target_range=target_range)
    
    # Should use AVERAGEIF with wildcard at start
    assert "AVERAGEIF" in formula, "'endswith' with mean should use AVERAGEIF"
    assert "*son" in formula, "Should use wildcard at start"
    print(f"   Formula: {formula}")
    print("✅ 'endswith' operator with mean works")


def test_operator_is_null_with_sum():
    """Test 'is_null' operator with sum operation."""
    print("\n📂 Testing 'is_null' operator with sum")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Email": "Sheet1!$D:$D"}
    target_range = "Sheet1!$E:$E"
    
    filters = [FilterCondition(column="Email", operator="is_null", value=None)]
    formula = gen.generate_from_filter("sum", filters, column_ranges, target_range=target_range)
    
    # Should return NA() - not supported
    assert "NA()" in formula or formula is None, "'is_null' with sum should return NA() or None"
    print(f"   Formula: {formula}")
    print("✅ 'is_null' operator with sum handled correctly")


def test_operator_is_not_null_with_sum():
    """Test 'is_not_null' operator with sum operation."""
    print("\n📂 Testing 'is_not_null' operator with sum")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Email": "Sheet1!$D:$D"}
    target_range = "Sheet1!$E:$E"
    
    filters = [FilterCondition(column="Email", operator="is_not_null", value=None)]
    formula = gen.generate_from_filter("sum", filters, column_ranges, target_range=target_range)
    
    # Should use SUM
    assert formula == "=SUM(Sheet1!$E:$E)", "'is_not_null' with sum should use SUM"
    print(f"   Formula: {formula}")
    print("✅ 'is_not_null' operator with sum works")


def test_operator_is_not_null_with_mean():
    """Test 'is_not_null' operator with mean operation."""
    print("\n📂 Testing 'is_not_null' operator with mean")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Email": "Sheet1!$D:$D"}
    target_range = "Sheet1!$E:$E"
    
    filters = [FilterCondition(column="Email", operator="is_not_null", value=None)]
    formula = gen.generate_from_filter("mean", filters, column_ranges, target_range=target_range)
    
    # Should use AVERAGE
    assert formula == "=AVERAGE(Sheet1!$E:$E)", "'is_not_null' with mean should use AVERAGE"
    print(f"   Formula: {formula}")
    print("✅ 'is_not_null' operator with mean works")


def test_datetime_in_operator_conversion():
    """Test datetime conversion in 'in' operator."""
    print("\n📂 Testing datetime conversion in 'in' operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Date": "Sheet1!$E:$E"}
    column_types = {"Date": "datetime"}
    
    # String datetime values in 'in' operator should be converted
    filters = [FilterCondition(column="Date", operator="in", values=["2024-01-01", "2024-02-01"])]
    formula = gen.generate_from_filter("count", filters, column_ranges, column_types=column_types)
    
    # Should contain DATE() functions
    assert "DATE(" in formula, "DateTime 'in' operator should use DATE() functions"
    print(f"   Formula: {formula}")
    print("✅ DateTime 'in' operator conversion works")


def test_format_criteria_with_string_comparison():
    """Test _format_criteria with string comparison operators."""
    print("\n📂 Testing _format_criteria with string comparison")
    
    gen = FormulaGenerator("Sheet1")
    
    # Test != with string
    criteria_ne = gen._format_criteria("!=", "test")
    assert criteria_ne == '"<>test"', "!= with string should format correctly"
    print(f"   != string: {criteria_ne}")
    
    # Test > with string
    criteria_gt = gen._format_criteria(">", "abc")
    assert criteria_gt == '">abc"', "> with string should format correctly"
    print(f"   > string: {criteria_gt}")
    
    # Test >= with string
    criteria_gte = gen._format_criteria(">=", "abc")
    assert criteria_gte == '">=abc"', ">= with string should format correctly"
    print(f"   >= string: {criteria_gte}")
    
    print("✅ String comparison criteria formatted correctly")


def test_format_criteria_text_operators_with_numbers():
    """Test _format_criteria with text operators on numeric values."""
    print("\n📂 Testing _format_criteria text operators with numbers")
    
    gen = FormulaGenerator("Sheet1")
    
    # Test contains with number
    criteria_contains = gen._format_criteria("contains", 123)
    assert criteria_contains == '"*123*"', "contains with number should convert to string"
    print(f"   contains number: {criteria_contains}")
    
    # Test startswith with number
    criteria_starts = gen._format_criteria("startswith", 456)
    assert criteria_starts == '"456*"', "startswith with number should convert to string"
    print(f"   startswith number: {criteria_starts}")
    
    # Test endswith with number
    criteria_ends = gen._format_criteria("endswith", 789)
    assert criteria_ends == '"*789"', "endswith with number should convert to string"
    print(f"   endswith number: {criteria_ends}")
    
    print("✅ Text operators with numbers formatted correctly")


def test_no_filters_all_operations():
    """Test all operations without filters."""
    print("\n📂 Testing all operations without filters")
    
    gen = FormulaGenerator("Sheet1")
    target_range = "Sheet1!$C:$C"
    
    # Test median
    formula_median = gen.generate_from_filter("median", [], {}, target_range=target_range)
    assert formula_median == "=MEDIAN(Sheet1!$C:$C)", "Median without filters should use MEDIAN"
    
    # Test min
    formula_min = gen.generate_from_filter("min", [], {}, target_range=target_range)
    assert formula_min == "=MIN(Sheet1!$C:$C)", "Min without filters should use MIN"
    
    # Test max
    formula_max = gen.generate_from_filter("max", [], {}, target_range=target_range)
    assert formula_max == "=MAX(Sheet1!$C:$C)", "Max without filters should use MAX"
    
    # Test std
    formula_std = gen.generate_from_filter("std", [], {}, target_range=target_range)
    assert formula_std == "=STDEV(Sheet1!$C:$C)", "Std without filters should use STDEV"
    
    # Test var
    formula_var = gen.generate_from_filter("var", [], {}, target_range=target_range)
    assert formula_var == "=VAR(Sheet1!$C:$C)", "Var without filters should use VAR"
    
    print("✅ All operations without filters work correctly")


def test_unsupported_operation_error():
    """Test error for unsupported operation without filters."""
    print("\n📂 Testing unsupported operation error")
    
    gen = FormulaGenerator("Sheet1")
    
    with pytest.raises(ValueError) as exc_info:
        gen.generate_from_filter("unsupported", [], {})
    
    assert "requires" in str(exc_info.value).lower(), "Should raise error for unsupported operation"
    print(f"   ✅ Error caught: {exc_info.value}")
    print("✅ Unsupported operation error handled correctly")


def test_multiple_filters_unsupported_operation():
    """Test unsupported operation with multiple filters."""
    print("\n📂 Testing unsupported operation with multiple filters")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {
        "City": "Sheet1!$A:$A",
        "Age": "Sheet1!$B:$B"
    }
    
    filters = [
        FilterCondition(column="City", operator="==", value="Moscow"),
        FilterCondition(column="Age", operator=">", value=25)
    ]
    
    # Test mean with multiple filters (not fully supported)
    formula = gen.generate_from_filter("mean", filters, column_ranges)
    
    # Should return NA() message
    assert "NA()" in formula, "Unsupported operation with multiple filters should return NA()"
    print(f"   Formula: {formula}")
    print("✅ Unsupported operation with multiple filters handled correctly")


def test_datetime_not_equal_operator():
    """Test datetime with != operator."""
    print("\n📂 Testing datetime != operator")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Date": "Sheet1!$E:$E"}
    column_types = {"Date": "datetime"}
    
    filters = [FilterCondition(column="Date", operator="!=", value="2024-01-01")]
    formula = gen.generate_from_filter("count", filters, column_ranges, column_types=column_types)
    
    # Should use "<>"&DATE(...)
    assert "<>" in formula and "DATE(" in formula, "!= with datetime should use <>&DATE()"
    print(f"   Formula: {formula}")
    print("✅ DateTime != operator works")


def test_operator_in_empty_values():
    """Test 'in' operator with empty values list."""
    print("\n📂 Testing 'in' operator with empty values")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$C:$C"}
    
    filters = [FilterCondition(column="Status", operator="in", values=[])]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should return NA() - requires values
    assert "NA()" in formula, "'in' with empty values should return NA()"
    print(f"   Formula: {formula}")
    print("✅ 'in' operator with empty values handled correctly")


def test_operator_not_in_empty_values():
    """Test 'not_in' operator with empty values list."""
    print("\n📂 Testing 'not_in' operator with empty values")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$C:$C"}
    
    filters = [FilterCondition(column="Status", operator="not_in", values=[])]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    # Should return NA() - requires values
    assert "NA()" in formula, "'not_in' with empty values should return NA()"
    print(f"   Formula: {formula}")
    print("✅ 'not_in' operator with empty values handled correctly")


# ============================================================================
# NEGATION OPERATOR (NOT) TESTS
# ============================================================================

def test_formula_generation_returns_none_with_negation():
    """Test that formula generation returns None when negate=True."""
    print("\n📂 Testing formula generation with negation")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$A:$A"}
    
    filters = [FilterCondition(column="Status", operator="==", value="Active", negate=True)]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    print(f"   Formula: {formula}")
    
    # Formula should not be generated for negation
    assert formula is None, "Formula should be None when negate=True"
    print("✅ Formula generation correctly returns None for negation")


def test_formula_generation_multiple_filters_with_negation():
    """Test that formula returns None when any filter has negate=True."""
    print("\n📂 Testing formula generation with multiple filters (one negated)")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$A:$A", "Age": "Sheet1!$B:$B"}
    
    filters = [
        FilterCondition(column="Status", operator="==", value="Active"),
        FilterCondition(column="Age", operator=">", value=30, negate=True)
    ]
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    print(f"   Formula: {formula}")
    
    # Formula should not be generated if any filter has negate=True
    assert formula is None, "Formula should be None when any filter has negate=True"
    print("✅ Formula generation correctly returns None when any filter is negated")


def test_convert_datetime_filters_preserves_negate():
    """Test that datetime conversion preserves negate field."""
    print("\n📂 Testing datetime conversion preserves negate")
    
    gen = FormulaGenerator("Sheet1")
    column_types = {"Date": "datetime"}
    
    filters = [FilterCondition(column="Date", operator=">=", value="2024-01-01", negate=True)]
    converted = gen._convert_datetime_filters(filters, column_types)
    
    print(f"   Original negate: {filters[0].negate}")
    print(f"   Converted negate: {converted[0].negate}")
    
    # negate should be preserved after conversion
    assert converted[0].negate == True, "negate field should be preserved after datetime conversion"
    print("✅ Datetime conversion preserves negate field")


# ============================================================================
# NESTED FILTER GROUPS TESTS
# ============================================================================

def test_generate_formula_with_nested_group_returns_none():
    """Test that formula generation returns None for nested groups."""
    print("\n📂 Testing formula generation with nested group")
    
    from mcp_excel.models.requests import FilterGroup
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$A:$A", "Age": "Sheet1!$B:$B"}
    
    # Simple nested group: (Status=Active AND Age>30) OR Status=VIP
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Status", operator="==", value="Active"),
                FilterCondition(column="Age", operator=">", value=30)
            ],
            logic="AND"
        ),
        FilterCondition(column="Status", operator="==", value="VIP")
    ]
    
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    print(f"   Formula: {formula}")
    
    # Nested groups are not supported in Excel formulas
    assert formula is None, "Formula should be None for nested groups"
    print("✅ Formula generation correctly returns None for nested groups")


def test_generate_formula_flat_filters_still_works():
    """Test that flat filters still generate formulas after nested group support."""
    print("\n📂 Testing flat filters still work")
    
    gen = FormulaGenerator("Sheet1")
    column_ranges = {"Status": "Sheet1!$A:$A", "Age": "Sheet1!$B:$B"}
    
    # Flat filters (no groups)
    filters = [
        FilterCondition(column="Status", operator="==", value="Active"),
        FilterCondition(column="Age", operator=">", value=30)
    ]
    
    formula = gen.generate_from_filter("count", filters, column_ranges)
    
    print(f"   Formula: {formula}")
    
    # Flat filters should still generate formulas
    assert formula is not None, "Formula should be generated for flat filters"
    assert "COUNTIFS" in formula, "Should use COUNTIFS for multiple flat conditions"
    print("✅ Flat filters still generate formulas correctly")


def test_convert_datetime_filters_nested_group():
    """Test datetime conversion in nested groups."""
    print("\n📂 Testing datetime conversion in nested groups")
    
    from mcp_excel.models.requests import FilterGroup
    
    gen = FormulaGenerator("Sheet1")
    column_types = {"Date": "datetime", "Status": "string"}
    
    # Nested group with datetime filter
    filters = [
        FilterGroup(
            filters=[
                FilterCondition(column="Date", operator=">=", value="2024-01-01"),
                FilterCondition(column="Status", operator="==", value="Active")
            ],
            logic="AND"
        )
    ]
    
    converted = gen._convert_datetime_filters(filters, column_types)
    
    print(f"   Original type: {type(filters[0])}")
    print(f"   Converted type: {type(converted[0])}")
    
    # Should preserve FilterGroup structure
    assert isinstance(converted[0], FilterGroup), "Should preserve FilterGroup type"
    assert len(converted[0].filters) == 2, "Should have 2 filters in group"
    
    # Check datetime conversion happened inside group
    date_filter = converted[0].filters[0]
    assert isinstance(date_filter.value, pd.Timestamp), "Date value should be converted to Timestamp"
    
    print("✅ Datetime conversion works in nested groups")


def test_convert_datetime_filters_deep_nesting():
    """Test datetime conversion in deeply nested groups (3+ levels)."""
    print("\n📂 Testing datetime conversion in deep nesting")
    
    from mcp_excel.models.requests import FilterGroup
    
    gen = FormulaGenerator("Sheet1")
    column_types = {"Date": "datetime", "Status": "string", "Amount": "float"}
    
    # Deep nesting: ((Date >= X AND Status = Y) OR Amount > Z)
    filters = [
        FilterGroup(
            filters=[
                FilterGroup(
                    filters=[
                        FilterCondition(column="Date", operator=">=", value="2024-01-01"),
                        FilterCondition(column="Status", operator="==", value="Active")
                    ],
                    logic="AND"
                ),
                FilterCondition(column="Amount", operator=">", value=1000)
            ],
            logic="OR"
        )
    ]
    
    converted = gen._convert_datetime_filters(filters, column_types)
    
    print(f"   Nesting levels: 3")
    print(f"   Converted successfully: {isinstance(converted[0], FilterGroup)}")
    
    # Navigate to the datetime filter at level 3
    outer_group = converted[0]
    inner_group = outer_group.filters[0]
    date_filter = inner_group.filters[0]
    
    assert isinstance(date_filter.value, pd.Timestamp), "Date value should be converted at deep nesting level"
    print("✅ Datetime conversion works at deep nesting levels")
