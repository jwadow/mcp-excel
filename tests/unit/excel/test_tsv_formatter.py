# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Unit tests for TSVFormatter component.

Tests cover:
- Table formatting (headers + rows)
- Single value formatting (with/without formula)
- Key-value pairs formatting
- Matrix formatting (with row/column labels)
- Cell value formatting (strings, numbers, booleans, None)
- Special character escaping (tabs, newlines)
"""

import pytest

from mcp_excel.excel.tsv_formatter import TSVFormatter


# ============================================================================
# TABLE FORMATTING TESTS
# ============================================================================

def test_format_table_simple():
    """Test simple table formatting."""
    print("\n📂 Testing simple table formatting")
    
    formatter = TSVFormatter()
    
    headers = ["Name", "Age", "City"]
    rows = [
        ["Alice", 25, "Moscow"],
        ["Bob", 30, "London"],
        ["Charlie", 35, "Paris"]
    ]
    
    tsv = formatter.format_table(headers, rows)
    
    lines = tsv.split("\n")
    assert len(lines) == 4, "Should have 4 lines (1 header + 3 rows)"
    assert lines[0] == "Name\tAge\tCity", "Header should be tab-separated"
    assert lines[1] == "Alice\t25\tMoscow", "First row should be correct"
    assert lines[2] == "Bob\t30\tLondon", "Second row should be correct"
    
    print(f"✅ Table formatted correctly:\n{tsv}")


def test_format_table_empty_rows():
    """Test table formatting with empty rows list."""
    print("\n📂 Testing table with empty rows")
    
    formatter = TSVFormatter()
    
    headers = ["Name", "Age"]
    rows = []
    
    tsv = formatter.format_table(headers, rows)
    
    assert tsv == "Name\tAge", "Should only have header line"
    print(f"✅ Empty table formatted correctly: {tsv}")


def test_format_table_with_nulls():
    """Test table formatting with None values."""
    print("\n📂 Testing table with None values")
    
    formatter = TSVFormatter()
    
    headers = ["Name", "Email", "Phone"]
    rows = [
        ["Alice", "alice@example.com", None],
        ["Bob", None, "123-456"],
        [None, "unknown@example.com", None]
    ]
    
    tsv = formatter.format_table(headers, rows)
    
    lines = tsv.split("\n")
    assert lines[1] == "Alice\talice@example.com\t", "None should become empty string"
    assert lines[2] == "Bob\t\t123-456", "None in middle should be empty"
    assert lines[3] == "\tunknown@example.com\t", "None at start should be empty"
    
    print(f"✅ Table with nulls formatted correctly:\n{tsv}")


def test_format_table_cyrillic():
    """Test table formatting with Cyrillic data."""
    print("\n📂 Testing table with Cyrillic")
    
    formatter = TSVFormatter()
    
    headers = ["Имя", "Возраст", "Город"]
    rows = [
        ["Алексей", 25, "Москва"],
        ["Мария", 30, "Санкт-Петербург"]
    ]
    
    tsv = formatter.format_table(headers, rows)
    
    assert "Имя\tВозраст\tГород" in tsv, "Cyrillic headers should work"
    assert "Алексей\t25\tМосква" in tsv, "Cyrillic data should work"
    
    print(f"✅ Cyrillic table formatted correctly:\n{tsv}")


# ============================================================================
# SINGLE VALUE FORMATTING TESTS
# ============================================================================

def test_format_single_value_without_formula():
    """Test single value formatting without formula."""
    print("\n📂 Testing single value (no formula)")
    
    formatter = TSVFormatter()
    
    tsv = formatter.format_single_value("Total", 12345)
    
    assert tsv == "Total\t12345", "Should format as label<tab>value"
    print(f"✅ Single value formatted: {tsv}")


def test_format_single_value_with_formula():
    """Test single value formatting with formula."""
    print("\n📂 Testing single value (with formula)")
    
    formatter = TSVFormatter()
    
    tsv = formatter.format_single_value("Total", 12345, formula="=SUM(A1:A10)")
    
    assert tsv == "Total\t=SUM(A1:A10)", "Should use formula instead of value"
    print(f"✅ Single value with formula formatted: {tsv}")


def test_format_single_value_string():
    """Test single value formatting with string value."""
    print("\n📂 Testing single value (string)")
    
    formatter = TSVFormatter()
    
    tsv = formatter.format_single_value("Status", "Active")
    
    assert tsv == "Status\tActive", "String value should work"
    print(f"✅ String value formatted: {tsv}")


# ============================================================================
# KEY-VALUE PAIRS FORMATTING TESTS
# ============================================================================

def test_format_key_value_pairs_simple():
    """Test key-value pairs formatting."""
    print("\n📂 Testing key-value pairs")
    
    formatter = TSVFormatter()
    
    pairs = {
        "Name": "Alice",
        "Age": 25,
        "City": "Moscow"
    }
    
    tsv = formatter.format_key_value_pairs(pairs)
    
    lines = tsv.split("\n")
    assert len(lines) == 3, "Should have 3 lines"
    assert "Name\tAlice" in tsv, "Name pair should be present"
    assert "Age\t25" in tsv, "Age pair should be present"
    assert "City\tMoscow" in tsv, "City pair should be present"
    
    print(f"✅ Key-value pairs formatted:\n{tsv}")


def test_format_key_value_pairs_with_nulls():
    """Test key-value pairs with None values."""
    print("\n📂 Testing key-value pairs with nulls")
    
    formatter = TSVFormatter()
    
    pairs = {
        "Name": "Alice",
        "Email": None,
        "Phone": "123-456"
    }
    
    tsv = formatter.format_key_value_pairs(pairs)
    
    assert "Email\t\n" in tsv or "Email\t" in tsv, "None should become empty"
    print(f"✅ Key-value pairs with nulls formatted:\n{tsv}")


def test_format_key_value_pairs_empty():
    """Test key-value pairs with empty dict."""
    print("\n📂 Testing empty key-value pairs")
    
    formatter = TSVFormatter()
    
    pairs = {}
    
    tsv = formatter.format_key_value_pairs(pairs)
    
    assert tsv == "", "Empty dict should produce empty string"
    print("✅ Empty key-value pairs handled correctly")


# ============================================================================
# MATRIX FORMATTING TESTS
# ============================================================================

def test_format_matrix_simple():
    """Test matrix formatting with row and column labels."""
    print("\n📂 Testing matrix formatting")
    
    formatter = TSVFormatter()
    
    row_labels = ["Row1", "Row2", "Row3"]
    col_labels = ["ColA", "ColB", "ColC"]
    data = [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9]
    ]
    
    tsv = formatter.format_matrix(row_labels, col_labels, data)
    
    lines = tsv.split("\n")
    assert len(lines) == 4, "Should have 4 lines (1 header + 3 data rows)"
    assert lines[0] == "\tColA\tColB\tColC", "Header should have empty first cell"
    assert lines[1] == "Row1\t1\t2\t3", "First data row should be correct"
    assert lines[2] == "Row2\t4\t5\t6", "Second data row should be correct"
    
    print(f"✅ Matrix formatted correctly:\n{tsv}")


def test_format_matrix_correlation():
    """Test matrix formatting for correlation matrix (common use case)."""
    print("\n📂 Testing correlation matrix")
    
    formatter = TSVFormatter()
    
    columns = ["Price", "Quantity", "Discount"]
    data = [
        [1.0, 0.85, -0.3],
        [0.85, 1.0, -0.2],
        [-0.3, -0.2, 1.0]
    ]
    
    tsv = formatter.format_matrix(columns, columns, data)
    
    lines = tsv.split("\n")
    assert "\tPrice\tQuantity\tDiscount" in lines[0], "Column headers should be correct"
    assert "Price\t1.0\t0.85\t-0.3" in lines[1], "First row should be correct"
    
    print(f"✅ Correlation matrix formatted correctly:\n{tsv}")


def test_format_matrix_with_nulls():
    """Test matrix formatting with None values."""
    print("\n📂 Testing matrix with nulls")
    
    formatter = TSVFormatter()
    
    row_labels = ["A", "B"]
    col_labels = ["X", "Y"]
    data = [
        [1, None],
        [None, 2]
    ]
    
    tsv = formatter.format_matrix(row_labels, col_labels, data)
    
    lines = tsv.split("\n")
    assert lines[1] == "A\t1\t", "None should become empty"
    assert lines[2] == "B\t\t2", "None should become empty"
    
    print(f"✅ Matrix with nulls formatted correctly:\n{tsv}")


# ============================================================================
# CELL FORMATTING TESTS
# ============================================================================

def test_format_cell_string():
    """Test cell formatting for strings."""
    print("\n📂 Testing cell formatting (string)")
    
    formatter = TSVFormatter()
    
    assert formatter._format_cell("test") == "test", "Simple string should pass through"
    assert formatter._format_cell("Москва") == "Москва", "Cyrillic should work"
    assert formatter._format_cell("") == "", "Empty string should work"
    
    print("✅ String cell formatting works")


def test_format_cell_numbers():
    """Test cell formatting for numbers."""
    print("\n📂 Testing cell formatting (numbers)")
    
    formatter = TSVFormatter()
    
    assert formatter._format_cell(42) == "42", "Integer should be converted to string"
    assert formatter._format_cell(3.14) == "3.14", "Float should be converted to string"
    assert formatter._format_cell(0) == "0", "Zero should work"
    assert formatter._format_cell(-100) == "-100", "Negative should work"
    
    print("✅ Number cell formatting works")


def test_format_cell_boolean():
    """Test cell formatting for booleans."""
    print("\n📂 Testing cell formatting (boolean)")
    
    formatter = TSVFormatter()
    
    assert formatter._format_cell(True) == "TRUE", "True should become TRUE"
    assert formatter._format_cell(False) == "FALSE", "False should become FALSE"
    
    print("✅ Boolean cell formatting works")


def test_format_cell_none():
    """Test cell formatting for None."""
    print("\n📂 Testing cell formatting (None)")
    
    formatter = TSVFormatter()
    
    assert formatter._format_cell(None) == "", "None should become empty string"
    
    print("✅ None cell formatting works")


def test_format_cell_special_chars():
    """Test cell formatting with special characters."""
    print("\n📂 Testing cell formatting (special chars)")
    
    formatter = TSVFormatter()
    
    # Tabs should be replaced with spaces
    assert formatter._format_cell("text\twith\ttabs") == "text with tabs", "Tabs should be replaced"
    
    # Newlines should be replaced with spaces
    assert formatter._format_cell("text\nwith\nnewlines") == "text with newlines", "Newlines should be replaced"
    
    # Carriage returns should be removed
    assert formatter._format_cell("text\rwith\rCR") == "textwithCR", "CR should be removed"
    
    # Combined
    assert formatter._format_cell("text\t\n\rwith\tall") == "text  with all", "All special chars should be handled"
    
    print("✅ Special character escaping works")


def test_format_cell_mixed_special_chars():
    """Test cell formatting with mixed content."""
    print("\n📂 Testing cell formatting (mixed)")
    
    formatter = TSVFormatter()
    
    # Real-world example: address with newlines
    address = "123 Main St\nApt 4B\nMoscow, Russia"
    formatted = formatter._format_cell(address)
    
    assert "\n" not in formatted, "Newlines should be removed"
    assert "\t" not in formatted, "Tabs should be removed"
    assert "123 Main St Apt 4B Moscow, Russia" == formatted, "Should be single line"
    
    print(f"✅ Mixed content formatted: {formatted}")


# ============================================================================
# EDGE CASES TESTS
# ============================================================================

def test_format_table_single_column():
    """Test table with single column."""
    print("\n📂 Testing single column table")
    
    formatter = TSVFormatter()
    
    headers = ["Value"]
    rows = [[1], [2], [3]]
    
    tsv = formatter.format_table(headers, rows)
    
    lines = tsv.split("\n")
    assert lines[0] == "Value", "Single column header should work"
    assert lines[1] == "1", "Single column row should work"
    
    print(f"✅ Single column table formatted:\n{tsv}")


def test_format_table_single_row():
    """Test table with single row."""
    print("\n📂 Testing single row table")
    
    formatter = TSVFormatter()
    
    headers = ["A", "B", "C"]
    rows = [[1, 2, 3]]
    
    tsv = formatter.format_table(headers, rows)
    
    lines = tsv.split("\n")
    assert len(lines) == 2, "Should have 2 lines (header + 1 row)"
    
    print(f"✅ Single row table formatted:\n{tsv}")


def test_format_table_large_numbers():
    """Test table with large numbers."""
    print("\n📂 Testing large numbers")
    
    formatter = TSVFormatter()
    
    headers = ["ID", "Value"]
    rows = [
        [50089416, 1234567890],
        [99999999, 9876543210]
    ]
    
    tsv = formatter.format_table(headers, rows)
    
    assert "50089416" in tsv, "Large integers should be preserved"
    assert "1234567890" in tsv, "Large integers should be preserved"
    
    print(f"✅ Large numbers formatted correctly:\n{tsv}")


def test_format_table_floats_precision():
    """Test table with float precision."""
    print("\n📂 Testing float precision")
    
    formatter = TSVFormatter()
    
    headers = ["Value"]
    rows = [[3.14159265359], [2.71828182846]]
    
    tsv = formatter.format_table(headers, rows)
    
    # Should preserve full precision
    assert "3.14159265359" in tsv, "Float precision should be preserved"
    
    print(f"✅ Float precision preserved:\n{tsv}")


def test_format_matrix_empty():
    """Test matrix with empty data."""
    print("\n📂 Testing empty matrix")
    
    formatter = TSVFormatter()
    
    row_labels = []
    col_labels = ["A", "B"]
    data = []
    
    tsv = formatter.format_matrix(row_labels, col_labels, data)
    
    # Should only have header row
    assert tsv == "\tA\tB", "Empty matrix should only have header"
    
    print(f"✅ Empty matrix formatted: {tsv}")


def test_unicode_emoji():
    """Test formatting with emoji and special unicode."""
    print("\n📂 Testing emoji and unicode")
    
    formatter = TSVFormatter()
    
    headers = ["Name", "Status"]
    rows = [
        ["Alice", "Active ✓"],
        ["Bob", "VIP 🌟"],
        ["Charlie", "New 🎉"]
    ]
    
    tsv = formatter.format_table(headers, rows)
    
    assert "✓" in tsv, "Checkmark should be preserved"
    assert "🌟" in tsv, "Star emoji should be preserved"
    assert "🎉" in tsv, "Party emoji should be preserved"
    
    print(f"✅ Emoji and unicode preserved:\n{tsv}")


def test_format_cell_object_fallback():
    """Test cell formatting with custom objects (fallback to str())."""
    print("\n📂 Testing object fallback")
    
    formatter = TSVFormatter()
    
    class CustomObject:
        def __str__(self):
            return "CustomValue"
    
    obj = CustomObject()
    formatted = formatter._format_cell(obj)
    
    assert formatted == "CustomValue", "Should use str() for unknown types"
    
    print(f"✅ Object fallback works: {formatted}")
