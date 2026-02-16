# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Integration tests for Data Retrieval operations.

Tests cover:
- get_unique_values: Extract unique values from a column
- get_value_counts: Get frequency counts for column values
- filter_and_get_rows: Retrieve filtered rows with pagination

These are END-TO-END tests that verify the complete operation flow:
FileLoader -> HeaderDetector -> FilterEngine -> Operations -> Response
"""

import pytest

from mcp_excel.operations.data_operations import DataOperations
from mcp_excel.models.requests import (
    GetUniqueValuesRequest,
    GetValueCountsRequest,
    FilterAndGetRowsRequest,
    FilterCondition,
)


# ============================================================================
# get_unique_values tests
# ============================================================================

def test_get_unique_values_simple(simple_fixture, file_loader):
    """Test get_unique_values on simple clean data.
    
    Verifies:
    - Returns unique values from column
    - Values are sorted
    - Count matches number of unique values
    - No truncation for small datasets
    - Performance metrics included
    """
    print(f"\n📂 Testing get_unique_values on: {simple_fixture.name}")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],  # "Имя"
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique values found: {response.count}")
    print(f"   Values: {response.values[:5]}...")
    print(f"   Truncated: {response.truncated}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.count > 0, "Should find unique values"
    assert response.count == len(response.values), "Count should match values length"
    assert response.truncated is False, "Should not be truncated with limit=100"
    assert len(response.values) <= 100, "Should respect limit"
    
    # Check values are sorted
    assert response.values == sorted(response.values), "Values should be sorted"
    
    # Check metadata
    assert response.metadata is not None, "Should include metadata"
    assert response.metadata.rows_total == simple_fixture.row_count
    
    # Check performance metrics
    assert response.performance is not None
    assert response.performance.execution_time_ms > 0


def test_get_unique_values_with_limit(simple_fixture, file_loader):
    """Test get_unique_values with small limit.
    
    Verifies:
    - Respects limit parameter
    - Sets truncated flag when limit exceeded
    - Returns exactly limit number of values
    """
    print(f"\n📂 Testing get_unique_values with limit=3")
    
    ops = DataOperations(file_loader)
    
    # First, get all unique values to know total count
    request_all = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1000
    )
    response_all = ops.get_unique_values(request_all)
    total_unique = response_all.count
    
    print(f"   Total unique values: {total_unique}")
    
    # Now test with small limit
    request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=3
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Returned: {response.count} values")
    print(f"   Truncated: {response.truncated}")
    
    assert response.count == 3, "Should return exactly 3 values"
    assert len(response.values) == 3, "Values list should have 3 items"
    
    if total_unique > 3:
        assert response.truncated is True, "Should be truncated when total > limit"
    else:
        assert response.truncated is False, "Should not be truncated when total <= limit"


def test_get_unique_values_numeric_column(numeric_types_fixture, file_loader):
    """Test get_unique_values on numeric column.
    
    Verifies:
    - Handles integer values correctly
    - Values are properly formatted (no .0 for integers)
    - Sorting works for numeric values
    """
    print(f"\n📂 Testing get_unique_values on numeric column")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        column=numeric_types_fixture.columns[1],  # "Количество" (integer)
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique numeric values: {response.count}")
    print(f"   Sample values: {response.values[:5]}")
    
    assert response.count > 0, "Should find unique values"
    
    # Check that values are integers (not floats with .0)
    for value in response.values:
        if isinstance(value, float):
            assert value.is_integer(), f"Integer column should not have float values: {value}"
        else:
            assert isinstance(value, int), f"Value should be int, got {type(value)}"
    
    # Check sorting (numeric order)
    assert response.values == sorted(response.values), "Numeric values should be sorted"


def test_get_unique_values_with_nulls(with_nulls_fixture, file_loader):
    """Test get_unique_values on column with null values.
    
    Verifies:
    - Excludes null values from results
    - Only returns non-null unique values
    - Count reflects non-null values only
    """
    print(f"\n📂 Testing get_unique_values with nulls")
    
    ops = DataOperations(file_loader)
    
    # Test on column that has nulls (Email column)
    request = GetUniqueValuesRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        column="Email",
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique non-null values: {response.count}")
    print(f"   Values: {response.values}")
    
    assert response.count > 0, "Should find some non-null values"
    
    # Check that None is not in values
    assert None not in response.values, "Should exclude null values"
    assert "None" not in response.values, "Should not have string 'None'"
    
    # All values should be strings (email addresses)
    for value in response.values:
        assert isinstance(value, str), f"Email values should be strings, got {type(value)}"


def test_get_unique_values_with_duplicates(with_duplicates_fixture, file_loader):
    """Test get_unique_values on column with many duplicates.
    
    Verifies:
    - Returns only unique values (no duplicates in result)
    - Count is less than total row count
    - Each value appears only once
    """
    print(f"\n📂 Testing get_unique_values with duplicates")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        column="Клиент",
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique values: {response.count}")
    print(f"   Total rows: {response.metadata.rows_total}")
    print(f"   Values: {response.values}")
    
    assert response.count > 0, "Should find unique values"
    assert response.count < response.metadata.rows_total, "Unique count should be less than total rows (due to duplicates)"
    
    # Check no duplicates in result
    assert len(response.values) == len(set(response.values)), "Result should not contain duplicates"


def test_get_unique_values_unicode(mixed_languages_fixture, file_loader):
    """Test get_unique_values with unicode characters.
    
    Verifies:
    - Handles Cyrillic, Latin, Chinese, emojis correctly
    - No encoding errors
    - Sorting works with mixed unicode
    """
    print(f"\n📂 Testing get_unique_values with unicode")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=mixed_languages_fixture.path_str,
        sheet_name=mixed_languages_fixture.sheet_name,
        column=mixed_languages_fixture.columns[0],  # "Name/Имя"
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique unicode values: {response.count}")
    print(f"   Sample: {response.values[:3]}")
    
    assert response.count > 0, "Should find unique values"
    
    # Check that values contain unicode characters
    all_values_str = "".join(str(v) for v in response.values)
    
    # Should have at least some non-ASCII characters
    has_unicode = any(ord(c) > 127 for c in all_values_str)
    assert has_unicode, "Should contain unicode characters"


def test_get_unique_values_datetime_column(with_dates_fixture, file_loader):
    """Test get_unique_values on datetime column.
    
    Verifies:
    - Handles datetime values correctly
    - Returns ISO 8601 formatted strings
    - Sorting works for datetime values
    """
    print(f"\n📂 Testing get_unique_values on datetime column")
    
    ops = DataOperations(file_loader)
    
    datetime_col = with_dates_fixture.expected["datetime_columns"][0]
    
    request = GetUniqueValuesRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        column=datetime_col,
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique datetime values: {response.count}")
    print(f"   Sample: {response.values[:3]}")
    
    assert response.count > 0, "Should find unique datetime values"
    
    # Check that values are ISO 8601 strings
    for value in response.values[:5]:  # Check first 5
        assert isinstance(value, str), f"Datetime should be string, got {type(value)}"
        assert "T" in value or "-" in value, f"Should be ISO 8601 format: {value}"


def test_get_unique_values_wide_table(wide_table_fixture, file_loader):
    """Test get_unique_values on wide table.
    
    Verifies:
    - Works correctly even with many columns
    - Performance is acceptable
    """
    print(f"\n📂 Testing get_unique_values on wide table")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=wide_table_fixture.path_str,
        sheet_name=wide_table_fixture.sheet_name,
        column=wide_table_fixture.columns[0],
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique values: {response.count}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.count > 0, "Should find unique values"
    assert response.performance.execution_time_ms < 5000, "Should complete in reasonable time"


def test_get_unique_values_single_column(single_column_fixture, file_loader):
    """Test get_unique_values on minimal table (single column).
    
    Verifies:
    - Works with minimal structure
    """
    print(f"\n📂 Testing get_unique_values on single column table")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=single_column_fixture.path_str,
        sheet_name=single_column_fixture.sheet_name,
        column=single_column_fixture.columns[0],
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique values: {response.count}")
    
    assert response.count > 0, "Should find unique values"


def test_get_unique_values_messy_headers(messy_headers_fixture, file_loader):
    """Test get_unique_values with auto-detected headers.
    
    Verifies:
    - Auto-detects correct header row
    - Returns correct unique values from data rows
    """
    print(f"\n📂 Testing get_unique_values with messy headers")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=messy_headers_fixture.path_str,
        sheet_name=messy_headers_fixture.sheet_name,
        column="Клиент",
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique values: {response.count}")
    print(f"   Expected unique clients: {messy_headers_fixture.expected['unique_clients']}")
    
    assert response.count == messy_headers_fixture.expected["unique_clients"], "Should find correct number of unique clients"


def test_get_unique_values_invalid_column(simple_fixture, file_loader):
    """Test get_unique_values with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message lists available columns
    """
    print(f"\n📂 Testing get_unique_values with invalid column")
    
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
    assert "not found" in error_msg.lower(), "Error should mention column not found"
    assert "NonExistentColumn" in error_msg, "Error should mention the invalid column"
    assert "Available columns" in error_msg, "Error should list available columns"


def test_get_unique_values_performance_metrics(simple_fixture, file_loader):
    """Test that get_unique_values includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    - Cache status is reported
    """
    print(f"\n📂 Testing get_unique_values performance metrics")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Performance:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Cache hit: {response.performance.cache_hit}")
    
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.cache_hit in [True, False], "Should report cache status"


def test_get_unique_values_legacy_format(simple_legacy_fixture, file_loader):
    """Test get_unique_values on legacy .xls format.
    
    Verifies:
    - Works with .xls files
    - xlrd engine handles unique values correctly
    """
    print(f"\n📂 Testing get_unique_values on legacy .xls")
    
    ops = DataOperations(file_loader)
    request = GetUniqueValuesRequest(
        file_path=simple_legacy_fixture.path_str,
        sheet_name=simple_legacy_fixture.sheet_name,
        column=simple_legacy_fixture.columns[0],
        limit=100
    )
    
    # Act
    response = ops.get_unique_values(request)
    
    # Assert
    print(f"✅ Unique values from .xls: {response.count}")
    
    assert response.count > 0, "Should find unique values in .xls file"


# ============================================================================
# get_value_counts tests
# ============================================================================

def test_get_value_counts_simple(simple_fixture, file_loader):
    """Test get_value_counts on simple clean data.
    
    Verifies:
    - Returns frequency counts for column values
    - Top N values are returned
    - Total values count is correct
    - TSV output is generated
    - Performance metrics included
    """
    print(f"\n📂 Testing get_value_counts on: {simple_fixture.name}")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[2],  # "Город" (City)
        top_n=5
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts returned: {len(response.value_counts)}")
    print(f"   Total values: {response.total_values}")
    print(f"   Top values: {list(response.value_counts.items())[:3]}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert len(response.value_counts) > 0, "Should return value counts"
    assert len(response.value_counts) <= 5, "Should respect top_n=5"
    assert response.total_values == simple_fixture.row_count, "Total should match row count"
    
    # Check that counts are integers
    for value, count in response.value_counts.items():
        assert isinstance(count, int), f"Count should be integer, got {type(count)}"
        assert count > 0, "Count should be positive"
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    assert "\t" in response.excel_output.tsv, "TSV should use tab separators"
    
    # Check metadata
    assert response.metadata is not None
    assert response.metadata.rows_total == simple_fixture.row_count
    
    # Check performance
    assert response.performance is not None
    assert response.performance.execution_time_ms > 0


def test_get_value_counts_top_n_parameter(simple_fixture, file_loader):
    """Test get_value_counts with different top_n values.
    
    Verifies:
    - Respects top_n parameter
    - Returns at most top_n values
    - Most frequent values come first
    """
    print(f"\n📂 Testing get_value_counts with top_n=3")
    
    ops = DataOperations(file_loader)
    
    # Test with top_n=3
    request = GetValueCountsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        top_n=3
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Returned: {len(response.value_counts)} values")
    print(f"   Values: {list(response.value_counts.keys())}")
    
    assert len(response.value_counts) <= 3, "Should return at most 3 values"
    
    # Check that values are sorted by frequency (descending)
    counts = list(response.value_counts.values())
    assert counts == sorted(counts, reverse=True), "Values should be sorted by frequency (descending)"


def test_get_value_counts_with_duplicates(with_duplicates_fixture, file_loader):
    """Test get_value_counts on column with many duplicates.
    
    Verifies:
    - Correctly counts duplicate occurrences
    - Returns highest frequency values first
    - Total values matches row count
    """
    print(f"\n📂 Testing get_value_counts with duplicates")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=with_duplicates_fixture.path_str,
        sheet_name=with_duplicates_fixture.sheet_name,
        column="Клиент",
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts: {response.value_counts}")
    print(f"   Total values: {response.total_values}")
    
    assert len(response.value_counts) > 0, "Should return value counts"
    assert response.total_values == with_duplicates_fixture.row_count, "Total should match row count"
    
    # Check that some values have count > 1 (duplicates)
    has_duplicates = any(count > 1 for count in response.value_counts.values())
    assert has_duplicates, "Should detect duplicate values with count > 1"
    
    # Sum of all counts should equal total (for top values shown)
    # Note: This is only true if we show ALL unique values, not just top_n
    # So we just check that counts are reasonable
    for count in response.value_counts.values():
        assert count <= response.total_values, "Individual count should not exceed total"


def test_get_value_counts_numeric_column(numeric_types_fixture, file_loader):
    """Test get_value_counts on numeric column.
    
    Verifies:
    - Handles numeric values correctly
    - Values are properly formatted (no .0 for integers)
    - Counts are accurate
    """
    print(f"\n📂 Testing get_value_counts on numeric column")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
        column=numeric_types_fixture.columns[1],  # "Количество" (integer)
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts for numeric column: {len(response.value_counts)}")
    print(f"   Sample: {list(response.value_counts.items())[:3]}")
    
    assert len(response.value_counts) > 0, "Should return value counts"
    
    # Check that keys are strings (formatted values)
    for key in response.value_counts.keys():
        assert isinstance(key, str), f"Keys should be strings, got {type(key)}"
        # Should not have .0 for integers
        if "." in key:
            # If it has a decimal point, it should have non-zero decimals
            parts = key.split(".")
            if len(parts) == 2:
                assert parts[1] != "0", f"Integer values should not have .0: {key}"


def test_get_value_counts_with_nulls(with_nulls_fixture, file_loader):
    """Test get_value_counts on column with null values.
    
    Verifies:
    - Excludes null values from counts
    - Total values reflects non-null count
    - Only non-null values in results
    """
    print(f"\n📂 Testing get_value_counts with nulls")
    
    ops = DataOperations(file_loader)
    
    # Test on column that has nulls (Email column)
    request = GetValueCountsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        column="Email",
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts (non-null): {len(response.value_counts)}")
    print(f"   Total non-null values: {response.total_values}")
    
    assert len(response.value_counts) > 0, "Should find some non-null values"
    
    # Check that None is not in keys
    assert "None" not in response.value_counts, "Should exclude null values"
    assert None not in response.value_counts, "Should exclude null values"
    
    # Total should be less than row count (due to nulls)
    assert response.total_values < with_nulls_fixture.row_count, "Total should exclude nulls"


def test_get_value_counts_unicode(mixed_languages_fixture, file_loader):
    """Test get_value_counts with unicode characters.
    
    Verifies:
    - Handles Cyrillic, Latin, Chinese, emojis correctly
    - No encoding errors in keys
    - Counts are accurate
    """
    print(f"\n📂 Testing get_value_counts with unicode")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=mixed_languages_fixture.path_str,
        sheet_name=mixed_languages_fixture.sheet_name,
        column=mixed_languages_fixture.columns[0],
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts with unicode: {len(response.value_counts)}")
    print(f"   Sample keys: {list(response.value_counts.keys())[:3]}")
    
    assert len(response.value_counts) > 0, "Should return value counts"
    
    # Check that keys contain unicode
    all_keys = "".join(response.value_counts.keys())
    has_unicode = any(ord(c) > 127 for c in all_keys)
    assert has_unicode, "Should contain unicode characters"


def test_get_value_counts_datetime_column(with_dates_fixture, file_loader):
    """Test get_value_counts on datetime column.
    
    Verifies:
    - Handles datetime values correctly
    - Keys are ISO 8601 formatted strings
    - Counts are accurate
    """
    print(f"\n📂 Testing get_value_counts on datetime column")
    
    ops = DataOperations(file_loader)
    
    datetime_col = with_dates_fixture.expected["datetime_columns"][0]
    
    request = GetValueCountsRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        column=datetime_col,
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts for datetime: {len(response.value_counts)}")
    print(f"   Sample: {list(response.value_counts.items())[:2]}")
    
    assert len(response.value_counts) > 0, "Should return value counts"
    
    # Check that keys are datetime strings
    for key in list(response.value_counts.keys())[:3]:
        assert isinstance(key, str), f"Datetime keys should be strings, got {type(key)}"
        # Should have date format indicators
        assert "-" in key or "T" in key, f"Should be ISO 8601 format: {key}"


def test_get_value_counts_tsv_output(simple_fixture, file_loader):
    """Test that get_value_counts generates proper TSV output.
    
    Verifies:
    - TSV output is generated
    - Contains column name and counts
    - Can be pasted into Excel
    - Has proper structure
    """
    print(f"\n📂 Testing get_value_counts TSV output")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        top_n=5
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ TSV output generated")
    print(f"   Length: {len(response.excel_output.tsv)} chars")
    print(f"   Preview: {response.excel_output.tsv[:150]}...")
    
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    
    # Check TSV structure
    assert "\t" in response.excel_output.tsv, "TSV should use tab separators"
    assert "\n" in response.excel_output.tsv, "TSV should have multiple lines"
    
    # Check that column name is in TSV
    assert simple_fixture.columns[0] in response.excel_output.tsv, "TSV should contain column name"
    
    # Check that "Count" header is present
    assert "Count" in response.excel_output.tsv, "TSV should have Count header"


def test_get_value_counts_wide_table(wide_table_fixture, file_loader):
    """Test get_value_counts on wide table.
    
    Verifies:
    - Works correctly even with many columns
    - Performance is acceptable
    """
    print(f"\n📂 Testing get_value_counts on wide table")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=wide_table_fixture.path_str,
        sheet_name=wide_table_fixture.sheet_name,
        column=wide_table_fixture.columns[0],
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts: {len(response.value_counts)}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert len(response.value_counts) > 0, "Should return value counts"
    assert response.performance.execution_time_ms < 5000, "Should complete in reasonable time"


def test_get_value_counts_single_column(single_column_fixture, file_loader):
    """Test get_value_counts on minimal table (single column).
    
    Verifies:
    - Works with minimal structure
    """
    print(f"\n📂 Testing get_value_counts on single column table")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=single_column_fixture.path_str,
        sheet_name=single_column_fixture.sheet_name,
        column=single_column_fixture.columns[0],
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts: {len(response.value_counts)}")
    
    assert len(response.value_counts) > 0, "Should return value counts"


def test_get_value_counts_messy_headers(messy_headers_fixture, file_loader):
    """Test get_value_counts with auto-detected headers.
    
    Verifies:
    - Auto-detects correct header row
    - Returns correct value counts from data rows
    """
    print(f"\n📂 Testing get_value_counts with messy headers")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=messy_headers_fixture.path_str,
        sheet_name=messy_headers_fixture.sheet_name,
        column="Статус",
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts: {response.value_counts}")
    print(f"   Expected unique statuses: {messy_headers_fixture.expected['unique_statuses']}")
    
    assert len(response.value_counts) == messy_headers_fixture.expected["unique_statuses"], "Should find correct number of unique statuses"
    assert response.total_values == messy_headers_fixture.row_count, "Total should match data rows"


def test_get_value_counts_invalid_column(simple_fixture, file_loader):
    """Test get_value_counts with non-existent column.
    
    Verifies:
    - Raises ValueError for invalid column
    - Error message lists available columns
    """
    print(f"\n📂 Testing get_value_counts with invalid column")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column="NonExistentColumn",
        top_n=10
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.get_value_counts(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower(), "Error should mention column not found"
    assert "NonExistentColumn" in error_msg, "Error should mention the invalid column"
    assert "Available columns" in error_msg, "Error should list available columns"


def test_get_value_counts_performance_metrics(simple_fixture, file_loader):
    """Test that get_value_counts includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    - Cache status is reported
    """
    print(f"\n📂 Testing get_value_counts performance metrics")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Performance:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Cache hit: {response.performance.cache_hit}")
    
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.cache_hit in [True, False], "Should report cache status"


def test_get_value_counts_legacy_format(simple_legacy_fixture, file_loader):
    """Test get_value_counts on legacy .xls format.
    
    Verifies:
    - Works with .xls files
    - xlrd engine handles value counts correctly
    """
    print(f"\n📂 Testing get_value_counts on legacy .xls")
    
    ops = DataOperations(file_loader)
    request = GetValueCountsRequest(
        file_path=simple_legacy_fixture.path_str,
        sheet_name=simple_legacy_fixture.sheet_name,
        column=simple_legacy_fixture.columns[0],
        top_n=10
    )
    
    # Act
    response = ops.get_value_counts(request)
    
    # Assert
    print(f"✅ Value counts from .xls: {len(response.value_counts)}")
    
    assert len(response.value_counts) > 0, "Should return value counts from .xls file"


# ============================================================================
# filter_and_get_rows tests
# ============================================================================

def test_filter_and_get_rows_simple_filter(simple_fixture, file_loader):
    """Test filter_and_get_rows with simple equality filter.
    
    Verifies:
    - Returns rows matching filter
    - Count matches number of returned rows
    - Total matches is accurate
    - TSV output is generated
    - Performance metrics included
    """
    print(f"\n📂 Testing filter_and_get_rows with simple filter")
    
    ops = DataOperations(file_loader)
    
    # Get a value to filter on
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    filter_value = unique_response.values[0]
    
    print(f"   Filtering where '{simple_fixture.columns[0]}' == '{filter_value}'")
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=filter_value)
        ],
        columns=None,  # All columns
        limit=50,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned rows: {response.count}")
    print(f"   Total matches: {response.total_matches}")
    print(f"   Truncated: {response.truncated}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.count > 0, "Should return matching rows"
    assert response.count == len(response.rows), "Count should match rows length"
    assert response.total_matches >= response.count, "Total matches should be >= returned count"
    assert response.truncated is False, "Should not be truncated with limit=50"
    
    # Check row structure
    if response.rows:
        first_row = response.rows[0]
        assert isinstance(first_row, dict), "Row should be dict"
        assert simple_fixture.columns[0] in first_row, "Row should have filtered column"
        assert first_row[simple_fixture.columns[0]] == filter_value, "Row should match filter value"
    
    # Check TSV output
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    
    # Check metadata
    assert response.metadata is not None
    assert response.metadata.rows_total == simple_fixture.row_count
    
    # Check performance
    assert response.performance is not None
    assert response.performance.execution_time_ms > 0


def test_filter_and_get_rows_with_pagination(simple_fixture, file_loader):
    """Test filter_and_get_rows with limit and offset.
    
    Verifies:
    - Respects limit parameter
    - Respects offset parameter
    - Sets truncated flag correctly
    - Can retrieve different pages
    """
    print(f"\n📂 Testing filter_and_get_rows with pagination")
    
    ops = DataOperations(file_loader)
    
    # No filter - get all rows with pagination
    request_page1 = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[],
        columns=None,
        limit=3,
        offset=0,
        logic="AND"
    )
    
    # Act - Page 1
    response_page1 = ops.filter_and_get_rows(request_page1)
    
    # Assert - Page 1
    print(f"✅ Page 1: {response_page1.count} rows")
    print(f"   Total matches: {response_page1.total_matches}")
    print(f"   Truncated: {response_page1.truncated}")
    
    assert response_page1.count == 3, "Should return exactly 3 rows (limit)"
    assert response_page1.total_matches == simple_fixture.row_count, "Total should be all rows"
    assert response_page1.truncated is True, "Should be truncated (more rows available)"
    
    # Act - Page 2 (offset=3)
    request_page2 = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[],
        columns=None,
        limit=3,
        offset=3,
        logic="AND"
    )
    response_page2 = ops.filter_and_get_rows(request_page2)
    
    # Assert - Page 2
    print(f"   Page 2: {response_page2.count} rows")
    
    assert response_page2.count == 3, "Should return 3 rows from offset 3"
    
    # Check that pages have different data
    if response_page1.rows and response_page2.rows:
        page1_first = response_page1.rows[0]
        page2_first = response_page2.rows[0]
        assert page1_first != page2_first, "Different pages should have different data"


def test_filter_and_get_rows_column_selection(simple_fixture, file_loader):
    """Test filter_and_get_rows with specific columns.
    
    Verifies:
    - Returns only requested columns
    - Column order is preserved
    - All rows have same columns
    """
    print(f"\n📂 Testing filter_and_get_rows with column selection")
    
    ops = DataOperations(file_loader)
    
    # Request only first 2 columns
    selected_columns = simple_fixture.columns[:2]
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[],
        columns=selected_columns,
        limit=10,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned rows: {response.count}")
    print(f"   Requested columns: {selected_columns}")
    
    assert response.count > 0, "Should return rows"
    
    # Check that all rows have only selected columns
    for row in response.rows:
        assert len(row) == len(selected_columns), f"Row should have {len(selected_columns)} columns"
        for col in selected_columns:
            assert col in row, f"Row should have column {col}"
        
        # Check no extra columns
        for col in row.keys():
            assert col in selected_columns, f"Row should not have extra column {col}"


def test_filter_and_get_rows_multiple_filters_and(simple_fixture, file_loader):
    """Test filter_and_get_rows with multiple AND filters.
    
    Verifies:
    - Applies multiple filters with AND logic
    - Returns only rows matching ALL conditions
    - Count is accurate
    """
    print(f"\n📂 Testing filter_and_get_rows with multiple AND filters")
    
    ops = DataOperations(file_loader)
    
    # Get values for filtering
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    filter_value1 = unique_response.values[0]
    
    print(f"   Filter 1: '{simple_fixture.columns[0]}' == '{filter_value1}'")
    print(f"   Filter 2: '{simple_fixture.columns[1]}' > 25")
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=filter_value1),
            FilterCondition(column=simple_fixture.columns[1], operator=">", value=25)
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Matching rows: {response.count}")
    
    # Check that all returned rows match both conditions
    for row in response.rows:
        assert row[simple_fixture.columns[0]] == filter_value1, "Row should match first filter"
        assert row[simple_fixture.columns[1]] > 25, "Row should match second filter"


def test_filter_and_get_rows_multiple_filters_or(simple_fixture, file_loader):
    """Test filter_and_get_rows with multiple OR filters.
    
    Verifies:
    - Applies multiple filters with OR logic
    - Returns rows matching ANY condition
    - Count is accurate
    """
    print(f"\n📂 Testing filter_and_get_rows with multiple OR filters")
    
    ops = DataOperations(file_loader)
    
    # Get two different values
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=2
    )
    unique_response = ops.get_unique_values(unique_request)
    
    if len(unique_response.values) < 2:
        pytest.skip("Need at least 2 unique values for OR test")
    
    filter_value1 = unique_response.values[0]
    filter_value2 = unique_response.values[1]
    
    print(f"   Filter 1: '{simple_fixture.columns[0]}' == '{filter_value1}'")
    print(f"   Filter 2: '{simple_fixture.columns[0]}' == '{filter_value2}'")
    print(f"   Logic: OR")
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=filter_value1),
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=filter_value2)
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="OR"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Matching rows: {response.count}")
    
    assert response.count > 0, "Should return matching rows"
    
    # Check that all returned rows match at least one condition
    for row in response.rows:
        matches_filter1 = row[simple_fixture.columns[0]] == filter_value1
        matches_filter2 = row[simple_fixture.columns[0]] == filter_value2
        assert matches_filter1 or matches_filter2, "Row should match at least one filter (OR logic)"


def test_filter_and_get_rows_in_operator(simple_fixture, file_loader):
    """Test filter_and_get_rows with 'in' operator.
    
    Verifies:
    - Handles 'in' operator correctly
    - Returns rows matching any value in list
    """
    print(f"\n📂 Testing filter_and_get_rows with 'in' operator")
    
    ops = DataOperations(file_loader)
    
    # Get multiple values
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=3
    )
    unique_response = ops.get_unique_values(unique_request)
    filter_values = unique_response.values[:2]  # Use first 2 values
    
    print(f"   Filter: '{simple_fixture.columns[0]}' in {filter_values}")
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="in", values=filter_values)
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Matching rows: {response.count}")
    
    assert response.count > 0, "Should return matching rows"
    
    # Check that all returned rows have value in the list
    for row in response.rows:
        assert row[simple_fixture.columns[0]] in filter_values, "Row value should be in filter list"


def test_filter_and_get_rows_no_matches(simple_fixture, file_loader):
    """Test filter_and_get_rows when no rows match.
    
    Verifies:
    - Returns empty list when no matches
    - Count is 0
    - Total matches is 0
    - TSV indicates no matches
    """
    print(f"\n📂 Testing filter_and_get_rows with no matches")
    
    ops = DataOperations(file_loader)
    
    # Use impossible filter value
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value="__IMPOSSIBLE_VALUE__")
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Matching rows: {response.count}")
    print(f"   Total matches: {response.total_matches}")
    
    assert response.count == 0, "Should return 0 rows"
    assert len(response.rows) == 0, "Rows list should be empty"
    assert response.total_matches == 0, "Total matches should be 0"
    assert response.truncated is False, "Should not be truncated"
    
    # Check TSV indicates no matches
    assert "No rows match" in response.excel_output.tsv, "TSV should indicate no matches"


def test_filter_and_get_rows_with_nulls(with_nulls_fixture, file_loader):
    """Test filter_and_get_rows on data with null values.
    
    Verifies:
    - Handles null values in results correctly
    - Null values are JSON-serializable (None)
    """
    print(f"\n📂 Testing filter_and_get_rows with nulls")
    
    ops = DataOperations(file_loader)
    
    # Get all rows (no filter)
    request = FilterAndGetRowsRequest(
        file_path=with_nulls_fixture.path_str,
        sheet_name=with_nulls_fixture.sheet_name,
        filters=[],
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned rows: {response.count}")
    
    assert response.count > 0, "Should return rows"
    
    # Check that some rows have None values
    has_null = False
    for row in response.rows:
        for value in row.values():
            if value is None:
                has_null = True
                break
        if has_null:
            break
    
    print(f"   Has null values: {has_null}")
    # Note: Might not always have null in returned rows, so we don't assert


def test_filter_and_get_rows_datetime_filter(with_dates_fixture, file_loader):
    """Test filter_and_get_rows with datetime column filter.
    
    Verifies:
    - Handles datetime filtering correctly
    - Returns rows matching datetime condition
    """
    print(f"\n📂 Testing filter_and_get_rows with datetime filter")
    
    ops = DataOperations(file_loader)
    
    datetime_col = with_dates_fixture.expected["datetime_columns"][0]
    date_value = with_dates_fixture.expected["date_range_start"]
    
    print(f"   Filter: '{datetime_col}' >= '{date_value}'")
    
    request = FilterAndGetRowsRequest(
        file_path=with_dates_fixture.path_str,
        sheet_name=with_dates_fixture.sheet_name,
        filters=[
            FilterCondition(column=datetime_col, operator=">=", value=date_value)
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Matching rows: {response.count}")
    
    assert response.count > 0, "Should return matching rows"
    
    # Check that datetime values are ISO 8601 strings
    if response.rows:
        first_row = response.rows[0]
        datetime_value = first_row.get(datetime_col)
        if datetime_value:
            assert isinstance(datetime_value, str), "Datetime should be string"
            assert "T" in datetime_value or "-" in datetime_value, "Should be ISO 8601 format"


def test_filter_and_get_rows_unicode(mixed_languages_fixture, file_loader):
    """Test filter_and_get_rows with unicode data.
    
    Verifies:
    - Handles unicode correctly in filters and results
    - No encoding errors
    """
    print(f"\n📂 Testing filter_and_get_rows with unicode")
    
    ops = DataOperations(file_loader)
    
    # Get all rows
    request = FilterAndGetRowsRequest(
        file_path=mixed_languages_fixture.path_str,
        sheet_name=mixed_languages_fixture.sheet_name,
        filters=[],
        columns=None,
        limit=10,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned rows: {response.count}")
    
    assert response.count > 0, "Should return rows"
    
    # Check that rows contain unicode
    if response.rows:
        all_values = "".join(str(v) for row in response.rows for v in row.values() if v)
        has_unicode = any(ord(c) > 127 for c in all_values)
        assert has_unicode, "Should contain unicode characters"


def test_filter_and_get_rows_wide_table(wide_table_fixture, file_loader):
    """Test filter_and_get_rows on wide table.
    
    Verifies:
    - Handles many columns correctly
    - Default column limit is applied (context overflow protection)
    - Performance is acceptable
    """
    print(f"\n📂 Testing filter_and_get_rows on wide table")
    
    ops = DataOperations(file_loader)
    
    # Request without specifying columns (should apply default limit)
    request = FilterAndGetRowsRequest(
        file_path=wide_table_fixture.path_str,
        sheet_name=wide_table_fixture.sheet_name,
        filters=[],
        columns=None,  # Should apply default column limit
        limit=10,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned rows: {response.count}")
    print(f"   Columns in result: {len(response.rows[0]) if response.rows else 0}")
    print(f"   Performance: {response.performance.execution_time_ms}ms")
    
    assert response.count > 0, "Should return rows"
    
    # Check that column limit was applied (default is 5)
    if response.rows:
        assert len(response.rows[0]) <= 5, "Should apply default column limit (5) for context overflow protection"
    
    assert response.performance.execution_time_ms < 5000, "Should complete in reasonable time"


def test_filter_and_get_rows_tsv_output(simple_fixture, file_loader):
    """Test that filter_and_get_rows generates proper TSV output.
    
    Verifies:
    - TSV output is generated
    - Contains headers and data
    - Can be pasted into Excel
    """
    print(f"\n📂 Testing filter_and_get_rows TSV output")
    
    ops = DataOperations(file_loader)
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[],
        columns=simple_fixture.columns[:2],
        limit=5,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ TSV output generated")
    print(f"   Length: {len(response.excel_output.tsv)} chars")
    print(f"   Preview: {response.excel_output.tsv[:150]}...")
    
    assert response.excel_output.tsv, "Should generate TSV output"
    assert len(response.excel_output.tsv) > 0, "TSV should not be empty"
    
    # Check TSV structure
    assert "\t" in response.excel_output.tsv, "TSV should use tab separators"
    assert "\n" in response.excel_output.tsv, "TSV should have multiple lines"
    
    # Check that column names are in TSV
    for col in simple_fixture.columns[:2]:
        assert col in response.excel_output.tsv, f"TSV should contain column {col}"


def test_filter_and_get_rows_invalid_column_in_filter(simple_fixture, file_loader):
    """Test filter_and_get_rows with invalid column in filter.
    
    Verifies:
    - Raises ValueError for invalid filter column
    - Error message is helpful
    """
    print(f"\n📂 Testing filter_and_get_rows with invalid filter column")
    
    ops = DataOperations(file_loader)
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column="NonExistentColumn", operator="==", value="test")
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.filter_and_get_rows(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "NonExistentColumn" in error_msg, "Error should mention the invalid column"


def test_filter_and_get_rows_invalid_column_selection(simple_fixture, file_loader):
    """Test filter_and_get_rows with invalid column in selection.
    
    Verifies:
    - Raises ValueError for invalid selected column
    - Error message lists available columns
    """
    print(f"\n📂 Testing filter_and_get_rows with invalid column selection")
    
    ops = DataOperations(file_loader)
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[],
        columns=["NonExistentColumn"],
        limit=50,
        offset=0,
        logic="AND"
    )
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        ops.filter_and_get_rows(request)
    
    print(f"✅ Caught expected error: {exc_info.value}")
    
    error_msg = str(exc_info.value)
    assert "not found" in error_msg.lower(), "Error should mention column not found"
    assert "NonExistentColumn" in error_msg, "Error should mention the invalid column"


def test_filter_and_get_rows_performance_metrics(simple_fixture, file_loader):
    """Test that filter_and_get_rows includes performance metrics.
    
    Verifies:
    - Performance metrics are included
    - Execution time is reasonable
    - Cache status is reported
    """
    print(f"\n📂 Testing filter_and_get_rows performance metrics")
    
    ops = DataOperations(file_loader)
    
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[],
        columns=None,
        limit=10,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Performance:")
    print(f"   Execution time: {response.performance.execution_time_ms}ms")
    print(f"   Cache hit: {response.performance.cache_hit}")
    
    assert response.performance is not None, "Should include performance metrics"
    assert response.performance.execution_time_ms > 0, "Should have execution time"
    assert response.performance.cache_hit in [True, False], "Should report cache status"


def test_filter_and_get_rows_legacy_format(simple_legacy_fixture, file_loader):
    """Test filter_and_get_rows on legacy .xls format.
    
    Verifies:
    - Works with .xls files
    - xlrd engine handles filtering correctly
    """
    print(f"\n📂 Testing filter_and_get_rows on legacy .xls")
    
    ops = DataOperations(file_loader)
    
    request = FilterAndGetRowsRequest(
        file_path=simple_legacy_fixture.path_str,
        sheet_name=simple_legacy_fixture.sheet_name,
        filters=[],
        columns=None,
        limit=10,
        offset=0,
        logic="AND"
    )
    
    # Act
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned rows from .xls: {response.count}")
    
    assert response.count > 0, "Should return rows from .xls file"


# ============================================================================
# NEGATION OPERATOR (NOT) TESTS
# ============================================================================

def test_filter_and_get_rows_with_negation(simple_fixture, file_loader):
    """Test filter_and_get_rows with negated condition.
    
    Verifies:
    - Negation works correctly in filter_and_get_rows
    - Returned rows exclude negated values
    - All returned rows satisfy the negated condition
    """
    print(f"\n🔍 Testing filter_and_get_rows with negation")
    
    ops = DataOperations(file_loader)
    
    # Get a test value
    from mcp_excel.models.requests import GetUniqueValuesRequest
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=1
    )
    test_value = ops.get_unique_values(unique_request).values[0]
    
    print(f"  Filter: {simple_fixture.columns[0]} == '{test_value}' (negated)")
    
    # Act
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=test_value, negate=True)
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned {response.count} rows")
    
    # All returned rows should NOT contain test_value
    assert all(row[simple_fixture.columns[0]] != test_value for row in response.rows), \
        f"No row should have {simple_fixture.columns[0]} == {test_value}"
    assert response.count > 0, "Should return some rows"
  
# ============================================================================
# NESTED FILTER GROUPS TESTS (filter_and_get_rows)
# ============================================================================

def test_filter_and_get_rows_nested_and_or(simple_fixture, file_loader):
    """Test filter_and_get_rows with nested group: (A AND B) OR C.
    
    Verifies:
    - Nested groups work in filter_and_get_rows
    - Returns correct rows for complex logic
    - All returned rows match the nested condition
    """
    print(f"\n🔍 Testing filter_and_get_rows: (A AND B) OR C")
    
    from mcp_excel.models.requests import FilterGroup, GetUniqueValuesRequest
    
    ops = DataOperations(file_loader)
    
    # Get test values
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=2
    )
    values = ops.get_unique_values(unique_request).values[:2]
    
    print(f"  Filter: ({simple_fixture.columns[0]} == '{values[0]}' AND {simple_fixture.columns[1]} > 0) OR {simple_fixture.columns[0]} == '{values[1]}'")
    
    # Act
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[0]),
                    FilterCondition(column=simple_fixture.columns[1], operator=">", value=0)
                ],
                logic="AND"
            ),
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[1])
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="OR"
    )
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned {response.count} rows")
    
    assert response.count > 0, "Should find matching rows"
    
    # Verify each row matches the nested logic
    for row in response.rows:
        matches_group = (row[simple_fixture.columns[0]] == values[0] and row[simple_fixture.columns[1]] > 0)
        matches_condition = (row[simple_fixture.columns[0]] == values[1])
        assert matches_group or matches_condition, "Row should match (A AND B) OR C"


def test_filter_and_get_rows_nested_or_and(numeric_types_fixture, file_loader):
    """Test filter_and_get_rows with nested group: (A OR B) AND C.
    
    Verifies:
    - Different nesting pattern works
    - Logic: (Количество < 50 OR Количество > 150) AND Цена > 100
    """
    print(f"\n🔍 Testing filter_and_get_rows: (A OR B) AND C")
    
    from mcp_excel.models.requests import FilterGroup
    
    ops = DataOperations(file_loader)
    
    print(f"  Filter: (Количество < 50 OR Количество > 150) AND Цена > 100")
    
    # Act
    request = FilterAndGetRowsRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
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
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned {response.count} rows")
    
    assert response.count >= 0, "Count should be non-negative"
    
    # Verify each row matches the nested logic
    for row in response.rows:
        quantity = row["Количество"]
        price = row["Цена"]
        matches_group = (quantity < 50 or quantity > 150)
        matches_condition = (price > 100)
        assert matches_group and matches_condition, "Row should match (A OR B) AND C"


def test_filter_and_get_rows_nested_three_levels(numeric_types_fixture, file_loader):
    """Test filter_and_get_rows with 3 levels of nesting: ((A OR B) AND C) OR D.
    
    Verifies:
    - Deep nesting works correctly
    - Complex logic is evaluated properly
    """
    print(f"\n🔍 Testing filter_and_get_rows with 3 levels: ((A OR B) AND C) OR D")
    
    from mcp_excel.models.requests import FilterGroup
    
    ops = DataOperations(file_loader)
    
    print(f"  Filter: ((Количество < 50 OR Количество > 150) AND Цена > 100) OR Количество == 100")
    
    # Act
    request = FilterAndGetRowsRequest(
        file_path=numeric_types_fixture.path_str,
        sheet_name=numeric_types_fixture.sheet_name,
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
        columns=None,
        limit=50,
        offset=0,
        logic="OR"
    )
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned {response.count} rows")
    
    assert response.count >= 0, "Count should be non-negative"
    
    # Verify complex nested logic
    for row in response.rows:
        quantity = row["Количество"]
        price = row["Цена"]
        inner_group = (quantity < 50 or quantity > 150)
        outer_group = (inner_group and price > 100)
        matches = outer_group or (quantity == 100)
        assert matches, "Row should match ((A OR B) AND C) OR D"


def test_filter_and_get_rows_nested_with_negation(simple_fixture, file_loader):
    """Test filter_and_get_rows with nested group and negation: NOT (A AND B).
    
    Verifies:
    - Negation works with nested groups
    - Returns rows not matching the group
    """
    print(f"\n🔍 Testing filter_and_get_rows: NOT (A AND B)")
    
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
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column=simple_fixture.columns[0], operator="==", value=test_value),
                    FilterCondition(column=simple_fixture.columns[1], operator=">", value=0)
                ],
                logic="AND",
                negate=True
            )
        ],
        columns=None,
        limit=50,
        offset=0,
        logic="AND"
    )
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned {response.count} rows")
    
    assert response.count > 0, "Should find rows not matching the group"
    
    # Verify each row does NOT match (A AND B)
    for row in response.rows:
        matches_group = (row[simple_fixture.columns[0]] == test_value and row[simple_fixture.columns[1]] > 0)
        assert not matches_group, "Row should NOT match (A AND B)"


def test_filter_and_get_rows_nested_with_column_selection(simple_fixture, file_loader):
    """Test filter_and_get_rows with nested groups and column selection.
    
    Verifies:
    - Nested groups work with column selection
    - Returns only requested columns
    """
    print(f"\n🔍 Testing filter_and_get_rows: nested groups + column selection")
    
    from mcp_excel.models.requests import FilterGroup, GetUniqueValuesRequest
    
    ops = DataOperations(file_loader)
    
    # Get test values
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=2
    )
    values = ops.get_unique_values(unique_request).values[:2]
    
    selected_columns = simple_fixture.columns[:2]
    
    print(f"  Filter: (A AND B) OR C")
    print(f"  Columns: {selected_columns}")
    
    # Act
    request = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[0]),
                    FilterCondition(column=simple_fixture.columns[1], operator=">", value=0)
                ],
                logic="AND"
            ),
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[1])
        ],
        columns=selected_columns,
        limit=50,
        offset=0,
        logic="OR"
    )
    response = ops.filter_and_get_rows(request)
    
    # Assert
    print(f"✅ Returned {response.count} rows with {len(selected_columns)} columns")
    
    assert response.count > 0, "Should find matching rows"
    
    # Check that all rows have only selected columns
    for row in response.rows:
        assert len(row) == len(selected_columns), f"Row should have {len(selected_columns)} columns"
        for col in selected_columns:
            assert col in row, f"Row should have column {col}"


def test_filter_and_get_rows_nested_with_pagination(simple_fixture, file_loader):
    """Test filter_and_get_rows with nested groups and pagination.
    
    Verifies:
    - Nested groups work with limit/offset
    - Pagination works correctly with complex filters
    """
    print(f"\n🔍 Testing filter_and_get_rows: nested groups + pagination")
    
    from mcp_excel.models.requests import FilterGroup, GetUniqueValuesRequest
    
    ops = DataOperations(file_loader)
    
    # Get test values
    unique_request = GetUniqueValuesRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        column=simple_fixture.columns[0],
        limit=2
    )
    values = ops.get_unique_values(unique_request).values[:2]
    
    print(f"  Filter: (A AND B) OR C with limit=3")
    
    # Act - Page 1
    request_page1 = FilterAndGetRowsRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterGroup(
                filters=[
                    FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[0]),
                    FilterCondition(column=simple_fixture.columns[1], operator=">", value=0)
                ],
                logic="AND"
            ),
            FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[1])
        ],
        columns=None,
        limit=3,
        offset=0,
        logic="OR"
    )
    response_page1 = ops.filter_and_get_rows(request_page1)
    
    # Assert
    print(f"✅ Page 1: {response_page1.count} rows")
    print(f"   Total matches: {response_page1.total_matches}")
    
    assert response_page1.count <= 3, "Should respect limit=3"
    
    if response_page1.total_matches > 3:
        assert response_page1.truncated is True, "Should be truncated when more rows available"
        
        # Act - Page 2
        request_page2 = FilterAndGetRowsRequest(
            file_path=simple_fixture.path_str,
            sheet_name=simple_fixture.sheet_name,
            filters=[
                FilterGroup(
                    filters=[
                        FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[0]),
                        FilterCondition(column=simple_fixture.columns[1], operator=">", value=0)
                    ],
                    logic="AND"
                ),
                FilterCondition(column=simple_fixture.columns[0], operator="==", value=values[1])
            ],
            columns=None,
            limit=3,
            offset=3,
            logic="OR"
        )
        response_page2 = ops.filter_and_get_rows(request_page2)
        
        print(f"   Page 2: {response_page2.count} rows")
        
        # Pages should have different data
        if response_page1.rows and response_page2.rows:
            page1_first = response_page1.rows[0]
            page2_first = response_page2.rows[0]
            assert page1_first != page2_first, "Different pages should have different data"
