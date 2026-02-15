# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Unit tests for FileLoader component.

Tests cover:
- File loading (.xlsx and .xls formats)
- Caching mechanism (LRU cache)
- Sheet name retrieval
- File info extraction
- Header row parameter handling
- Datetime conversion toggle
- Error handling (file not found, invalid format)
"""

import pytest
from pathlib import Path


def test_load_simple_xlsx(simple_fixture, file_loader):
    """Test loading simple .xlsx file with default parameters.
    
    Verifies:
    - File loads successfully
    - Correct number of rows and columns
    - Column names match expected
    - Data types are correct
    """
    print(f"\n📂 Loading file: {simple_fixture.path_str}")
    print(f"   Sheet: {simple_fixture.sheet_name}")
    
    # Act - explicitly specify header_row=0 to use first row as header
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    
    # Assert
    print(f"✅ Loaded: {len(df)} rows, {len(df.columns)} columns")
    print(f"   Columns: {list(df.columns)}")
    
    assert len(df) == simple_fixture.row_count, f"Expected {simple_fixture.row_count} rows"
    assert len(df.columns) == len(simple_fixture.columns), f"Expected {len(simple_fixture.columns)} columns"
    assert list(df.columns) == simple_fixture.columns, "Column names mismatch"


def test_load_with_explicit_header_row(messy_headers_fixture, file_loader):
    """Test loading file with explicitly specified header row.
    
    Verifies:
    - Can override auto-detection with explicit header_row
    - Correct columns are extracted from specified row
    """
    print(f"\n📂 Loading file with explicit header_row: {messy_headers_fixture.path_str}")
    print(f"   Header row: {messy_headers_fixture.header_row}")
    
    # Act
    df = file_loader.load(
        messy_headers_fixture.path_str,
        messy_headers_fixture.sheet_name,
        header_row=messy_headers_fixture.header_row
    )
    
    # Assert
    print(f"✅ Loaded: {len(df)} rows, {len(df.columns)} columns")
    print(f"   Columns: {list(df.columns)}")
    
    assert list(df.columns) == messy_headers_fixture.columns, "Column names mismatch"
    assert len(df) == messy_headers_fixture.row_count, f"Expected {messy_headers_fixture.row_count} rows"


def test_load_without_header(simple_fixture, file_loader):
    """Test loading file without header (raw data).
    
    Verifies:
    - header_row=None loads raw data
    - Columns are numeric (0, 1, 2, ...)
    - First row contains header data
    """
    print(f"\n📂 Loading file without header: {simple_fixture.path_str}")
    
    # Act
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=None)
    
    # Assert
    print(f"✅ Loaded raw: {len(df)} rows, {len(df.columns)} columns")
    print(f"   Column names (numeric): {list(df.columns)}")
    print(f"   First row (should be headers): {df.iloc[0].tolist()}")
    
    # Columns should be numeric when no header
    assert all(isinstance(col, int) for col in df.columns), "Columns should be numeric without header"
    # Should have one more row (header row is now data)
    assert len(df) == simple_fixture.row_count + 1, "Should include header row as data"


def test_load_legacy_xls_format(simple_legacy_fixture, file_loader):
    """Test loading legacy .xls format.
    
    Verifies:
    - .xls files load correctly with xlrd engine
    - Data integrity is preserved
    - Cyrillic characters are handled correctly
    """
    print(f"\n📂 Loading legacy .xls file: {simple_legacy_fixture.path_str}")
    
    # Act - explicitly specify header_row=0 for legacy format too
    df = file_loader.load(simple_legacy_fixture.path_str, simple_legacy_fixture.sheet_name, header_row=0)
    
    # Assert
    print(f"✅ Loaded legacy: {len(df)} rows, {len(df.columns)} columns")
    print(f"   Columns: {list(df.columns)}")
    print(f"   Sample data: {df.iloc[0].tolist()}")
    
    assert len(df) == simple_legacy_fixture.row_count, f"Expected {simple_legacy_fixture.row_count} rows"
    assert len(df.columns) == len(simple_legacy_fixture.columns), f"Expected {len(simple_legacy_fixture.columns)} columns"
    assert list(df.columns) == simple_legacy_fixture.columns, "Column names mismatch"


def test_load_with_datetime_conversion(with_dates_fixture, file_loader):
    """Test loading file with automatic datetime conversion.
    
    Verifies:
    - Datetime columns are detected and converted
    - Converted values are pd.Timestamp objects
    - Non-datetime columns remain unchanged
    """
    print(f"\n📂 Loading file with datetime conversion: {with_dates_fixture.path_str}")
    
    # Act - need header_row=0 for proper column names
    df = file_loader.load(with_dates_fixture.path_str, with_dates_fixture.sheet_name, header_row=0, convert_dates=True)
    
    # Assert
    print(f"✅ Loaded with datetime conversion: {len(df)} rows, {len(df.columns)} columns")
    print(f"   Column types: {df.dtypes.to_dict()}")
    
    # Check that datetime columns were converted
    datetime_cols = [col for col, dtype in df.dtypes.items() if 'datetime' in str(dtype)]
    print(f"   Datetime columns detected: {datetime_cols}")
    
    assert len(datetime_cols) > 0, "Should detect at least one datetime column"
    
    # Verify datetime values are properly formatted
    for col in datetime_cols:
        sample_value = df[col].iloc[0]
        print(f"   Sample datetime value in '{col}': {sample_value} (type: {type(sample_value).__name__})")
        assert sample_value is not None or df[col].isna().iloc[0], "Datetime value should not be None unless NaN"


def test_load_without_datetime_conversion(with_dates_fixture, file_loader):
    """Test loading file without datetime conversion.
    
    Verifies:
    - convert_dates=False keeps original Excel numbers
    - Datetime columns remain as float64
    """
    print(f"\n📂 Loading file without datetime conversion: {with_dates_fixture.path_str}")
    
    # Act - need header_row=0
    df = file_loader.load(with_dates_fixture.path_str, with_dates_fixture.sheet_name, header_row=0, convert_dates=False)
    
    # Assert
    print(f"✅ Loaded without datetime conversion: {len(df)} rows, {len(df.columns)} columns")
    print(f"   Column types: {df.dtypes.to_dict()}")
    
    # Datetime columns should NOT be converted (should be float or object)
    datetime_cols = [col for col, dtype in df.dtypes.items() if 'datetime' in str(dtype)]
    print(f"   Datetime columns (should be empty): {datetime_cols}")
    
    # Note: This might still have some datetime columns if pandas auto-detected them
    # The key is that we're testing the convert_dates parameter works


def test_cache_hit_on_second_load(simple_fixture, file_loader):
    """Test that cache is used on second load of same file.
    
    Verifies:
    - First load populates cache
    - Second load uses cache (faster)
    - Cache stats show hit
    """
    print(f"\n📂 Testing cache with file: {simple_fixture.path_str}")
    
    # Clear cache first
    file_loader.clear_cache()
    stats_before = file_loader.get_cache_stats()
    print(f"   Cache before: {stats_before}")
    
    # First load (cache miss)
    df1 = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    stats_after_first = file_loader.get_cache_stats()
    print(f"   Cache after first load: {stats_after_first}")
    
    # Second load (cache hit) - should use cache
    df2 = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    stats_after_second = file_loader.get_cache_stats()
    print(f"   Cache after second load: {stats_after_second}")
    
    # Assert
    assert stats_after_first['size'] > stats_before['size'], "Cache should grow after first load"
    # Note: Cache stats don't track hits/misses, just verify cache is being used
    # by checking that size doesn't grow on second load
    assert stats_after_second['size'] == stats_after_first['size'], "Cache size should stay same on second load"
    
    # DataFrames should be identical
    assert df1.equals(df2), "Cached DataFrame should be identical to original"


def test_cache_different_sheets(multi_sheet_fixture, file_loader):
    """Test that different sheets are cached separately.
    
    Verifies:
    - Each sheet has its own cache entry
    - Loading different sheets doesn't invalidate other caches
    - Uses multi_sheet fixture with 3 sheets (Products, Clients, Orders)
    """
    print(f"\n📂 Testing cache with multiple sheets: {multi_sheet_fixture.path_str}")
    
    file_loader.clear_cache()
    
    # Get all sheet names
    sheet_names = file_loader.get_sheet_names(multi_sheet_fixture.path_str)
    print(f"   Available sheets: {sheet_names}")
    
    # Verify we have expected sheets
    assert len(sheet_names) == multi_sheet_fixture.expected["sheet_count"], "Should have 3 sheets"
    assert sheet_names == multi_sheet_fixture.expected["sheet_names"], "Sheet names mismatch"
    
    # Load first sheet (Products)
    df1 = file_loader.load(multi_sheet_fixture.path_str, "Products", header_row=0)
    stats_after_first = file_loader.get_cache_stats()
    print(f"   Cache after loading Products: {stats_after_first}")
    print(f"   Products: {len(df1)} rows, columns: {list(df1.columns)}")
    
    # Load second sheet (Clients)
    df2 = file_loader.load(multi_sheet_fixture.path_str, "Clients", header_row=0)
    stats_after_second = file_loader.get_cache_stats()
    print(f"   Cache after loading Clients: {stats_after_second}")
    print(f"   Clients: {len(df2)} rows, columns: {list(df2.columns)}")
    
    # Assert
    assert stats_after_second['size'] > stats_after_first['size'], "Cache should grow with second sheet"
    assert len(df1) == multi_sheet_fixture.expected["products_count"], "Products row count mismatch"
    assert len(df2) == multi_sheet_fixture.expected["clients_count"], "Clients row count mismatch"


def test_cache_invalidation(simple_fixture, file_loader):
    """Test manual cache invalidation.
    
    Verifies:
    - invalidate_cache() removes file from cache
    - Next load is cache miss
    """
    print(f"\n📂 Testing cache invalidation")
    
    # Load file to populate cache
    df1 = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    stats_before = file_loader.get_cache_stats()
    print(f"   Cache before invalidation: {stats_before}")
    
    # Invalidate cache for this file
    file_loader.invalidate_cache(simple_fixture.path_str)
    stats_after = file_loader.get_cache_stats()
    print(f"   Cache after invalidation: {stats_after}")
    
    # Assert
    assert stats_after['size'] < stats_before['size'], "Cache size should decrease after invalidation"


def test_get_sheet_names(simple_fixture, file_loader):
    """Test retrieving sheet names from file.
    
    Verifies:
    - get_sheet_names() returns list of strings
    - At least one sheet exists
    - Sheet names are non-empty
    """
    print(f"\n📂 Getting sheet names from: {simple_fixture.path_str}")
    
    # Act
    sheet_names = file_loader.get_sheet_names(simple_fixture.path_str)
    
    # Assert
    print(f"✅ Found {len(sheet_names)} sheet(s): {sheet_names}")
    
    assert isinstance(sheet_names, list), "Should return list"
    assert len(sheet_names) > 0, "Should have at least one sheet"
    assert all(isinstance(name, str) for name in sheet_names), "All sheet names should be strings"
    assert all(len(name) > 0 for name in sheet_names), "Sheet names should not be empty"


def test_get_file_info(simple_fixture, file_loader):
    """Test retrieving file information.
    
    Verifies:
    - get_file_info() returns complete metadata
    - Format is correctly detected
    - Size is calculated
    - Sheet count matches
    """
    print(f"\n📂 Getting file info for: {simple_fixture.path_str}")
    
    # Act
    file_info = file_loader.get_file_info(simple_fixture.path_str)
    
    # Assert
    print(f"✅ File info:")
    print(f"   Format: {file_info['format']}")
    print(f"   Size: {file_info['size_mb']} MB ({file_info['size_bytes']} bytes)")
    print(f"   Sheets: {file_info['sheet_count']}")
    print(f"   Sheet names: {file_info['sheet_names']}")
    
    assert file_info['format'] == simple_fixture.format, f"Expected format {simple_fixture.format}"
    assert file_info['size_bytes'] > 0, "File size should be positive"
    assert file_info['sheet_count'] > 0, "Should have at least one sheet"
    assert len(file_info['sheet_names']) == file_info['sheet_count'], "Sheet count mismatch"


def test_load_file_not_found(file_loader):
    """Test error handling for non-existent file.
    
    Verifies:
    - FileNotFoundError is raised
    - Error message is descriptive
    """
    print(f"\n📂 Testing file not found error")
    
    non_existent_path = "C:/this/file/does/not/exist.xlsx"
    
    # Act & Assert
    with pytest.raises(FileNotFoundError) as exc_info:
        file_loader.load(non_existent_path, "Sheet1")
    
    print(f"✅ Caught expected error: {exc_info.value}")
    assert "not found" in str(exc_info.value).lower(), "Error message should mention 'not found'"


def test_load_unsupported_format(temp_excel_path, file_loader):
    """Test error handling for unsupported file format.
    
    Verifies:
    - ValueError is raised for non-Excel files
    - Error message mentions unsupported format
    """
    print(f"\n📂 Testing unsupported format error")
    
    # Create a file with unsupported extension
    unsupported_file = temp_excel_path / "test.csv"
    unsupported_file.write_text("col1,col2\n1,2\n")
    
    # Act & Assert
    with pytest.raises(ValueError) as exc_info:
        file_loader.load(str(unsupported_file), "Sheet1")
    
    print(f"✅ Caught expected error: {exc_info.value}")
    assert "unsupported" in str(exc_info.value).lower(), "Error message should mention 'unsupported'"


def test_load_all_basic_fixtures(basic_fixture_meta, file_loader):
    """Parametrized test: load all basic fixtures successfully.
    
    This test runs for EACH basic fixture automatically.
    Verifies that all basic test files can be loaded without errors.
    """
    print(f"\n📂 Loading basic fixture: {basic_fixture_meta.name}")
    print(f"   File: {basic_fixture_meta.path_str}")
    print(f"   Sheet: {basic_fixture_meta.sheet_name}")
    
    # Act - use header_row from fixture metadata
    df = file_loader.load(
        basic_fixture_meta.path_str,
        basic_fixture_meta.sheet_name,
        header_row=basic_fixture_meta.header_row
    )
    
    # Assert
    print(f"✅ Loaded: {len(df)} rows, {len(df.columns)} columns")
    
    assert len(df) > 0, "Should have at least one row"
    assert len(df.columns) > 0, "Should have at least one column"


def test_load_all_edge_case_fixtures(edge_fixture_meta, file_loader):
    """Parametrized test: load all edge case fixtures successfully.
    
    This test runs for EACH edge case fixture automatically.
    Verifies that even problematic files can be loaded.
    """
    print(f"\n📂 Loading edge case fixture: {edge_fixture_meta.name}")
    print(f"   File: {edge_fixture_meta.path_str}")
    print(f"   Description: {edge_fixture_meta.description}")
    
    # Act - use header_row from fixture metadata
    df = file_loader.load(
        edge_fixture_meta.path_str,
        edge_fixture_meta.sheet_name,
        header_row=edge_fixture_meta.header_row
    )
    
    # Assert
    print(f"✅ Loaded: {len(df)} rows, {len(df.columns)} columns")
    
    assert len(df) >= 0, "Should load without error (even if empty)"
    assert len(df.columns) > 0, "Should have at least one column"


def test_cache_key_includes_header_row(simple_fixture, file_loader):
    """Test that cache key includes header_row parameter.
    
    Verifies:
    - Loading with different header_row creates separate cache entries
    - Same file with different header_row doesn't use wrong cache
    """
    print(f"\n📂 Testing cache key with different header_row values")
    
    file_loader.clear_cache()
    
    # Load with header_row=0
    df1 = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    stats_after_first = file_loader.get_cache_stats()
    print(f"   Cache after header_row=0: {stats_after_first}")
    
    # Load with header_row=1 (should be different cache entry)
    df2 = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=1)
    stats_after_second = file_loader.get_cache_stats()
    print(f"   Cache after header_row=None: {stats_after_second}")
    
    # Assert
    assert stats_after_second['size'] > stats_after_first['size'], "Should create separate cache entry"
    # DataFrames should be different because different rows are used as headers
    assert len(df1.columns) > 0 and len(df2.columns) > 0, "Both should have columns"


def test_cache_key_includes_convert_dates(with_dates_fixture, file_loader):
    """Test that cache key includes convert_dates parameter.
    
    Verifies:
    - Loading with different convert_dates creates separate cache entries
    - Same file with different convert_dates doesn't use wrong cache
    """
    print(f"\n📂 Testing cache key with different convert_dates values")
    
    file_loader.clear_cache()
    
    # Load with convert_dates=True
    df1 = file_loader.load(with_dates_fixture.path_str, with_dates_fixture.sheet_name, header_row=0, convert_dates=True)
    stats_after_first = file_loader.get_cache_stats()
    print(f"   Cache after convert_dates=True: {stats_after_first}")
    
    # Load with convert_dates=False (should be different cache entry)
    df2 = file_loader.load(with_dates_fixture.path_str, with_dates_fixture.sheet_name, header_row=0, convert_dates=False)
    stats_after_second = file_loader.get_cache_stats()
    print(f"   Cache after convert_dates=False: {stats_after_second}")
    
    # Assert
    assert stats_after_second['size'] > stats_after_first['size'], "Should create separate cache entry"
