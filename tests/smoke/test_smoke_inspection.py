# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Smoke tests for inspection tools.

Tests for 5 inspection tools:
- inspect_file
- get_sheet_info
- get_column_names
- get_data_profile
- find_column

Each tool is tested with:
- Basic functionality (happy path)
- Different parameter combinations
- Edge cases
- Full response structure validation
- Excel output validation where applicable
"""

import pytest


# ============================================================================
# INSPECT_FILE TESTS
# ============================================================================

def test_inspect_file_basic(mcp_call_tool, simple_fixture):
    """Smoke: inspect_file returns complete file structure information."""
    print(f"\n📂 Testing inspect_file with {simple_fixture.file_name}...")
    
    result = mcp_call_tool("inspect_file", {
        "file_path": str(simple_fixture.path_str)
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from InspectFileResponse
    assert "format" in result, "Missing 'format' field"
    assert "size_bytes" in result, "Missing 'size_bytes' field"
    assert "size_mb" in result, "Missing 'size_mb' field"
    assert "sheet_count" in result, "Missing 'sheet_count' field"
    assert "sheet_names" in result, "Missing 'sheet_names' field"
    assert "sheets_info" in result, "Missing 'sheets_info' field"
    assert "metadata" in result, "Missing 'metadata' field"
    assert "performance" in result, "Missing 'performance' field"
    
    # Verify format
    assert result["format"] in ["xlsx", "xls"], f"Invalid format: {result['format']}"
    assert result["format"] == "xlsx", f"Expected xlsx, got {result['format']}"
    
    # Verify size
    assert result["size_bytes"] > 0, "File size should be positive"
    assert result["size_mb"] >= 0, "File size in MB should be non-negative"
    # Verify size conversion (allow small rounding differences)
    expected_mb = result["size_bytes"] / (1024 * 1024)
    assert abs(result["size_mb"] - expected_mb) < 0.01, f"Size conversion incorrect: {result['size_mb']} != {expected_mb}"
    
    # Verify sheets
    assert result["sheet_count"] >= 1, "Should have at least one sheet"
    assert len(result["sheet_names"]) == result["sheet_count"], "Sheet names count mismatch"
    assert isinstance(result["sheet_names"], list), "sheet_names should be list"
    assert all(isinstance(name, str) for name in result["sheet_names"]), "All sheet names should be strings"
    
    # Verify sheets_info
    assert isinstance(result["sheets_info"], list), "sheets_info should be list"
    assert len(result["sheets_info"]) == result["sheet_count"], "sheets_info count mismatch"
    for i, sheet_info in enumerate(result["sheets_info"]):
        assert isinstance(sheet_info, dict), f"Sheet info {i} should be dict"
        # Check for common keys (structure may vary)
        # At minimum should have some identifying info
        assert len(sheet_info) > 0, f"Sheet info {i} should not be empty"
    
    # Verify metadata structure
    metadata = result["metadata"]
    assert "file_format" in metadata, "Metadata missing 'file_format'"
    assert metadata["file_format"] == result["format"], "Metadata format mismatch"
    
    # Verify performance metrics
    performance = result["performance"]
    assert "execution_time_ms" in performance, "Performance missing 'execution_time_ms'"
    assert "rows_processed" in performance, "Performance missing 'rows_processed'"
    assert "cache_hit" in performance, "Performance missing 'cache_hit'"
    assert "memory_used_mb" in performance, "Performance missing 'memory_used_mb'"
    assert performance["execution_time_ms"] >= 0, "Execution time should be non-negative"
    assert isinstance(performance["cache_hit"], bool), "cache_hit should be boolean"
    
    print(f"  ✅ Format: {result['format']}, Sheets: {result['sheet_count']}, Size: {result['size_mb']:.2f} MB")
    print(f"  ✅ Execution time: {performance['execution_time_ms']:.2f}ms, Cache hit: {performance['cache_hit']}")


def test_inspect_file_multi_sheet(mcp_call_tool, multi_sheet_fixture):
    """Smoke: inspect_file correctly reports multiple sheets with details."""
    print(f"\n📂 Testing inspect_file with multi-sheet file...")
    
    result = mcp_call_tool("inspect_file", {
        "file_path": str(multi_sheet_fixture.path_str)
    })
    
    # Verify multiple sheets
    assert result["sheet_count"] == 3, f"Expected 3 sheets, got {result['sheet_count']}"
    assert len(result["sheet_names"]) == 3, f"Expected 3 sheet names, got {len(result['sheet_names'])}"
    assert len(result["sheets_info"]) == 3, f"Expected 3 sheets_info, got {len(result['sheets_info'])}"
    
    # Verify expected sheet names
    expected_sheets = {"Products", "Clients", "Orders"}
    actual_sheets = set(result["sheet_names"])
    assert actual_sheets == expected_sheets, f"Expected {expected_sheets}, got {actual_sheets}"
    
    # Verify each sheet has info (structure may vary, so be flexible)
    for i, sheet_info in enumerate(result["sheets_info"]):
        assert isinstance(sheet_info, dict), f"Sheet info {i} should be dict"
        assert len(sheet_info) > 0, f"Sheet info {i} should not be empty"
        # If it has name/rows/columns, verify them
        if "name" in sheet_info:
            assert sheet_info["name"] in expected_sheets, f"Unexpected sheet: {sheet_info['name']}"
        if "rows" in sheet_info:
            assert sheet_info["rows"] > 0, f"Sheet should have rows"
        if "columns" in sheet_info:
            assert sheet_info["columns"] > 0, f"Sheet should have columns"
    
    print(f"  ✅ Found 3 sheets: {result['sheet_names']}")
    print(f"  Sheets info structure: {result['sheets_info']}")


# ============================================================================
# GET_SHEET_INFO TESTS
# ============================================================================

def test_get_sheet_info_basic(mcp_call_tool, simple_fixture):
    """Smoke: get_sheet_info returns complete sheet information."""
    print(f"\n📋 Testing get_sheet_info...")
    
    result = mcp_call_tool("get_sheet_info", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from GetSheetInfoResponse
    assert "sheet_name" in result, "Missing 'sheet_name'"
    assert "column_names" in result, "Missing 'column_names'"
    assert "column_count" in result, "Missing 'column_count'"
    assert "column_types" in result, "Missing 'column_types'"
    assert "row_count" in result, "Missing 'row_count'"
    assert "data_start_row" in result, "Missing 'data_start_row'"
    assert "sample_rows" in result, "Missing 'sample_rows'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify sheet name
    assert result["sheet_name"] == simple_fixture.sheet_name, "Sheet name mismatch"
    
    # Verify columns
    assert result["column_count"] == len(simple_fixture.columns), f"Expected {len(simple_fixture.columns)} columns, got {result['column_count']}"
    assert len(result["column_names"]) == result["column_count"], "Column names count mismatch"
    assert result["column_names"] == simple_fixture.columns, f"Column names mismatch: expected {simple_fixture.columns}, got {result['column_names']}"
    
    # Verify column types
    assert isinstance(result["column_types"], dict), "column_types should be dict"
    assert len(result["column_types"]) == result["column_count"], "column_types count mismatch"
    for col_name in result["column_names"]:
        assert col_name in result["column_types"], f"Missing type for column '{col_name}'"
        col_type = result["column_types"][col_name]
        assert col_type in ["integer", "float", "string", "datetime", "boolean"], f"Invalid type '{col_type}' for column '{col_name}'"
    
    # Verify rows
    assert result["row_count"] == simple_fixture.row_count, f"Expected {simple_fixture.row_count} rows, got {result['row_count']}"
    assert result["data_start_row"] >= 0, "data_start_row should be non-negative"
    
    # Verify sample rows
    assert isinstance(result["sample_rows"], list), "sample_rows should be list"
    assert len(result["sample_rows"]) <= 3, "Should have at most 3 sample rows"
    assert len(result["sample_rows"]) <= result["row_count"], "Sample rows can't exceed total rows"
    for row in result["sample_rows"]:
        assert isinstance(row, dict), "Each sample row should be dict"
        # Each row should have values for all columns
        for col_name in result["column_names"]:
            assert col_name in row, f"Sample row missing column '{col_name}'"
    
    # Verify metadata
    metadata = result["metadata"]
    assert metadata["sheet_name"] == result["sheet_name"], "Metadata sheet_name mismatch"
    assert metadata["rows_total"] == result["row_count"], "Metadata rows_total mismatch"
    assert metadata["columns_total"] == result["column_count"], "Metadata columns_total mismatch"
    
    print(f"  ✅ Columns: {result['column_count']}, Rows: {result['row_count']}, Start row: {result['data_start_row']}")
    print(f"  Column names: {result['column_names']}")
    print(f"  Column types: {result['column_types']}")


def test_get_sheet_info_with_dates(mcp_call_tool, with_dates_fixture):
    """Smoke: get_sheet_info correctly identifies and handles datetime columns."""
    print(f"\n📋 Testing get_sheet_info with datetime columns...")
    
    result = mcp_call_tool("get_sheet_info", {
        "file_path": str(with_dates_fixture.path_str),
        "sheet_name": with_dates_fixture.sheet_name
    })
    
    # Verify datetime columns are detected
    column_types = result["column_types"]
    print(f"  Column types: {column_types}")
    
    # Should have at least one datetime column
    datetime_columns = [col for col, dtype in column_types.items() if "datetime" in str(dtype).lower()]
    assert len(datetime_columns) > 0, f"Expected datetime columns, got types: {column_types}"
    
    print(f"  ✅ Found {len(datetime_columns)} datetime column(s): {datetime_columns}")
    
    # Verify sample rows have datetime values
    if result["sample_rows"]:
        first_row = result["sample_rows"][0]
        for dt_col in datetime_columns:
            value = first_row[dt_col]
            print(f"    {dt_col}: {value} (type: {type(value).__name__})")


def test_get_sheet_info_with_header_detection(mcp_call_tool, simple_fixture):
    """Smoke: get_sheet_info includes header detection info when auto-detected."""
    print(f"\n📋 Testing get_sheet_info with header auto-detection...")
    
    # Don't specify header_row to trigger auto-detection
    result = mcp_call_tool("get_sheet_info", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name
        # No header_row parameter
    })
    
    # May or may not have header_detection depending on file
    if "header_detection" in result and result["header_detection"] is not None:
        header_detection = result["header_detection"]
        print(f"  Header detection info present")
        
        # Verify header detection structure
        assert "header_row" in header_detection, "header_detection missing 'header_row'"
        assert "confidence" in header_detection, "header_detection missing 'confidence'"
        assert isinstance(header_detection["header_row"], int), "header_row should be int"
        assert 0 <= header_detection["confidence"] <= 1, f"Confidence should be 0-1, got {header_detection['confidence']}"
        
        print(f"  ✅ Header detected at row {header_detection['header_row']} (confidence: {header_detection['confidence']:.2f})")
    else:
        print(f"  ℹ️  No header detection info (file has clean headers)")


# ============================================================================
# GET_COLUMN_NAMES TESTS
# ============================================================================

def test_get_column_names_basic(mcp_call_tool, simple_fixture):
    """Smoke: get_column_names returns complete list of column names."""
    print(f"\n📝 Testing get_column_names...")
    
    result = mcp_call_tool("get_column_names", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from GetColumnNamesResponse
    assert "column_names" in result, "Missing 'column_names'"
    assert "column_count" in result, "Missing 'column_count'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify column names
    assert isinstance(result["column_names"], list), "column_names should be list"
    assert result["column_count"] == len(simple_fixture.columns), f"Expected {len(simple_fixture.columns)} columns, got {result['column_count']}"
    assert len(result["column_names"]) == result["column_count"], "Column names count mismatch"
    assert result["column_names"] == simple_fixture.columns, f"Column names mismatch"
    
    # Verify all column names are strings
    assert all(isinstance(name, str) for name in result["column_names"]), "All column names should be strings"
    assert all(len(name) > 0 for name in result["column_names"]), "All column names should be non-empty"
    
    # Verify metadata
    metadata = result["metadata"]
    assert metadata["sheet_name"] == simple_fixture.sheet_name, "Metadata sheet_name mismatch"
    
    print(f"  ✅ Found {result['column_count']} columns: {result['column_names']}")


# ============================================================================
# GET_DATA_PROFILE TESTS
# ============================================================================

def test_get_data_profile_all_columns(mcp_call_tool, simple_fixture):
    """Smoke: get_data_profile returns comprehensive profiles for all columns."""
    print(f"\n📊 Testing get_data_profile (all columns)...")
    
    result = mcp_call_tool("get_data_profile", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "top_n": 5
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from GetDataProfileResponse
    assert "profiles" in result, "Missing 'profiles'"
    assert "columns_profiled" in result, "Missing 'columns_profiled'"
    assert "excel_output" in result, "Missing 'excel_output'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify profiles
    profiles = result["profiles"]
    assert isinstance(profiles, dict), "profiles should be dict"
    assert len(profiles) > 0, "Should have at least one column profile"
    assert result["columns_profiled"] == len(profiles), "columns_profiled count mismatch"
    
    # Verify each profile structure (from ColumnProfile model)
    for col_name, profile in profiles.items():
        print(f"  Checking profile for '{col_name}'...")
        
        # Required fields
        assert "column_name" in profile, f"Profile for '{col_name}' missing 'column_name'"
        assert "data_type" in profile, f"Profile for '{col_name}' missing 'data_type'"
        assert "total_count" in profile, f"Profile for '{col_name}' missing 'total_count'"
        assert "null_count" in profile, f"Profile for '{col_name}' missing 'null_count'"
        assert "null_percentage" in profile, f"Profile for '{col_name}' missing 'null_percentage'"
        assert "unique_count" in profile, f"Profile for '{col_name}' missing 'unique_count'"
        assert "top_values" in profile, f"Profile for '{col_name}' missing 'top_values'"
        
        # Verify values
        assert profile["column_name"] == col_name, "column_name mismatch"
        assert profile["data_type"] in ["integer", "float", "string", "datetime", "boolean"], f"Invalid data_type: {profile['data_type']}"
        assert profile["total_count"] >= 0, "total_count should be non-negative"
        assert profile["null_count"] >= 0, "null_count should be non-negative"
        assert profile["null_count"] <= profile["total_count"], "null_count can't exceed total_count"
        assert 0 <= profile["null_percentage"] <= 100, f"null_percentage should be 0-100, got {profile['null_percentage']}"
        assert profile["unique_count"] >= 0, "unique_count should be non-negative"
        
        # Verify top_values structure
        assert isinstance(profile["top_values"], list), "top_values should be list"
        assert len(profile["top_values"]) <= 5, "Should have at most 5 top values (top_n=5)"
        for top_value in profile["top_values"]:
            assert "value" in top_value, "top_value missing 'value'"
            assert "count" in top_value, "top_value missing 'count'"
            assert "percentage" in top_value, "top_value missing 'percentage'"
            assert top_value["count"] > 0, "top_value count should be positive"
            assert 0 < top_value["percentage"] <= 100, "top_value percentage should be 0-100"
        
        # Verify stats for numeric columns
        if profile["data_type"] in ["integer", "float"]:
            if "stats" in profile and profile["stats"] is not None:
                stats = profile["stats"]
                assert "count" in stats, "stats missing 'count'"
                assert "min" in stats, "stats missing 'min'"
                assert "max" in stats, "stats missing 'max'"
                # mean, median, std, q25, q75 are optional
    
    # Verify excel_output
    excel_output = result["excel_output"]
    assert "tsv" in excel_output, "excel_output missing 'tsv'"
    if excel_output["tsv"]:
        assert isinstance(excel_output["tsv"], str), "tsv should be string"
        assert len(excel_output["tsv"]) > 0, "tsv should not be empty"
    
    print(f"  ✅ Profiled {result['columns_profiled']} columns successfully")


def test_get_data_profile_specific_columns(mcp_call_tool, simple_fixture):
    """Smoke: get_data_profile works with specific column selection."""
    print(f"\n📊 Testing get_data_profile with specific columns...")
    
    # Profile only first 2 columns
    columns_to_profile = simple_fixture.columns[:2]
    print(f"  Profiling columns: {columns_to_profile}")
    
    result = mcp_call_tool("get_data_profile", {
        "file_path": str(simple_fixture.path_str),
        "sheet_name": simple_fixture.sheet_name,
        "columns": columns_to_profile,
        "top_n": 3
    })
    
    # Verify only requested columns were profiled
    profiles = result["profiles"]
    assert len(profiles) == 2, f"Expected 2 profiles, got {len(profiles)}"
    assert result["columns_profiled"] == 2, f"Expected columns_profiled=2, got {result['columns_profiled']}"
    
    profiled_columns = set(profiles.keys())
    expected_columns = set(columns_to_profile)
    assert profiled_columns == expected_columns, f"Expected {expected_columns}, got {profiled_columns}"
    
    # Verify top_n was respected
    for col_name, profile in profiles.items():
        assert len(profile["top_values"]) <= 3, f"Expected at most 3 top values, got {len(profile['top_values'])}"
    
    print(f"  ✅ Profiled only requested columns: {list(profiles.keys())}")


# ============================================================================
# FIND_COLUMN TESTS
# ============================================================================

def test_find_column_single_sheet(mcp_call_tool, simple_fixture):
    """Smoke: find_column locates column in single sheet with complete info."""
    print(f"\n🔍 Testing find_column in single sheet...")
    
    # Search for first column
    column_to_find = simple_fixture.columns[0]
    print(f"  Searching for column: '{column_to_find}'")
    
    result = mcp_call_tool("find_column", {
        "file_path": str(simple_fixture.path_str),
        "column_name": column_to_find,
        "search_all_sheets": False
    })
    
    print(f"  Result keys: {list(result.keys())}")
    
    # Verify ALL required fields from FindColumnResponse
    assert "found_in" in result, "Missing 'found_in'"
    assert "total_matches" in result, "Missing 'total_matches'"
    assert "metadata" in result, "Missing 'metadata'"
    assert "performance" in result, "Missing 'performance'"
    
    # Verify column was found
    assert result["total_matches"] >= 1, f"Column '{column_to_find}' should be found"
    assert len(result["found_in"]) >= 1, "found_in should have at least one match"
    assert len(result["found_in"]) == result["total_matches"], "found_in count should match total_matches"
    
    # Verify match structure
    for match in result["found_in"]:
        assert "sheet" in match, "Match missing 'sheet'"
        assert "column_index" in match, "Match missing 'column_index'"
        assert "row_count" in match, "Match missing 'row_count'"
        assert isinstance(match["sheet"], str), "sheet should be string"
        assert isinstance(match["column_index"], int), "column_index should be int"
        assert match["column_index"] >= 0, "column_index should be non-negative"
        assert isinstance(match["row_count"], int), "row_count should be int"
        assert match["row_count"] > 0, "row_count should be positive"
    
    first_match = result["found_in"][0]
    print(f"  ✅ Found '{column_to_find}' in sheet '{first_match['sheet']}' at index {first_match['column_index']}")


def test_find_column_multi_sheet(mcp_call_tool, multi_sheet_fixture):
    """Smoke: find_column searches across multiple sheets correctly."""
    print(f"\n🔍 Testing find_column across multiple sheets...")
    
    # Search for a column that might exist in multiple sheets
    result = mcp_call_tool("find_column", {
        "file_path": str(multi_sheet_fixture.path_str),
        "column_name": "Name",
        "search_all_sheets": True
    })
    
    # Verify response structure
    assert "found_in" in result
    assert "total_matches" in result
    assert isinstance(result["found_in"], list), "found_in should be list"
    assert result["total_matches"] == len(result["found_in"]), "total_matches should match found_in length"
    
    # If found, verify all matches
    if result["total_matches"] > 0:
        print(f"  ✅ Found 'Name' in {result['total_matches']} sheet(s)")
        for match in result["found_in"]:
            print(f"    - Sheet: {match['sheet']}, Index: {match['column_index']}, Rows: {match['row_count']}")
            assert match["sheet"] in multi_sheet_fixture.expected["sheet_names"], f"Unexpected sheet: {match['sheet']}"
    else:
        print(f"  ℹ️  'Name' not found in any sheet (this is OK for smoke test)")


def test_find_column_not_found(mcp_call_tool, simple_fixture):
    """Smoke: find_column handles non-existent column gracefully."""
    print(f"\n🔍 Testing find_column with non-existent column...")
    
    result = mcp_call_tool("find_column", {
        "file_path": str(simple_fixture.path_str),
        "column_name": "NonExistentColumn12345XYZ",
        "search_all_sheets": True
    })
    
    # Should return empty results, not error
    assert result["total_matches"] == 0, "Non-existent column should have 0 matches"
    assert len(result["found_in"]) == 0, "found_in should be empty list"
    assert isinstance(result["found_in"], list), "found_in should still be list (empty)"
    
    # Verify metadata and performance are still present
    assert "metadata" in result, "metadata should be present even for no matches"
    assert "performance" in result, "performance should be present even for no matches"
    
    print(f"  ✅ Handled non-existent column gracefully (0 matches, no error)")


def test_find_column_case_insensitive(mcp_call_tool, simple_fixture):
    """Smoke: find_column search is case-insensitive."""
    print(f"\n🔍 Testing find_column case-insensitivity...")
    
    # Get first column name and search with different case
    original_column = simple_fixture.columns[0]
    # Try uppercase version
    search_column = original_column.upper() if original_column.islower() else original_column.lower()
    
    print(f"  Original column: '{original_column}'")
    print(f"  Searching for: '{search_column}'")
    
    result = mcp_call_tool("find_column", {
        "file_path": str(simple_fixture.path_str),
        "column_name": search_column,
        "search_all_sheets": False
    })
    
    # Should find the column despite case difference
    if original_column.lower() != original_column.upper():  # Only test if case matters
        assert result["total_matches"] >= 1, f"Case-insensitive search should find '{original_column}' when searching for '{search_column}'"
        print(f"  ✅ Found column with case-insensitive search")
    else:
        print(f"  ℹ️  Column name is case-insensitive by nature (e.g., numbers)")
