# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Manual testing script for MCP Excel Server.

This script allows you to test the core functionality without running the MCP server.
You can directly call operations and see the results.
"""

import json
import sys
from pathlib import Path

# Add src to path for direct execution without installation
sys.path.insert(0, str(Path(__file__).parent / "src"))

from mcp_excel.core.file_loader import FileLoader
from mcp_excel.core.header_detector import HeaderDetector
from mcp_excel.models.requests import (
    AggregateRequest,
    FilterAndCountRequest,
    FilterAndGetRowsRequest,
    FilterCondition,
    GetColumnNamesRequest,
    GetSheetInfoRequest,
    GetUniqueValuesRequest,
    GetValueCountsRequest,
    GroupByRequest,
    InspectFileRequest,
)
from mcp_excel.operations.data_operations import DataOperations
from mcp_excel.operations.inspection import InspectionOperations


def print_section(title: str) -> None:
    """Print section header."""
    print("\n" + "=" * 80)
    print(f"  {title}")
    print("=" * 80 + "\n")


def test_file_loader(file_path: str) -> None:
    """Test FileLoader functionality."""
    print_section("Testing FileLoader")

    loader = FileLoader()

    # Get file info
    print("📁 File Information:")
    file_info = loader.get_file_info(file_path)
    print(json.dumps(file_info, indent=2))

    # Get sheet names
    print("\n📋 Sheet Names:")
    sheet_names = loader.get_sheet_names(file_path)
    for idx, name in enumerate(sheet_names, 1):
        print(f"  {idx}. {name}")

    # Load first sheet
    print(f"\n📊 Loading first sheet: {sheet_names[0]}")
    df = loader.load(file_path, sheet_names[0])
    print(f"  Rows: {len(df)}")
    print(f"  Columns: {len(df.columns)}")
    print(f"  Column names: {list(df.columns)}")

    # Cache stats
    print("\n💾 Cache Statistics:")
    cache_stats = loader.get_cache_stats()
    print(json.dumps(cache_stats, indent=2))


def test_header_detector(file_path: str, sheet_name: str) -> None:
    """Test HeaderDetector functionality."""
    print_section("Testing HeaderDetector")

    loader = FileLoader()
    detector = HeaderDetector()

    # Load without header
    print(f"📊 Loading sheet '{sheet_name}' without header...")
    df_raw = loader.load(file_path, sheet_name, header_row=None)

    # Detect header
    print("\n🔍 Detecting header row...")
    result = detector.detect(df_raw)

    print(f"  Detected header row: {result.header_row}")
    print(f"  Confidence: {result.confidence:.2%}")

    if result.candidates:
        print("\n  Top candidates:")
        for candidate in result.candidates[:3]:
            print(f"    Row {candidate['row']}: score={candidate['score']:.2f}")
            print(f"      Preview: {candidate['preview'][:5]}")


def test_inspection_operations(file_path: str) -> None:
    """Test InspectionOperations functionality."""
    print_section("Testing InspectionOperations")

    loader = FileLoader()
    ops = InspectionOperations(loader)

    # Inspect file
    print("🔍 Inspecting file...")
    request = InspectFileRequest(file_path=file_path)
    response = ops.inspect_file(request)

    print(f"\n📁 File: {response.format.upper()}")
    print(f"  Size: {response.size_mb} MB")
    print(f"  Sheets: {response.sheet_count}")

    print("\n📋 Sheets:")
    for sheet_info in response.sheets_info:
        if "error" in sheet_info:
            print(f"  ❌ {sheet_info['sheet_name']}: {sheet_info['error']}")
        else:
            print(f"  ✅ {sheet_info['sheet_name']}: {sheet_info['row_count']} rows, {sheet_info['column_count']} columns")

    print(f"\n⚡ Performance:")
    print(f"  Execution time: {response.performance.execution_time_ms}ms")
    print(f"  Memory used: {response.performance.memory_used_mb}MB")

    # Get sheet info
    if response.sheet_names:
        sheet_name = response.sheet_names[0]
        print(f"\n\n🔍 Getting detailed info for sheet '{sheet_name}'...")

        request = GetSheetInfoRequest(file_path=file_path, sheet_name=sheet_name)
        response = ops.get_sheet_info(request)

        print(f"\n📊 Sheet: {response.sheet_name}")
        print(f"  Rows: {response.row_count}")
        print(f"  Columns: {response.column_count}")

        print("\n  Column Types:")
        for col_name, col_type in list(response.column_types.items())[:10]:
            print(f"    {col_name}: {col_type}")

        if response.header_detection:
            print(f"\n  Header Detection:")
            print(f"    Row: {response.header_detection.header_row}")
            print(f"    Confidence: {response.header_detection.confidence:.2%}")

        print("\n  Sample Data (first 3 rows):")
        for idx, row in enumerate(response.sample_rows, 1):
            print(f"    Row {idx}: {dict(list(row.items())[:3])}...")


def test_data_operations(file_path: str) -> None:
    """Test DataOperations functionality."""
    print_section("Testing DataOperations")

    loader = FileLoader()
    ops = DataOperations(loader)

    # Get first sheet name
    sheet_names = loader.get_sheet_names(file_path)
    if not sheet_names:
        print("❌ No sheets found in file")
        return

    sheet_name = sheet_names[0]
    print(f"📊 Using sheet: {sheet_name}")

    # Get sheet info to know available columns
    inspection_ops = InspectionOperations(loader)
    sheet_info_request = GetSheetInfoRequest(file_path=file_path, sheet_name=sheet_name)
    sheet_info = inspection_ops.get_sheet_info(sheet_info_request)
    
    if not sheet_info.column_names:
        print("❌ No columns found in sheet")
        return

    print(f"\n📋 Available columns: {', '.join(sheet_info.column_names[:5])}...")
    first_column = sheet_info.column_names[0]

    # Test 1: Get unique values
    print(f"\n\n🔍 Test 1: Getting unique values from '{first_column}'...")
    try:
        request = GetUniqueValuesRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            column=first_column,
            limit=10
        )
        response = ops.get_unique_values(request)
        
        print(f"  ✅ Found {response.count} unique values")
        print(f"  Truncated: {response.truncated}")
        print(f"  Values: {response.values[:5]}...")
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 2: Get value counts
    print(f"\n\n📊 Test 2: Getting value counts from '{first_column}'...")
    try:
        request = GetValueCountsRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            column=first_column,
            top_n=5
        )
        response = ops.get_value_counts(request)
        
        print(f"  ✅ Total values: {response.total_values}")
        print(f"  Top values:")
        for value, count in list(response.value_counts.items())[:5]:
            print(f"    {value}: {count}")
        print(f"\n  📋 TSV Output (first 100 chars):")
        print(f"    {response.excel_output.tsv[:100]}...")
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 3: Filter and count
    print(f"\n\n🔢 Test 3: Counting rows with filter...")
    unique_response = None  # Initialize to avoid scope issues
    try:
        # Get a value to filter on
        unique_request = GetUniqueValuesRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            column=first_column,
            limit=1
        )
        unique_response = ops.get_unique_values(unique_request)
        
        if unique_response.values:
            filter_value = unique_response.values[0]
            print(f"  Filtering where '{first_column}' == '{filter_value}'")
            
            request = FilterAndCountRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                filters=[
                    FilterCondition(column=first_column, operator="==", value=filter_value)
                ],
                logic="AND"
            )
            response = ops.filter_and_count(request)
            
            print(f"  ✅ Matching rows: {response.count}")
            print(f"  📋 Excel formula: {response.excel_output.formula}")
            print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
        else:
            print(f"  ⚠️ No values found to filter on")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 4: Filter and get rows
    print(f"\n\n📄 Test 4: Getting filtered rows...")
    try:
        # Use same filter as above
        if unique_response and unique_response.values:
            filter_value = unique_response.values[0]
            
            request = FilterAndGetRowsRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                filters=[
                    FilterCondition(column=first_column, operator="==", value=filter_value)
                ],
                columns=sheet_info.column_names[:3],  # First 3 columns only
                limit=5,
                offset=0,
                logic="AND"
            )
            response = ops.filter_and_get_rows(request)
            
            print(f"  ✅ Returned {response.count} rows (total matches: {response.total_matches})")
            print(f"  Truncated: {response.truncated}")
            print(f"\n  Sample rows:")
            for idx, row in enumerate(response.rows[:3], 1):
                print(f"    Row {idx}: {dict(list(row.items())[:3])}")
            print(f"\n  📋 TSV Output (first 150 chars):")
            print(f"    {response.excel_output.tsv[:150]}...")
            print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
        else:
            print(f"  ⚠️ No values found to filter on")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 5: Complex filter (multiple conditions)
    print(f"\n\n🔍 Test 5: Complex filter with multiple conditions...")
    try:
        if len(sheet_info.column_names) >= 2 and unique_response and unique_response.values:
            first_col = sheet_info.column_names[0]
            second_col = sheet_info.column_names[1]
            filter_value = unique_response.values[0]
            
            # Get a value from second column
            unique_request2 = GetUniqueValuesRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                column=second_col,
                limit=1
            )
            unique_response2 = ops.get_unique_values(unique_request2)
            
            if unique_response2.values:
                filter_value2 = unique_response2.values[0]
                print(f"  Filtering: '{first_col}' == '{filter_value}' AND '{second_col}' == '{filter_value2}'")
                
                request = FilterAndCountRequest(
                    file_path=file_path,
                    sheet_name=sheet_name,
                    filters=[
                        FilterCondition(column=first_col, operator="==", value=filter_value),
                        FilterCondition(column=second_col, operator="==", value=filter_value2)
                    ],
                    logic="AND"
                )
                response = ops.filter_and_count(request)
                
                print(f"  ✅ Matching rows: {response.count}")
                print(f"  📋 Excel formula: {response.excel_output.formula}")
                print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
            else:
                print(f"  ⚠️ Not enough data for complex filter test")
        else:
            print(f"  ⚠️ Not enough columns for complex filter test")
    except Exception as e:
        print(f"  ❌ Error: {e}")


def test_aggregation_operations(file_path: str) -> None:
    """Test aggregation operations."""
    print_section("Testing Aggregation Operations")

    loader = FileLoader()
    ops = DataOperations(loader)

    # Get first sheet name
    sheet_names = loader.get_sheet_names(file_path)
    if not sheet_names:
        print("❌ No sheets found in file")
        return

    sheet_name = sheet_names[0]
    print(f"📊 Using sheet: {sheet_name}")

    # Get sheet info to know available columns
    inspection_ops = InspectionOperations(loader)
    sheet_info_request = GetSheetInfoRequest(file_path=file_path, sheet_name=sheet_name)
    sheet_info = inspection_ops.get_sheet_info(sheet_info_request)
    
    if not sheet_info.column_names:
        print("❌ No columns found in sheet")
        return

    # Find a numeric column for aggregation
    numeric_column = None
    for col_name, col_type in sheet_info.column_types.items():
        if col_type in ["integer", "float"]:
            numeric_column = col_name
            break
    
    if not numeric_column:
        # Try first column as fallback (often contains IDs/numbers)
        numeric_column = sheet_info.column_names[0]
        print(f"  ℹ️ No numeric columns found, trying '{numeric_column}' (may contain numeric data as text)")

    print(f"\n📋 Using numeric column: {numeric_column}")
    first_column = sheet_info.column_names[0]

    # Test 1: Simple aggregation (count)
    print(f"\n\n🔢 Test 1: Count aggregation on '{numeric_column}'...")
    try:
        request = AggregateRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            operation="count",
            target_column=numeric_column,
            filters=[]
        )
        response = ops.aggregate(request)
        
        print(f"  ✅ Count: {response.value}")
        print(f"  Operation: {response.operation}")
        print(f"  📋 Excel formula: {response.excel_output.formula}")
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 2: Sum aggregation
    print(f"\n\n➕ Test 2: Sum aggregation on '{numeric_column}'...")
    try:
        request = AggregateRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            operation="sum",
            target_column=numeric_column,
            filters=[]
        )
        response = ops.aggregate(request)
        
        print(f"  ✅ Sum: {response.value}")
        print(f"  📋 Excel formula: {response.excel_output.formula}")
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 3: Mean aggregation
    print(f"\n\n📊 Test 3: Mean aggregation on '{numeric_column}'...")
    try:
        request = AggregateRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            operation="mean",
            target_column=numeric_column,
            filters=[]
        )
        response = ops.aggregate(request)
        
        print(f"  ✅ Mean: {response.value:.2f}")
        print(f"  📋 Excel formula: {response.excel_output.formula}")
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 4: Aggregation with filter
    print(f"\n\n🔍 Test 4: Aggregation with filter...")
    try:
        # Get a value to filter on
        unique_request = GetUniqueValuesRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            column=first_column,
            limit=1
        )
        unique_response = ops.get_unique_values(unique_request)
        
        if unique_response.values:
            filter_value = unique_response.values[0]
            print(f"  Filtering where '{first_column}' == '{filter_value}'")
            
            request = AggregateRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                operation="count",
                target_column=numeric_column,
                filters=[
                    FilterCondition(column=first_column, operator="==", value=filter_value)
                ]
            )
            response = ops.aggregate(request)
            
            print(f"  ✅ Filtered count: {response.value}")
            print(f"  📋 Excel formula: {response.excel_output.formula}")
            print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
        else:
            print(f"  ⚠️ No values found to filter on")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 5: Group by single column
    print(f"\n\n📊 Test 5: Group by '{first_column}' with sum...")
    try:
        request = GroupByRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            group_columns=[first_column],
            agg_column=numeric_column,
            agg_operation="sum",
            filters=[]
        )
        response = ops.group_by(request)
        
        print(f"  ✅ Found {len(response.groups)} groups")
        print(f"  Sample groups (first 3):")
        for idx, group in enumerate(response.groups[:3], 1):
            print(f"    Group {idx}: {dict(list(group.items())[:3])}")
        print(f"\n  📋 TSV Output (first 150 chars):")
        print(f"    {response.excel_output.tsv[:150]}...")
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 6: Group by multiple columns
    if len(sheet_info.column_names) >= 2:
        print(f"\n\n📊 Test 6: Group by multiple columns...")
        try:
            second_column = sheet_info.column_names[1]
            request = GroupByRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                group_columns=[first_column, second_column],
                agg_column=numeric_column,
                agg_operation="count",
                filters=[]
            )
            response = ops.group_by(request)
            
            print(f"  ✅ Found {len(response.groups)} groups")
            print(f"  Group columns: {response.group_columns}")
            print(f"  Sample groups (first 2):")
            for idx, group in enumerate(response.groups[:2], 1):
                print(f"    Group {idx}: {group}")
            print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
        except Exception as e:
            print(f"  ❌ Error: {e}")


def main() -> None:
    """Main test function."""
    print("\n" + "=" * 80)
    print("  MCP Excel Server - Manual Testing")
    print("=" * 80)

    # Check if file path provided
    if len(sys.argv) < 2:
        print("\n❌ Error: No file path provided")
        print("\nUsage:")
        print("  python test_manual.py <path_to_excel_file>")
        print("\nExample:")
        print("  python test_manual.py ./data/sample.xlsx")
        print("  python test_manual.py C:/Users/User/Documents/report.xls")
        sys.exit(1)

    file_path = sys.argv[1]

    # Check if file exists
    if not Path(file_path).exists():
        print(f"\n❌ Error: File not found: {file_path}")
        sys.exit(1)

    print(f"\n📂 Testing with file: {file_path}\n")

    try:
        # Run tests
        test_file_loader(file_path)

        # Get sheet name for header detection test
        loader = FileLoader()
        sheet_names = loader.get_sheet_names(file_path)
        if sheet_names:
            test_header_detector(file_path, sheet_names[0])

        test_inspection_operations(file_path)
        test_data_operations(file_path)
        test_aggregation_operations(file_path)

        print_section("✅ All Tests Completed Successfully")

    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
