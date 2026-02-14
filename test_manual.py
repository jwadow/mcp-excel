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
    CompareSheetsRequest,
    CorrelateRequest,
    DetectOutliersRequest,
    FilterAndCountRequest,
    FilterAndGetRowsRequest,
    FilterCondition,
    FindColumnRequest,
    GetColumnNamesRequest,
    GetColumnStatsRequest,
    GetSheetInfoRequest,
    GetUniqueValuesRequest,
    GetValueCountsRequest,
    GroupByRequest,
    InspectFileRequest,
    SearchAcrossSheetsRequest,
)
from mcp_excel.operations.data_operations import DataOperations
from mcp_excel.operations.inspection import InspectionOperations
from mcp_excel.operations.statistics import StatisticsOperations


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


def test_formula_generation(file_path: str) -> None:
    """Test Excel formula generation for all operators."""
    print_section("Testing Excel Formula Generation")

    loader = FileLoader()
    ops = DataOperations(loader)

    # Get first sheet name
    sheet_names = loader.get_sheet_names(file_path)
    if not sheet_names:
        print("❌ No sheets found in file")
        return

    sheet_name = sheet_names[0]
    print(f"📊 Using sheet: {sheet_name}")

    # Get sheet info
    inspection_ops = InspectionOperations(loader)
    sheet_info_request = GetSheetInfoRequest(file_path=file_path, sheet_name=sheet_name)
    sheet_info = inspection_ops.get_sheet_info(sheet_info_request)
    
    if not sheet_info.column_names or len(sheet_info.column_names) < 2:
        print("❌ Need at least 2 columns for testing")
        return

    first_column = sheet_info.column_names[0]
    
    # Get a sample value for testing
    unique_request = GetUniqueValuesRequest(
        file_path=file_path,
        sheet_name=sheet_name,
        column=first_column,
        limit=1
    )
    unique_response = ops.get_unique_values(unique_request)
    
    if not unique_response.values:
        print("❌ No values found for testing")
        return
    
    test_value = unique_response.values[0]
    print(f"\n📋 Test column: '{first_column}'")
    print(f"📋 Test value: '{test_value}'")

    # Test all comparison operators
    print(f"\n\n🔢 Testing Comparison Operators:")
    
    operators_to_test = [
        ("==", test_value, "Equal to"),
        ("!=", test_value, "Not equal to"),
        (">", test_value, "Greater than"),
        ("<", test_value, "Less than"),
        (">=", test_value, "Greater or equal"),
        ("<=", test_value, "Less or equal"),
    ]
    
    for operator, value, description in operators_to_test:
        try:
            request = FilterAndCountRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                filters=[
                    FilterCondition(column=first_column, operator=operator, value=value)
                ],
                logic="AND"
            )
            response = ops.filter_and_count(request)
            
            print(f"\n  {description} ({operator}):")
            print(f"    Count: {response.count}")
            print(f"    Formula: {response.excel_output.formula}")
        except Exception as e:
            print(f"\n  {description} ({operator}): ❌ Error: {e}")

    # Test text operators (if column is string)
    if isinstance(test_value, str) and len(test_value) > 2:
        print(f"\n\n📝 Testing Text Operators:")
        
        text_operators = [
            ("contains", test_value[:3], f"Contains '{test_value[:3]}'"),
            ("startswith", test_value[:2], f"Starts with '{test_value[:2]}'"),
            ("endswith", test_value[-2:], f"Ends with '{test_value[-2:]}'"),
        ]
        
        for operator, value, description in text_operators:
            try:
                request = FilterAndCountRequest(
                    file_path=file_path,
                    sheet_name=sheet_name,
                    filters=[
                        FilterCondition(column=first_column, operator=operator, value=value)
                    ],
                    logic="AND"
                )
                response = ops.filter_and_count(request)
                
                print(f"\n  {description}:")
                print(f"    Count: {response.count}")
                print(f"    Formula: {response.excel_output.formula}")
            except Exception as e:
                print(f"\n  {description}: ❌ Error: {e}")

    # Test 'in' operator
    print(f"\n\n📦 Testing Set Operators:")
    
    # Get multiple values for 'in' test
    unique_request_multi = GetUniqueValuesRequest(
        file_path=file_path,
        sheet_name=sheet_name,
        column=first_column,
        limit=3
    )
    unique_response_multi = ops.get_unique_values(unique_request_multi)
    
    if len(unique_response_multi.values) >= 2:
        test_values = unique_response_multi.values[:2]
        
        try:
            request = FilterAndCountRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                filters=[
                    FilterCondition(column=first_column, operator="in", values=test_values)
                ],
                logic="AND"
            )
            response = ops.filter_and_count(request)
            
            print(f"\n  In {test_values}:")
            print(f"    Count: {response.count}")
            print(f"    Formula: {response.excel_output.formula}")
        except Exception as e:
            print(f"\n  In operator: ❌ Error: {e}")

    # Test null operators
    print(f"\n\n🔍 Testing Null Operators:")
    
    null_operators = [
        ("is_null", None, "Is null"),
        ("is_not_null", None, "Is not null"),
    ]
    
    for operator, value, description in null_operators:
        try:
            request = FilterAndCountRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                filters=[
                    FilterCondition(column=first_column, operator=operator, value=value)
                ],
                logic="AND"
            )
            response = ops.filter_and_count(request)
            
            print(f"\n  {description}:")
            print(f"    Count: {response.count}")
            print(f"    Formula: {response.excel_output.formula}")
        except Exception as e:
            print(f"\n  {description}: ❌ Error: {e}")

    print(f"\n\n💡 Tip: Copy any formula above and paste it into Excel to verify it works!")


def test_statistics_operations(file_path: str) -> None:
    """Test statistical operations."""
    print_section("Testing Statistics Operations")

    loader = FileLoader()
    ops = StatisticsOperations(loader)

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

    # Find numeric columns for statistics
    numeric_columns = []
    for col_name, col_type in sheet_info.column_types.items():
        if col_type in ["integer", "float"]:
            numeric_columns.append(col_name)
    
    if not numeric_columns:
        # Try first column as fallback
        numeric_columns = [sheet_info.column_names[0]]
        print(f"  ℹ️ No numeric columns found, trying '{numeric_columns[0]}' (may contain numeric data as text)")
    
    print(f"\n📋 Numeric columns found: {', '.join(numeric_columns[:3])}...")
    test_column = numeric_columns[0]

    # Test 1: Get column statistics
    print(f"\n\n📊 Test 1: Getting statistics for '{test_column}'...")
    try:
        request = GetColumnStatsRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            column=test_column,
            filters=[]
        )
        response = ops.get_column_stats(request)
        
        print(f"  ✅ Statistics:")
        print(f"    Count: {response.stats.count}")
        print(f"    Mean: {response.stats.mean:.2f}")
        print(f"    Median: {response.stats.median:.2f}")
        std_str = f"{response.stats.std:.2f}" if response.stats.std is not None else "N/A"
        print(f"    Std Dev: {std_str}")
        print(f"    Min: {response.stats.min}")
        print(f"    Max: {response.stats.max}")
        print(f"    25th Percentile: {response.stats.q25:.2f}")
        print(f"    75th Percentile: {response.stats.q75:.2f}")
        print(f"    Null Count: {response.stats.null_count}")
        print(f"\n  📋 TSV Output (first 150 chars):")
        print(f"    {response.excel_output.tsv[:150]}...")
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 2: Correlation analysis (if we have 2+ numeric columns)
    if len(numeric_columns) >= 2:
        print(f"\n\n🔗 Test 2: Correlation analysis between columns...")
        try:
            # Use first 2-3 numeric columns
            corr_columns = numeric_columns[:min(3, len(numeric_columns))]
            print(f"  Analyzing: {', '.join(corr_columns)}")
            
            request = CorrelateRequest(
                file_path=file_path,
                sheet_name=sheet_name,
                columns=corr_columns,
                method="pearson",
                filters=[]
            )
            response = ops.correlate(request)
            
            print(f"  ✅ Correlation matrix ({response.method}):")
            for col1 in corr_columns:
                print(f"    {col1}:")
                for col2 in corr_columns:
                    corr_value = response.correlation_matrix[col1][col2]
                    print(f"      vs {col2}: {corr_value:.4f}")
            
            print(f"\n  📋 TSV Output (first 200 chars):")
            print(f"    {response.excel_output.tsv[:200]}...")
            print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
        except Exception as e:
            print(f"  ❌ Error: {e}")
    else:
        print(f"\n\n🔗 Test 2: Correlation analysis...")
        print(f"  ⚠️ Skipped: Need at least 2 numeric columns (found {len(numeric_columns)})")

    # Test 3: Outlier detection (IQR method)
    print(f"\n\n🎯 Test 3: Detecting outliers in '{test_column}' (IQR method)...")
    try:
        request = DetectOutliersRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            column=test_column,
            method="iqr",
            threshold=1.5
        )
        response = ops.detect_outliers(request)
        
        print(f"  ✅ Outliers detected: {response.outlier_count}")
        print(f"  Method: {response.method}")
        print(f"  Threshold: {response.threshold}")
        
        if response.outliers:
            print(f"\n  Sample outliers (first 3):")
            for idx, outlier in enumerate(response.outliers[:3], 1):
                # Show first 3 fields of each outlier
                outlier_preview = dict(list(outlier.items())[:3])
                print(f"    Outlier {idx}: {outlier_preview}")
            
            print(f"\n  📋 TSV Output (first 200 chars):")
            print(f"    {response.excel_output.tsv[:200]}...")
        else:
            print(f"  ℹ️ No outliers found with current threshold")
        
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 4: Outlier detection (Z-score method)
    print(f"\n\n🎯 Test 4: Detecting outliers in '{test_column}' (Z-score method)...")
    try:
        request = DetectOutliersRequest(
            file_path=file_path,
            sheet_name=sheet_name,
            column=test_column,
            method="zscore",
            threshold=3.0
        )
        response = ops.detect_outliers(request)
        
        print(f"  ✅ Outliers detected: {response.outlier_count}")
        print(f"  Method: {response.method}")
        print(f"  Threshold: {response.threshold}")
        
        if response.outliers:
            print(f"  Sample outliers (first 2):")
            for idx, outlier in enumerate(response.outliers[:2], 1):
                outlier_preview = dict(list(outlier.items())[:3])
                print(f"    Outlier {idx}: {outlier_preview}")
        else:
            print(f"  ℹ️ No outliers found with current threshold")
        
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")


def test_multisheet_operations(file_path: str) -> None:
    """Test multi-sheet operations."""
    print_section("Testing Multi-Sheet Operations (Block 4)")

    loader = FileLoader()
    ops = InspectionOperations(loader)

    # Get sheet names
    sheet_names = loader.get_sheet_names(file_path)
    if len(sheet_names) < 1:
        print("❌ No sheets found in file")
        return

    print(f"📊 File has {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")

    # Get first sheet info for column names
    inspection_ops = InspectionOperations(loader)
    sheet_info_request = GetSheetInfoRequest(file_path=file_path, sheet_name=sheet_names[0])
    sheet_info = inspection_ops.get_sheet_info(sheet_info_request)
    
    if not sheet_info.column_names:
        print("❌ No columns found in first sheet")
        return

    first_column = sheet_info.column_names[0]
    print(f"\n📋 Using column '{first_column}' for tests")

    # Test 1: find_column
    print(f"\n\n🔍 Test 1: Finding column '{first_column}' across all sheets...")
    try:
        request = FindColumnRequest(
            file_path=file_path,
            column_name=first_column,
            search_all_sheets=True
        )
        response = ops.find_column(request)
        
        print(f"  ✅ Found in {response.total_matches} location(s)")
        for match in response.found_in:
            print(f"    Sheet: {match['sheet']}, Column: {match['column_name']}, Index: {match['column_index']}, Rows: {match['row_count']}")
        print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 2: search_across_sheets
    print(f"\n\n🔎 Test 2: Searching for a value across all sheets...")
    try:
        # Get a sample value from first column
        data_ops = DataOperations(loader)
        unique_request = GetUniqueValuesRequest(
            file_path=file_path,
            sheet_name=sheet_names[0],
            column=first_column,
            limit=1
        )
        unique_response = data_ops.get_unique_values(unique_request)
        
        if unique_response.values:
            search_value = unique_response.values[0]
            print(f"  Searching for value: '{search_value}' in column '{first_column}'")
            
            request = SearchAcrossSheetsRequest(
                file_path=file_path,
                column_name=first_column,
                value=search_value
            )
            response = ops.search_across_sheets(request)
            
            print(f"  ✅ Total matches: {response.total_matches}")
            print(f"  Found in {len(response.matches)} sheet(s):")
            for match in response.matches:
                print(f"    Sheet: {match['sheet']}, Matches: {match['match_count']}/{match['total_rows']} rows")
            print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
        else:
            print(f"  ⚠️ No values found to search for")
    except Exception as e:
        print(f"  ❌ Error: {e}")

    # Test 3: compare_sheets (only if we have 2+ sheets)
    if len(sheet_names) >= 2:
        print(f"\n\n⚖️ Test 3: Comparing sheets '{sheet_names[0]}' and '{sheet_names[1]}'...")
        try:
            # Get columns from both sheets
            sheet1_info = inspection_ops.get_sheet_info(
                GetSheetInfoRequest(file_path=file_path, sheet_name=sheet_names[0])
            )
            sheet2_info = inspection_ops.get_sheet_info(
                GetSheetInfoRequest(file_path=file_path, sheet_name=sheet_names[1])
            )
            
            # Find common columns
            common_columns = set(sheet1_info.column_names) & set(sheet2_info.column_names)
            
            if len(common_columns) >= 2:
                common_list = list(common_columns)
                key_column = common_list[0]
                compare_columns = common_list[1:min(3, len(common_list))]  # Compare up to 2 columns
                
                print(f"  Key column: '{key_column}'")
                print(f"  Comparing columns: {compare_columns}")
                
                request = CompareSheetsRequest(
                    file_path=file_path,
                    sheet1=sheet_names[0],
                    sheet2=sheet_names[1],
                    key_column=key_column,
                    compare_columns=compare_columns
                )
                response = ops.compare_sheets(request)
                
                print(f"  ✅ Differences found: {response.difference_count}")
                if response.differences:
                    print(f"  Sample differences (first 3):")
                    for idx, diff in enumerate(response.differences[:3], 1):
                        status = diff.get('status', 'unknown')
                        key_val = diff.get(key_column, 'N/A')
                        print(f"    Diff {idx}: {key_val} - {status}")
                    
                    print(f"\n  📋 TSV Output (first 200 chars):")
                    print(f"    {response.excel_output.tsv[:200]}...")
                else:
                    print(f"  ℹ️ No differences found - sheets are identical")
                
                print(f"  ⚡ Execution time: {response.performance.execution_time_ms}ms")
            else:
                print(f"  ⚠️ Not enough common columns between sheets (found {len(common_columns)})")
        except Exception as e:
            print(f"  ❌ Error: {e}")
    else:
        print(f"\n\n⚖️ Test 3: Comparing sheets...")
        print(f"  ⚠️ Skipped: Need at least 2 sheets (found {len(sheet_names)})")


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
        test_statistics_operations(file_path)
        test_multisheet_operations(file_path)

        # This test should be at the very end for ease of copying and pasting
        test_formula_generation(file_path)

        print_section("✅ All Tests Completed Successfully")

    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
