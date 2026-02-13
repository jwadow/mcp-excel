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
    GetColumnNamesRequest,
    GetSheetInfoRequest,
    InspectFileRequest,
)
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

        print_section("✅ All Tests Completed Successfully")

    except Exception as e:
        print(f"\n❌ Error during testing: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
