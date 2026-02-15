# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Smoke tests - verify testing infrastructure works.

These tests check that:
- Fixtures can be loaded
- Registry works correctly
- Core components can be imported and instantiated
- Basic operations succeed

Run: pytest tests/test_smoke.py -v
"""

import pytest
from pathlib import Path


def test_fixtures_exist():
    """Verify all fixture files exist on disk."""
    from tests.fixtures.registry import FIXTURES
    
    print("\n🔍 Checking fixture files...")
    missing = []
    
    for name, fixture in FIXTURES.items():
        if not fixture.path.exists():
            missing.append(f"{name}: {fixture.path}")
            print(f"  ❌ Missing: {name}")
        else:
            print(f"  ✅ Found: {name} ({fixture.path.stat().st_size} bytes)")
    
    assert len(missing) == 0, f"Missing fixtures: {missing}"
    print(f"\n✅ All {len(FIXTURES)} fixtures exist")


def test_registry_metadata():
    """Verify registry has correct structure."""
    from tests.fixtures.registry import FIXTURES, get_fixture, get_fixtures_by_category
    
    print("\n🔍 Checking registry metadata...")
    
    # Check total count
    assert len(FIXTURES) == 16, f"Expected 16 fixtures, got {len(FIXTURES)}"
    print(f"  ✅ Registry has 16 fixtures")
    
    # Check categories
    basic = get_fixtures_by_category("basic")
    messy = get_fixtures_by_category("messy")
    edge = get_fixtures_by_category("edge_cases")
    legacy = get_fixtures_by_category("legacy")
    
    assert len(basic) == 3, f"Expected 3 basic fixtures, got {len(basic)}"
    assert len(messy) == 4, f"Expected 4 messy fixtures, got {len(messy)}"
    assert len(edge) == 8, f"Expected 8 edge case fixtures, got {len(edge)}"
    assert len(legacy) == 1, f"Expected 1 legacy fixture, got {len(legacy)}"
    
    print(f"  ✅ Categories: basic={len(basic)}, messy={len(messy)}, edge={len(edge)}, legacy={len(legacy)}")
    
    # Check get_fixture works
    simple = get_fixture("simple")
    assert simple.name == "simple"
    assert simple.columns == ["Имя", "Возраст", "Город"]
    print(f"  ✅ get_fixture() works")


def test_core_components_import():
    """Verify core components can be imported."""
    print("\n🔍 Importing core components...")
    
    from mcp_excel.core.file_loader import FileLoader
    from mcp_excel.core.header_detector import HeaderDetector
    from mcp_excel.core.datetime_detector import DateTimeDetector
    from mcp_excel.core.datetime_converter import DateTimeConverter
    from mcp_excel.operations.filtering import FilterEngine
    from mcp_excel.excel.formula_generator import FormulaGenerator
    from mcp_excel.excel.tsv_formatter import TSVFormatter
    
    print("  ✅ All core components imported successfully")


def test_file_loader_basic(simple_fixture, file_loader):
    """Verify FileLoader can load a simple fixture."""
    print(f"\n🔍 Testing FileLoader with {simple_fixture.name}...")
    print(f"  Path: {simple_fixture.path_str}")
    
    # Load file with header_row
    df = file_loader.load(
        simple_fixture.path_str,
        simple_fixture.sheet_name,
        header_row=simple_fixture.header_row
    )
    
    # Verify structure
    assert len(df.columns) == len(simple_fixture.columns), \
        f"Expected {len(simple_fixture.columns)} columns, got {len(df.columns)}"
    assert len(df) == simple_fixture.row_count, \
        f"Expected {simple_fixture.row_count} rows, got {len(df)}"
    
    print(f"  ✅ Loaded: {len(df)} rows, {len(df.columns)} columns")
    print(f"  Columns: {list(df.columns)}")


def test_file_loader_with_dates(with_dates_fixture, file_loader):
    """Verify FileLoader can load fixture with dates."""
    print(f"\n🔍 Testing FileLoader with {with_dates_fixture.name}...")
    
    # Load file with datetime conversion and header_row
    df = file_loader.load(
        with_dates_fixture.path_str,
        with_dates_fixture.sheet_name,
        header_row=with_dates_fixture.header_row,
        convert_dates=True
    )
    
    # Verify datetime columns exist
    assert len(df) == with_dates_fixture.row_count
    print(f"  ✅ Loaded: {len(df)} rows with datetime columns")
    
    # Check column types
    for col in df.columns:
        dtype = str(df[col].dtype)
        print(f"    {col}: {dtype}")


def test_legacy_format(simple_legacy_fixture, file_loader):
    """Verify FileLoader can load legacy .xls format."""
    print(f"\n🔍 Testing legacy .xls format...")
    
    df = file_loader.load(
        simple_legacy_fixture.path_str,
        simple_legacy_fixture.sheet_name,
        header_row=simple_legacy_fixture.header_row
    )
    
    assert len(df) == simple_legacy_fixture.row_count
    assert len(df.columns) == len(simple_legacy_fixture.columns)
    
    print(f"  ✅ Legacy .xls loaded: {len(df)} rows, {len(df.columns)} columns")


def test_conftest_fixtures_available(
    file_loader,
    header_detector,
    datetime_detector,
    filter_engine,
    tsv_formatter
):
    """Verify all conftest.py fixtures are available."""
    print("\n🔍 Checking conftest.py fixtures...")
    
    assert file_loader is not None
    assert header_detector is not None
    assert datetime_detector is not None
    assert filter_engine is not None
    assert tsv_formatter is not None
    
    print("  ✅ All conftest fixtures available")


def test_parametrize_works(basic_fixture_meta, file_loader):
    """Verify parametrization works (runs 3 times for basic fixtures)."""
    print(f"\n🔍 Testing parametrization with {basic_fixture_meta.name}...")
    
    # This test runs once per basic fixture (3 times total)
    df = file_loader.load(
        basic_fixture_meta.path_str,
        basic_fixture_meta.sheet_name,
        header_row=basic_fixture_meta.header_row
    )
    
    assert len(df) > 0
    assert len(df.columns) > 0
    
    print(f"  ✅ {basic_fixture_meta.name}: {len(df)} rows, {len(df.columns)} columns")


def test_cache_works(simple_fixture, file_loader):
    """Verify FileLoader caching works."""
    print("\n🔍 Testing FileLoader cache...")
    
    # First load
    df1 = file_loader.load(
        simple_fixture.path_str,
        simple_fixture.sheet_name,
        header_row=simple_fixture.header_row
    )
    
    # Second load (should be from cache)
    df2 = file_loader.load(
        simple_fixture.path_str,
        simple_fixture.sheet_name,
        header_row=simple_fixture.header_row
    )
    
    # Should be same object (from cache)
    assert df1 is df2, "Second load should return cached DataFrame"
    
    # Check cache stats
    stats = file_loader.get_cache_stats()
    print(f"  Cache stats: {stats}")
    assert stats["size"] > 0, "Cache should have entries"
    
    print("  ✅ Cache working correctly")


def test_helper_functions(assert_dataframe_equals, assert_excel_formula):
    """Verify helper functions from conftest work."""
    print("\n🔍 Testing helper functions...")
    
    import pandas as pd
    
    # Test assert_dataframe_equals
    df1 = pd.DataFrame({"A": [1, 2, 3]})
    df2 = pd.DataFrame({"A": [1.0, 2.0, 3.0]})  # Different dtype
    assert_dataframe_equals(df1, df2)  # Should pass (check_dtype=False by default)
    print("  ✅ assert_dataframe_equals works")
    
    # Test assert_excel_formula
    formula = "=SUM(A1:A10)"
    assert_excel_formula(formula, starts_with="=", contains="SUM")
    print("  ✅ assert_excel_formula works")


if __name__ == "__main__":
    # Allow running directly for quick checks
    pytest.main([__file__, "-v", "-s"])
