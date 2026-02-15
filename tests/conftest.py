# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Pytest configuration and shared fixtures.

This module provides reusable pytest fixtures for all tests.
All fixtures use the Fixture Registry for metadata - no hardcoded paths or values.

Usage in tests:
    def test_something(simple_fixture, file_loader):
        # simple_fixture provides metadata
        # file_loader is ready to use
        df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name)
        assert len(df.columns) == len(simple_fixture.columns)
"""

import sys
from pathlib import Path
from typing import Generator

import pytest

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from mcp_excel.core.file_loader import FileLoader
from mcp_excel.core.header_detector import HeaderDetector
from mcp_excel.core.datetime_detector import DateTimeDetector
from mcp_excel.core.datetime_converter import DateTimeConverter
from mcp_excel.operations.filtering import FilterEngine
from mcp_excel.excel.formula_generator import FormulaGenerator
from mcp_excel.excel.tsv_formatter import TSVFormatter

from tests.fixtures.registry import (
    FIXTURES,
    FixtureMetadata,
    get_fixture,
    get_fixtures_by_category,
)


# ============================================================================
# CORE COMPONENTS (Singletons for performance)
# ============================================================================

@pytest.fixture(scope="session")
def file_loader() -> FileLoader:
    """Provides FileLoader instance (session-scoped for caching)."""
    return FileLoader()


@pytest.fixture(scope="session")
def header_detector() -> HeaderDetector:
    """Provides HeaderDetector instance."""
    return HeaderDetector()


@pytest.fixture(scope="session")
def datetime_detector() -> DateTimeDetector:
    """Provides DateTimeDetector instance."""
    return DateTimeDetector()


@pytest.fixture(scope="session")
def datetime_converter() -> DateTimeConverter:
    """Provides DateTimeConverter instance."""
    return DateTimeConverter()


@pytest.fixture
def filter_engine() -> FilterEngine:
    """Provides FilterEngine instance (function-scoped)."""
    return FilterEngine()


@pytest.fixture
def tsv_formatter() -> TSVFormatter:
    """Provides TSVFormatter instance."""
    return TSVFormatter()


# ============================================================================
# FIXTURE METADATA (Individual fixtures)
# ============================================================================

@pytest.fixture
def simple_fixture() -> FixtureMetadata:
    """Simple table with Cyrillic data."""
    return get_fixture("simple")


@pytest.fixture
def with_dates_fixture() -> FixtureMetadata:
    """Table with datetime columns."""
    return get_fixture("with_dates")


@pytest.fixture
def numeric_types_fixture() -> FixtureMetadata:
    """Table with different numeric types."""
    return get_fixture("numeric_types")


@pytest.fixture
def multi_sheet_fixture() -> FixtureMetadata:
    """File with 3 sheets (Products, Clients, Orders)."""
    return get_fixture("multi_sheet")


@pytest.fixture
def messy_headers_fixture() -> FixtureMetadata:
    """Table with headers from row 4, junk above."""
    return get_fixture("messy_headers")


@pytest.fixture
def merged_cells_fixture() -> FixtureMetadata:
    """Table with merged cells in headers."""
    return get_fixture("merged_cells")


@pytest.fixture
def multilevel_headers_fixture() -> FixtureMetadata:
    """Table with 3-level header hierarchy."""
    return get_fixture("multilevel_headers")


@pytest.fixture
def enterprise_chaos_fixture() -> FixtureMetadata:
    """Worst case: junk + merged + multi-level + formulas."""
    return get_fixture("enterprise_chaos")


@pytest.fixture
def with_nulls_fixture() -> FixtureMetadata:
    """Table with null/empty values."""
    return get_fixture("with_nulls")


@pytest.fixture
def with_duplicates_fixture() -> FixtureMetadata:
    """Table with duplicate rows."""
    return get_fixture("with_duplicates")


@pytest.fixture
def wide_table_fixture() -> FixtureMetadata:
    """Wide table with 50 columns."""
    return get_fixture("wide_table")


@pytest.fixture
def single_column_fixture() -> FixtureMetadata:
    """Minimal table with single column."""
    return get_fixture("single_column")


@pytest.fixture
def mixed_languages_fixture() -> FixtureMetadata:
    """Mixed Cyrillic, Latin, Chinese, emojis."""
    return get_fixture("mixed_languages")


@pytest.fixture
def special_chars_fixture() -> FixtureMetadata:
    """Formula injection tests, special symbols."""
    return get_fixture("special_chars")


@pytest.fixture
def with_formulas_fixture() -> FixtureMetadata:
    """Excel formulas in cells."""
    return get_fixture("with_formulas")


@pytest.fixture
def complex_formatting_fixture() -> FixtureMetadata:
    """Various number formats."""
    return get_fixture("complex_formatting")


@pytest.fixture
def simple_legacy_fixture() -> FixtureMetadata:
    """Legacy .xls format."""
    return get_fixture("simple_legacy")


@pytest.fixture
def large_10k_fixture() -> FixtureMetadata:
    """Large table with 10,000 rows for performance testing."""
    return get_fixture("large_10k")


@pytest.fixture
def large_50k_fixture() -> FixtureMetadata:
    """Large table with 50,000 rows for stress testing."""
    return get_fixture("large_50k")


@pytest.fixture
def large_100k_fixture() -> FixtureMetadata:
    """Very large table with 100,000 rows for extreme stress testing."""
    return get_fixture("large_100k")


# ============================================================================
# FIXTURE COLLECTIONS (By category)
# ============================================================================

@pytest.fixture
def basic_fixtures() -> list[FixtureMetadata]:
    """All basic fixtures."""
    return get_fixtures_by_category("basic")


@pytest.fixture
def messy_fixtures() -> list[FixtureMetadata]:
    """All messy (real world) fixtures."""
    return get_fixtures_by_category("messy")


@pytest.fixture
def edge_case_fixtures() -> list[FixtureMetadata]:
    """All edge case fixtures."""
    return get_fixtures_by_category("edge_cases")


@pytest.fixture
def legacy_fixtures() -> list[FixtureMetadata]:
    """All legacy format fixtures."""
    return get_fixtures_by_category("legacy")


@pytest.fixture
def performance_fixtures() -> list[FixtureMetadata]:
    """All performance fixtures (large files)."""
    return get_fixtures_by_category("performance")


@pytest.fixture
def all_fixtures() -> list[FixtureMetadata]:
    """All available fixtures."""
    return list(FIXTURES.values())


# ============================================================================
# PARAMETRIZE HELPERS
# ============================================================================

def pytest_generate_tests(metafunc):
    """Automatically parametrize tests based on fixture names.
    
    Usage in tests:
        def test_all_basic(fixture_meta, file_loader):
            # This test will run for each basic fixture
            pass
    
    Supported parameter names:
        - fixture_meta: Parametrized with all fixtures
        - basic_fixture_meta: Parametrized with basic fixtures
        - messy_fixture_meta: Parametrized with messy fixtures
        - edge_fixture_meta: Parametrized with edge case fixtures
    """
    if "fixture_meta" in metafunc.fixturenames:
        # Parametrize with all fixtures
        metafunc.parametrize(
            "fixture_meta",
            list(FIXTURES.values()),
            ids=lambda f: f.name
        )
    
    elif "basic_fixture_meta" in metafunc.fixturenames:
        # Parametrize with basic fixtures only
        metafunc.parametrize(
            "basic_fixture_meta",
            get_fixtures_by_category("basic"),
            ids=lambda f: f.name
        )
    
    elif "messy_fixture_meta" in metafunc.fixturenames:
        # Parametrize with messy fixtures only
        metafunc.parametrize(
            "messy_fixture_meta",
            get_fixtures_by_category("messy"),
            ids=lambda f: f.name
        )
    
    elif "edge_fixture_meta" in metafunc.fixturenames:
        # Parametrize with edge case fixtures only
        metafunc.parametrize(
            "edge_fixture_meta",
            get_fixtures_by_category("edge_cases"),
            ids=lambda f: f.name
        )


# ============================================================================
# UTILITY FIXTURES
# ============================================================================

@pytest.fixture
def temp_excel_path(tmp_path) -> Generator[Path, None, None]:
    """Provides temporary path for creating test Excel files.
    
    Usage:
        def test_something(temp_excel_path):
            # Create temporary Excel file
            wb = openpyxl.Workbook()
            wb.save(temp_excel_path / "test.xlsx")
    """
    yield tmp_path


@pytest.fixture
def assert_dataframe_equals():
    """Provides helper function for comparing DataFrames in tests.
    
    Usage:
        def test_something(assert_dataframe_equals):
            assert_dataframe_equals(df1, df2, check_dtype=False)
    """
    import pandas as pd
    from pandas.testing import assert_frame_equal
    
    def _assert_equals(df1, df2, **kwargs):
        """Compare two DataFrames with sensible defaults."""
        defaults = {
            "check_dtype": False,  # Don't be strict about int vs float
            "check_names": True,
            "check_exact": False,
            "rtol": 1e-5,
        }
        defaults.update(kwargs)
        assert_frame_equal(df1, df2, **defaults)
    
    return _assert_equals


@pytest.fixture
def assert_excel_formula():
    """Provides helper for validating Excel formulas.
    
    Usage:
        def test_formula(assert_excel_formula):
            formula = "=SUM(A1:A10)"
            assert_excel_formula(formula, starts_with="=", contains="SUM")
    """
    def _assert_formula(formula: str, starts_with: str = "=", contains: str = None):
        """Validate Excel formula structure."""
        assert formula.startswith(starts_with), f"Formula should start with '{starts_with}'"
        if contains:
            assert contains in formula, f"Formula should contain '{contains}'"
        # Check for common formula injection patterns
        dangerous = ["=cmd", "=system", "|", "&"]
        for pattern in dangerous:
            assert pattern not in formula.lower(), f"Formula contains dangerous pattern: {pattern}"
    
    return _assert_formula


# ============================================================================
# PYTEST CONFIGURATION
# ============================================================================

def pytest_configure(config):
    """Register custom markers."""
    config.addinivalue_line(
        "markers", "unit: Unit tests (fast, isolated)"
    )
    config.addinivalue_line(
        "markers", "integration: Integration tests (slower, end-to-end)"
    )
    config.addinivalue_line(
        "markers", "slow: Slow tests (> 1 second)"
    )
    config.addinivalue_line(
        "markers", "legacy: Tests for legacy .xls format"
    )
    config.addinivalue_line(
        "markers", "datetime: Tests for datetime handling"
    )
    config.addinivalue_line(
        "markers", "edge_case: Tests for edge cases"
    )


def pytest_collection_modifyitems(config, items):
    """Auto-mark tests based on their location."""
    for item in items:
        # Auto-mark based on test file location
        if "unit" in str(item.fspath):
            item.add_marker(pytest.mark.unit)
        elif "integration" in str(item.fspath):
            item.add_marker(pytest.mark.integration)
        
        # Auto-mark based on fixture usage
        if any("legacy" in fixture for fixture in item.fixturenames):
            item.add_marker(pytest.mark.legacy)
        if any("date" in fixture for fixture in item.fixturenames):
            item.add_marker(pytest.mark.datetime)
