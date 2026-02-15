# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Fixture Registry - metadata about all test Excel files.

This registry provides structured information about each fixture:
- Location and format
- Structure (columns, rows, header location)
- What it tests (edge cases, features)
- Expected values for assertions

Usage in tests:
    from tests.fixtures.registry import FIXTURES, get_fixture
    
    fixture = get_fixture("simple")
    assert fixture.columns == ["Имя", "Возраст", "Город"]
"""

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Dict, Any


@dataclass
class FixtureMetadata:
    """Metadata about a single test fixture."""
    
    # Identity
    name: str
    category: str  # basic, messy, edge_cases, legacy
    file_name: str
    
    # Structure
    format: str  # xlsx, xls
    sheet_name: str
    header_row: Optional[int]  # 0-based index, None if needs detection
    columns: List[str]
    row_count: int  # Data rows (excluding headers)
    
    # What it tests
    description: str
    tests: List[str]  # List of features/edge cases this fixture tests
    
    # Expected values (for assertions)
    expected: Dict[str, Any]
    
    @property
    def path(self) -> Path:
        """Get full path to fixture file."""
        base = Path(__file__).parent
        return base / self.category / self.file_name
    
    @property
    def path_str(self) -> str:
        """Get path as string (for FileLoader)."""
        return str(self.path)


# ============================================================================
# BASIC FIXTURES
# ============================================================================

SIMPLE = FixtureMetadata(
    name="simple",
    category="basic",
    file_name="simple.xlsx",
    format="xlsx",
    sheet_name="Data",
    header_row=0,
    columns=["Имя", "Возраст", "Город"],
    row_count=10,
    description="Simple table with Cyrillic data, clean structure",
    tests=[
        "basic_loading",
        "cyrillic_encoding",
        "column_detection",
        "data_types_string_int",
    ],
    expected={
        "first_row": {"Имя": "Алексей", "Возраст": 25, "Город": "Москва"},
        "unique_cities": 10,
        "age_range": (25, 33),
    }
)

WITH_DATES = FixtureMetadata(
    name="with_dates",
    category="basic",
    file_name="with_dates.xlsx",
    format="xlsx",
    sheet_name="Sales",
    header_row=0,
    columns=["Номер заказа", "Клиент", "Сумма", "Дата заказа", "Дата доставки"],
    row_count=15,
    description="Table with datetime columns (date + time)",
    tests=[
        "datetime_detection",
        "datetime_conversion",
        "datetime_filtering",
        "date_formats",
    ],
    expected={
        "datetime_columns": ["Дата заказа", "Дата доставки"],
        "date_range_start": "2024-01-03",  # First order date
        "unique_clients": 5,
    }
)

NUMERIC_TYPES = FixtureMetadata(
    name="numeric_types",
    category="basic",
    file_name="numeric_types.xlsx",
    format="xlsx",
    sheet_name="Numbers",
    header_row=0,
    columns=["Код товара", "Количество", "Цена", "Скидка", "Итого"],
    row_count=20,
    description="Different numeric types (int, float) and large numbers",
    tests=[
        "numeric_type_detection",
        "int_vs_float",
        "large_integers",
        "float_precision",
        "numeric_aggregation",
    ],
    expected={
        "product_id_range": (50089401, 50089420),
        "quantity_sum": 2100,  # 10+20+30+...+200
        "has_floats": True,
    }
)

MULTI_SHEET = FixtureMetadata(
    name="multi_sheet",
    category="basic",
    file_name="multi_sheet.xlsx",
    format="xlsx",
    sheet_name="Products",  # Default sheet for single-sheet tests
    header_row=0,
    columns=["Товар", "Цена", "Категория"],
    row_count=5,
    description="File with 3 sheets (Products, Clients, Orders) for multi-sheet testing",
    tests=[
        "multi_sheet_operations",
        "cache_per_sheet",
        "sheet_navigation",
        "cross_sheet_operations",
    ],
    expected={
        "sheet_count": 3,
        "sheet_names": ["Products", "Clients", "Orders"],
        "products_count": 5,
        "clients_count": 4,
        "orders_count": 3,
    }
)

# ============================================================================
# MESSY FIXTURES (Real World)
# ============================================================================

MESSY_HEADERS = FixtureMetadata(
    name="messy_headers",
    category="messy",
    file_name="messy_headers.xlsx",
    format="xlsx",
    sheet_name="Report",
    header_row=3,  # Headers in row 4 (0-based index 3)
    columns=["Клиент", "Сумма", "Дата", "Статус"],
    row_count=20,
    description="Headers from row 4, junk above (company name, report title)",
    tests=[
        "header_detection",
        "junk_rows_handling",
        "auto_header_detection",
    ],
    expected={
        "junk_rows": 3,
        "unique_clients": 5,
        "unique_statuses": 3,
    }
)

MERGED_CELLS = FixtureMetadata(
    name="merged_cells",
    category="messy",
    file_name="merged_cells.xlsx",
    format="xlsx",
    sheet_name="Report",
    header_row=2,  # Row 3 (0-based)
    columns=["Регион", "Январь", "Февраль", "Март", "Апрель"],
    row_count=5,
    description="Merged cells in headers (typical enterprise report)",
    tests=[
        "merged_cells_handling",
        "complex_header_structure",
        "multi_column_headers",
    ],
    expected={
        "regions": 5,
        "quarters": 2,  # Q1 and Q2
    }
)

MULTILEVEL_HEADERS = FixtureMetadata(
    name="multilevel_headers",
    category="messy",
    file_name="multilevel_headers.xlsx",
    format="xlsx",
    sheet_name="Sales",
    header_row=2,  # Row 3 (0-based) - deepest level
    columns=["ID", "Клиент", "Q1", "Q2", "Q3", "Доход", "Расход"],
    row_count=10,
    description="3-level header hierarchy (company -> categories -> subcategories)",
    tests=[
        "multilevel_headers",
        "header_hierarchy",
        "complex_structure",
    ],
    expected={
        "header_levels": 3,
        "main_categories": 3,  # Информация, Продажи, Финансы
    }
)

ENTERPRISE_CHAOS = FixtureMetadata(
    name="enterprise_chaos",
    category="messy",
    file_name="enterprise_chaos.xlsx",
    format="xlsx",
    sheet_name="Отчёт",
    header_row=6,  # Row 7 (0-based) - after all junk and merges
    columns=["Контрагент", "Январь", "Февраль", "Март", "Сумма", "Количество"],
    row_count=5,
    description="Worst case: junk + merged + multi-level + formulas + empty rows",
    tests=[
        "worst_case_scenario",
        "all_edge_cases_combined",
        "formula_handling",
        "empty_rows",
    ],
    expected={
        "junk_rows_before_data": 7,
        "has_formulas": True,
        "has_merged_cells": True,
    }
)

# ============================================================================
# EDGE CASES
# ============================================================================

WITH_NULLS = FixtureMetadata(
    name="with_nulls",
    category="edge_cases",
    file_name="with_nulls.xlsx",
    format="xlsx",
    sheet_name="Data",
    header_row=0,
    columns=["ID", "Имя", "Email", "Телефон", "Примечания"],
    row_count=10,
    description="Table with null/empty values in various columns",
    tests=[
        "null_handling",
        "find_nulls_operation",
        "null_percentage",
    ],
    expected={
        "total_nulls": 10,  # Across all columns
        "columns_with_nulls": ["Имя", "Email", "Телефон", "Примечания"],
    }
)

WITH_DUPLICATES = FixtureMetadata(
    name="with_duplicates",
    category="edge_cases",
    file_name="with_duplicates.xlsx",
    format="xlsx",
    sheet_name="Orders",
    header_row=0,
    columns=["Клиент", "Товар", "Количество", "Дата"],
    row_count=8,
    description="Table with intentional duplicate rows",
    tests=[
        "find_duplicates_operation",
        "duplicate_detection",
    ],
    expected={
        "total_duplicates": 5,  # Including all occurrences
        "unique_rows": 5,
    }
)

WIDE_TABLE = FixtureMetadata(
    name="wide_table",
    category="edge_cases",
    file_name="wide_table.xlsx",
    format="xlsx",
    sheet_name="Wide",
    header_row=0,
    columns=[f"Колонка_{i+1}" for i in range(50)],
    row_count=10,
    description="Wide table with 50 columns",
    tests=[
        "many_columns",
        "column_limit_handling",
        "wide_table_performance",
    ],
    expected={
        "column_count": 50,
    }
)

SINGLE_COLUMN = FixtureMetadata(
    name="single_column",
    category="edge_cases",
    file_name="single_column.xlsx",
    format="xlsx",
    sheet_name="Single",
    header_row=0,
    columns=["Значение"],
    row_count=10,
    description="Minimal table with single column",
    tests=[
        "minimal_structure",
        "single_column_handling",
    ],
    expected={
        "column_count": 1,
    }
)

MIXED_LANGUAGES = FixtureMetadata(
    name="mixed_languages",
    category="edge_cases",
    file_name="mixed_languages.xlsx",
    format="xlsx",
    sheet_name="Mixed",
    header_row=0,
    columns=["Name/Имя", "Age/Возраст", "City/Город", "Comment/Комментарий"],
    row_count=8,
    description="Mixed Cyrillic, Latin, Chinese, emojis, special chars",
    tests=[
        "unicode_handling",
        "mixed_encodings",
        "emoji_support",
        "special_characters",
    ],
    expected={
        "has_emoji": True,
        "has_chinese": True,
        "has_french": True,
    }
)

SPECIAL_CHARS = FixtureMetadata(
    name="special_chars",
    category="edge_cases",
    file_name="special_chars.xlsx",
    format="xlsx",
    sheet_name="Special",
    header_row=0,
    columns=["ID", "Текст", "Спецсимволы"],
    row_count=10,
    description="Formula injection tests, special symbols",
    tests=[
        "formula_injection_protection",
        "special_char_escaping",
        "newline_tab_handling",
    ],
    expected={
        "has_formula_prefix": True,  # =1+1, +7, -100, @username
        "has_quotes": True,
        "has_newlines": True,
    }
)

WITH_FORMULAS = FixtureMetadata(
    name="with_formulas",
    category="edge_cases",
    file_name="with_formulas.xlsx",
    format="xlsx",
    sheet_name="Calculations",
    header_row=0,
    columns=["Товар", "Цена", "Количество", "Сумма", "НДС 20%", "Итого"],
    row_count=5,
    description="Excel formulas in cells (calculations)",
    tests=[
        "formula_handling",
        "formula_evaluation",
        "calculated_columns",
    ],
    expected={
        "has_formulas": True,
        "formula_columns": ["Сумма", "НДС 20%", "Итого"],
    }
)

COMPLEX_FORMATTING = FixtureMetadata(
    name="complex_formatting",
    category="edge_cases",
    file_name="complex_formatting.xlsx",
    format="xlsx",
    sheet_name="Formats",
    header_row=0,
    columns=["Описание", "Значение", "Формат"],
    row_count=10,
    description="Various number formats (%, currency, dates, scientific)",
    tests=[
        "number_format_detection",
        "percentage_handling",
        "currency_handling",
        "scientific_notation",
    ],
    expected={
        "format_types": ["General", "0.00", "0.00%", "currency", "date", "time", "scientific"],
    }
)

# ============================================================================
# LEGACY FORMAT
# ============================================================================

SIMPLE_LEGACY = FixtureMetadata(
    name="simple_legacy",
    category="legacy",
    file_name="simple_legacy.xls",
    format="xls",
    sheet_name="Data",
    header_row=0,
    columns=["Имя", "Возраст", "Город"],
    row_count=5,
    description="Legacy .xls format for xlrd engine testing",
    tests=[
        "xls_format",
        "xlrd_engine",
        "legacy_support",
    ],
    expected={
        "format": "xls",
    }
)

# ============================================================================
# PERFORMANCE FIXTURES
# ============================================================================

LARGE_10K = FixtureMetadata(
    name="large_10k",
    category="performance",
    file_name="large_10k.xlsx",
    format="xlsx",
    sheet_name="Orders",
    header_row=0,
    columns=["Order ID", "Customer", "Product", "Quantity", "Price", "Total", "Date", "Status", "Region"],
    row_count=10000,
    description="Large table with 10,000 rows for basic performance testing",
    tests=[
        "performance_filtering",
        "performance_aggregation",
        "performance_statistics",
        "cache_performance",
    ],
    expected={
        "unique_customers": 100,
        "unique_products": 50,
        "unique_statuses": 5,
        "unique_regions": 10,
        "has_datetime": True,
    }
)

LARGE_50K = FixtureMetadata(
    name="large_50k",
    category="performance",
    file_name="large_50k.xlsx",
    format="xlsx",
    sheet_name="Orders",
    header_row=0,
    columns=["Order ID", "Customer", "Product", "Quantity", "Price", "Total", "Date", "Status", "Region"],
    row_count=50000,
    description="Large table with 50,000 rows for aggregation stress testing",
    tests=[
        "performance_aggregation_stress",
        "performance_grouping",
        "performance_complex_filters",
        "multi_sheet_performance",
    ],
    expected={
        "unique_customers": 200,
        "unique_products": 100,
        "unique_statuses": 5,
        "unique_regions": 10,
        "has_datetime": True,
    }
)

LARGE_100K = FixtureMetadata(
    name="large_100k",
    category="performance",
    file_name="large_100k.xlsx",
    format="xlsx",
    sheet_name="Orders",
    header_row=0,
    columns=["Order ID", "Customer", "Product", "Quantity", "Price", "Total", "Date", "Status", "Region"],
    row_count=100000,
    description="Very large table with 100,000 rows for extreme stress testing",
    tests=[
        "performance_extreme_stress",
        "performance_statistics_large",
        "performance_filtering_large",
        "memory_usage",
    ],
    expected={
        "unique_customers": 500,
        "unique_products": 200,
        "unique_statuses": 5,
        "unique_regions": 10,
        "has_datetime": True,
    }
)

# ============================================================================
# REGISTRY
# ============================================================================

FIXTURES: Dict[str, FixtureMetadata] = {
    # Basic
    "simple": SIMPLE,
    "with_dates": WITH_DATES,
    "numeric_types": NUMERIC_TYPES,
    "multi_sheet": MULTI_SHEET,
    
    # Messy
    "messy_headers": MESSY_HEADERS,
    "merged_cells": MERGED_CELLS,
    "multilevel_headers": MULTILEVEL_HEADERS,
    "enterprise_chaos": ENTERPRISE_CHAOS,
    
    # Edge cases
    "with_nulls": WITH_NULLS,
    "with_duplicates": WITH_DUPLICATES,
    "wide_table": WIDE_TABLE,
    "single_column": SINGLE_COLUMN,
    "mixed_languages": MIXED_LANGUAGES,
    "special_chars": SPECIAL_CHARS,
    "with_formulas": WITH_FORMULAS,
    "complex_formatting": COMPLEX_FORMATTING,
    
    # Legacy
    "simple_legacy": SIMPLE_LEGACY,
    
    # Performance
    "large_10k": LARGE_10K,
    "large_50k": LARGE_50K,
    "large_100k": LARGE_100K,
}


def get_fixture(name: str) -> FixtureMetadata:
    """Get fixture metadata by name.
    
    Args:
        name: Fixture name (e.g., "simple", "messy_headers")
        
    Returns:
        FixtureMetadata object
        
    Raises:
        KeyError: If fixture not found
    """
    if name not in FIXTURES:
        available = ", ".join(FIXTURES.keys())
        raise KeyError(f"Fixture '{name}' not found. Available: {available}")
    return FIXTURES[name]


def get_fixtures_by_category(category: str) -> List[FixtureMetadata]:
    """Get all fixtures in a category.
    
    Args:
        category: Category name (basic, messy, edge_cases, legacy)
        
    Returns:
        List of FixtureMetadata objects
    """
    return [f for f in FIXTURES.values() if f.category == category]


def get_fixtures_by_test(test_name: str) -> List[FixtureMetadata]:
    """Get all fixtures that test a specific feature.
    
    Args:
        test_name: Test/feature name (e.g., "datetime_detection")
        
    Returns:
        List of FixtureMetadata objects
    """
    return [f for f in FIXTURES.values() if test_name in f.tests]
