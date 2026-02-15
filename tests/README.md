# Testing Infrastructure Documentation

**For AI Agents:** This document describes the complete testing infrastructure for MCP Excel Server. Read this before writing any tests.

## Table of Contents

- [Overview](#overview)
- [Directory Structure](#directory-structure)
- [Test Fixtures](#test-fixtures)
- [Writing Tests](#writing-tests)
- [Running Tests](#running-tests)
- [Adding New Fixtures](#adding-new-fixtures)
- [Future Work](#future-work)

---

## Overview

### Testing Philosophy

**Principle:** Tests must be deterministic, fast, and use NO real data.

**Architecture:**
- **Static Fixtures:** 17 pre-generated Excel files covering all scenarios (basic, messy, edge cases, legacy)
- **Fixture Registry:** Central metadata store ([`registry.py`](fixtures/registry.py:1)) - NO hardcoded paths/values in tests
- **Pytest Fixtures:** Reusable components in [`conftest.py`](conftest.py:1)
- **Separation:** Unit tests (isolated, fast) vs Integration tests (end-to-end, slower)

### What's Tested

**Core Components:**
- [`FileLoader`](../src/mcp_excel/core/file_loader.py:1) - loading .xls/.xlsx, caching
- [`HeaderDetector`](../src/mcp_excel/core/header_detector.py:1) - auto-detection of header rows
- [`DateTimeDetector`](../src/mcp_excel/core/datetime_detector.py:1) - datetime column detection
- [`FilterEngine`](../src/mcp_excel/operations/filtering.py:1) - 12 filter operators
- [`FormulaGenerator`](../src/mcp_excel/excel/formula_generator.py:1) - Excel formula generation

**Operations:** All 23 MCP tools (inspection, data operations, statistics, validation, etc.)

---

## Directory Structure

```
tests/
├── conftest.py              # Pytest configuration & shared fixtures
├── README.md                # This file
│
├── fixtures/                # Test Excel files (committed to git)
│   ├── registry.py          # Fixture metadata (paths, columns, expected values)
│   ├── basic/               # 3 files: simple, with_dates, numeric_types
│   ├── messy/               # 4 files: enterprise scenarios (merged cells, multi-level headers)
│   ├── edge_cases/          # 9 files: nulls, duplicates, wide tables, special chars, etc.
│   └── legacy/              # 1 file: simple_legacy.xls
│
├── builders/                # Fixture generation scripts
│   └── generate_fixtures.py # Run ONCE to regenerate all fixtures
│
├── unit/                    # Unit tests (fast, isolated)
│   ├── core/                # Tests for core components
│   ├── excel/               # Tests for formula/TSV generators
│   └── operations/          # Tests for filtering engine
│
└── integration/             # Integration tests (end-to-end)
    ├── test_inspection.py
    ├── test_data_operations.py
    └── ...
```

---

## Test Fixtures

### Static Fixtures (Primary Approach)

**Location:** `tests/fixtures/`  
**Count:** 17 Excel files  
**Status:** Pre-generated, committed to git

**Categories:**

1. **Basic (3 files)** - Clean, simple tables
   - `simple.xlsx` - 3 columns, 10 rows, Cyrillic data
   - `with_dates.xlsx` - datetime columns
   - `numeric_types.xlsx` - int/float types, large numbers

2. **Messy (4 files)** - Real enterprise scenarios
   - `messy_headers.xlsx` - headers from row 4, junk above
   - `merged_cells.xlsx` - merged cells in headers
   - `multilevel_headers.xlsx` - 3-level header hierarchy
   - `enterprise_chaos.xlsx` - worst case: junk + merged + multi-level + formulas

3. **Edge Cases (9 files)** - Boundary conditions
   - `with_nulls.xlsx` - null/empty values
   - `with_duplicates.xlsx` - duplicate rows
   - `wide_table.xlsx` - 50 columns
   - `single_column.xlsx` - minimal structure
   - `mixed_languages.xlsx` - Cyrillic, Latin, Chinese, emojis
   - `special_chars.xlsx` - formula injection tests
   - `with_formulas.xlsx` - Excel formulas in cells
   - `complex_formatting.xlsx` - various number formats

4. **Legacy (1 file)** - Old format
   - `simple_legacy.xls` - for xlrd engine testing

### Fixture Registry

**File:** [`tests/fixtures/registry.py`](fixtures/registry.py:1)

**Purpose:** Central metadata store for all fixtures. NO hardcoded paths or expected values in tests.

**Structure:**
```python
from tests.fixtures.registry import get_fixture

fixture = get_fixture("simple")
# Access metadata:
fixture.path_str          # Full path to file
fixture.sheet_name        # "Data"
fixture.header_row        # 0 (0-based index)
fixture.columns           # ["Имя", "Возраст", "Город"]
fixture.row_count         # 10
fixture.expected          # Dict with expected values for assertions
```

**Available Functions:**
- `get_fixture(name)` - Get single fixture by name
- `get_fixtures_by_category(category)` - Get all fixtures in category
- `get_fixtures_by_test(test_name)` - Get fixtures that test specific feature

### Dynamic Fixtures (Rare Cases)

**When to use:** Only for very specific unit tests where static fixtures don't fit.

**Example:**
```python
def test_extreme_width(temp_excel_path, file_loader):
    # Create file with 1000 columns dynamically
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append([f"Col_{i}" for i in range(1000)])
    wb.save(temp_excel_path / "extreme.xlsx")
    
    # Test
    df = file_loader.load(str(temp_excel_path / "extreme.xlsx"), "Sheet")
    assert len(df.columns) == 1000
```

---

## Writing Tests

### Using Fixtures from conftest.py

**Example 1: Simple test with fixture metadata**
```python
def test_file_loading(simple_fixture, file_loader):
    """Test basic file loading."""
    # simple_fixture provides metadata from registry
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name)
    
    # Use metadata for assertions (no hardcoded values!)
    assert len(df.columns) == len(simple_fixture.columns)
    assert list(df.columns) == simple_fixture.columns
    assert len(df) == simple_fixture.row_count
```

**Example 2: Parametrized test across all basic fixtures**
```python
def test_all_basic_files(basic_fixture_meta, file_loader):
    """This test runs for EACH basic fixture automatically."""
    df = file_loader.load(basic_fixture_meta.path_str, basic_fixture_meta.sheet_name)
    assert len(df) > 0
    assert len(df.columns) > 0
```

**Example 3: Integration test**
```python
def test_filter_and_count(simple_fixture, file_loader):
    """Test filter_and_count operation end-to-end."""
    from mcp_excel.operations.data_operations import DataOperations
    from mcp_excel.models.requests import FilterAndCountRequest, FilterCondition
    
    ops = DataOperations(file_loader)
    
    request = FilterAndCountRequest(
        file_path=simple_fixture.path_str,
        sheet_name=simple_fixture.sheet_name,
        filters=[
            FilterCondition(column="Возраст", operator=">", value=30)
        ],
        logic="AND"
    )
    
    response = ops.filter_and_count(request)
    
    # Use expected values from registry
    assert response.count == simple_fixture.expected.get("age_over_30", 3)
    assert "COUNTIF" in response.excel_output.formula
```

### Available Pytest Fixtures

**Core Components (from conftest.py):**
- `file_loader` - FileLoader instance (session-scoped, cached)
- `header_detector` - HeaderDetector instance
- `datetime_detector` - DateTimeDetector instance
- `filter_engine` - FilterEngine instance
- `tsv_formatter` - TSVFormatter instance

**Fixture Metadata (individual):**
- `simple_fixture` - Simple table
- `with_dates_fixture` - Datetime columns
- `messy_headers_fixture` - Headers from row 4
- `enterprise_chaos_fixture` - Worst case scenario
- ... (see [`conftest.py`](conftest.py:1) for full list)

**Fixture Collections:**
- `basic_fixtures` - All basic fixtures (list)
- `messy_fixtures` - All messy fixtures (list)
- `edge_case_fixtures` - All edge case fixtures (list)
- `all_fixtures` - All 17 fixtures (list)

**Parametrize Helpers:**
- `fixture_meta` - Auto-parametrize with ALL fixtures
- `basic_fixture_meta` - Auto-parametrize with basic fixtures only
- `messy_fixture_meta` - Auto-parametrize with messy fixtures only
- `edge_fixture_meta` - Auto-parametrize with edge case fixtures only

**Utilities:**
- `temp_excel_path` - Temporary directory for dynamic file creation
- `assert_dataframe_equals` - Helper for comparing DataFrames
- `assert_excel_formula` - Helper for validating Excel formulas

### Test Markers

**Auto-applied markers:**
- `@pytest.mark.unit` - Fast, isolated tests (auto-applied to tests/unit/)
- `@pytest.mark.integration` - End-to-end tests (auto-applied to tests/integration/)
- `@pytest.mark.legacy` - Tests for .xls format (auto-applied when using legacy fixtures)
- `@pytest.mark.datetime` - Tests for datetime handling (auto-applied when using date fixtures)
- `@pytest.mark.slow` - Tests taking > 1 second (manual)
- `@pytest.mark.edge_case` - Edge case tests (manual)

**Usage:**
```python
@pytest.mark.slow
def test_large_file_performance(file_loader):
    # This test is marked as slow
    pass
```

---

## Running Tests

### Basic Commands

```bash
# Run all tests
pytest tests/

# Run only unit tests (fast)
pytest tests/unit/ -v

# Run only integration tests
pytest tests/integration/ -v

# Run tests for specific component
pytest tests/unit/core/test_file_loader.py -v

# Run with coverage
pytest tests/ --cov=src/mcp_excel --cov-report=html

# Run only fast tests (exclude slow)
pytest tests/ -m "not slow"

# Run only datetime-related tests
pytest tests/ -m datetime

# Run specific test by name
pytest tests/ -k "test_file_loading"

# Run in parallel (faster)
pytest tests/ -n auto
```

### Test Output

Tests use verbose output with `print()` statements for debugging. Example:
```
test_file_loader.py::test_simple_loading 
  Loading file: tests/fixtures/basic/simple.xlsx
  Rows: 10, Columns: 3
  Column names: ['Имя', 'Возраст', 'Город']
  ✅ PASSED
```

---

## Adding New Fixtures

### When to Add a New Fixture

**Add a new fixture when:**
- Testing a new edge case not covered by existing 17 files
- Need a specific structure for a new feature
- Found a real-world scenario that breaks current tests

**DON'T add a new fixture if:**
- Existing fixtures can be reused
- Can be tested with dynamic generation (`temp_excel_path`)

### Process

**Step 1: Add method to builder**

Edit [`tests/builders/generate_fixtures.py`](builders/generate_fixtures.py:1):

```python
def create_my_new_fixture_xlsx(self) -> Path:
    """Creates table with [description].
    
    Tests [what it tests].
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    
    # Create structure
    ws.append(["Header1", "Header2"])
    ws.append([1, 2])
    
    # Save to appropriate category
    output_path = self.edge_cases_dir / "my_new_fixture.xlsx"
    wb.save(output_path)
    return output_path
```

**Step 2: Add to main() in builder**

```python
def main():
    # ... existing code ...
    
    fixtures_created.append(("my_new_fixture.xlsx", builder.create_my_new_fixture_xlsx()))
    print(f"  ✅ my_new_fixture.xlsx - description")
```

**Step 3: Regenerate fixtures**

```bash
python tests/builders/generate_fixtures.py
```

**Step 4: Add metadata to registry**

Edit [`tests/fixtures/registry.py`](fixtures/registry.py:1):

```python
MY_NEW_FIXTURE = FixtureMetadata(
    name="my_new_fixture",
    category="edge_cases",
    file_name="my_new_fixture.xlsx",
    format="xlsx",
    sheet_name="Sheet1",
    header_row=0,
    columns=["Header1", "Header2"],
    row_count=1,
    description="Description of what this tests",
    tests=["feature_name", "edge_case_name"],
    expected={
        "some_value": 123,
    }
)

# Add to FIXTURES dict
FIXTURES = {
    # ... existing fixtures ...
    "my_new_fixture": MY_NEW_FIXTURE,
}
```

**Step 5: Add pytest fixture to conftest.py**

Edit [`tests/conftest.py`](conftest.py:1):

```python
@pytest.fixture
def my_new_fixture() -> FixtureMetadata:
    """Description."""
    return get_fixture("my_new_fixture")
```

**Step 6: Commit fixture file**

```bash
git add tests/fixtures/edge_cases/my_new_fixture.xlsx
git add tests/fixtures/registry.py
git add tests/conftest.py
git add tests/builders/generate_fixtures.py
git commit -m "test: add my_new_fixture for [feature]"
```

---

## Future Work

### Tests That Need to Be Written

**Priority 1: Core Components (Unit Tests)**
- [ ] `tests/unit/core/test_file_loader.py` - FileLoader (loading, caching, format detection)
- [ ] `tests/unit/core/test_header_detector.py` - HeaderDetector (auto-detection algorithm)
- [ ] `tests/unit/core/test_datetime_detector.py` - DateTimeDetector (datetime column detection)
- [ ] `tests/unit/core/test_datetime_converter.py` - DateTimeConverter (Excel number → datetime)
- [ ] `tests/unit/operations/test_filtering.py` - FilterEngine (12 operators, complex logic)
- [ ] `tests/unit/excel/test_formula_generator.py` - FormulaGenerator (all formula types)
- [ ] `tests/unit/excel/test_tsv_formatter.py` - TSVFormatter (TSV generation)

**Priority 2: Operations (Integration Tests)**
- [ ] `tests/integration/test_inspection.py` - inspect_file, get_sheet_info, get_column_names, find_column
- [ ] `tests/integration/test_data_operations.py` - get_unique_values, get_value_counts, filter_and_count, filter_and_get_rows, aggregate, group_by
- [ ] `tests/integration/test_statistics.py` - get_column_stats, correlate, detect_outliers
- [ ] `tests/integration/test_validation.py` - find_duplicates, find_nulls
- [ ] `tests/integration/test_timeseries.py` - calculate_period_change, calculate_running_total, calculate_moving_average
- [ ] `tests/integration/test_advanced.py` - rank_rows, calculate_expression
- [ ] `tests/integration/test_multisheet.py` - search_across_sheets, compare_sheets

**Priority 3: Edge Cases & Performance**
- [ ] Test all 12 filter operators with all data types
- [ ] Test formula generation for all operators
- [ ] Test datetime filtering with various formats
- [ ] Test merged cells handling
- [ ] Test multi-level headers detection
- [ ] Test performance with large files (10k+ rows)
- [ ] Test cache invalidation
- [ ] Test error handling (corrupted files, missing columns, etc.)

### Coverage Goals

**Target:** 80%+ code coverage for core components, 70%+ for operations

**Check coverage:**
```bash
pytest tests/ --cov=src/mcp_excel --cov-report=html
# Open htmlcov/index.html to see detailed report
```

### Testing Principles for Future Agents

1. **NO hardcoded values** - Always use Fixture Registry
2. **NO real data** - Only use test fixtures
3. **Deterministic** - Tests must pass consistently
4. **Fast** - Unit tests < 100ms, integration tests < 1s
5. **Isolated** - No dependencies between tests
6. **Clear** - Use descriptive names and print() for debugging
7. **Comprehensive** - Test happy path + edge cases + error cases

---

## Troubleshooting

### Common Issues

**Issue:** Fixtures not found  
**Solution:** Check that fixtures were generated: `python tests/builders/generate_fixtures.py`

**Issue:** Import errors  
**Solution:** Ensure `src/` is in path (conftest.py handles this automatically)

**Issue:** Tests fail on CI but pass locally  
**Solution:** Check that all fixtures are committed to git

**Issue:** Slow tests  
**Solution:** Use `@pytest.mark.slow` and exclude with `pytest -m "not slow"`

### Getting Help

- Read [`ARCHITECTURE.md`](../docs/ru/ARCHITECTURE.md:1) for system design
- Check [`test_manual.py`](../test_manual.py:1) for manual testing examples
- Look at existing tests for patterns

---

## Summary for AI Agents

**Before writing tests:**
1. Read this README completely
2. Check [`registry.py`](fixtures/registry.py:1) for available fixtures
3. Check [`conftest.py`](conftest.py:1) for available pytest fixtures
4. Look at existing tests for patterns

**When writing tests:**
1. Use fixtures from registry (NO hardcoded paths/values)
2. Use pytest fixtures from conftest.py
3. Add markers (`@pytest.mark.unit`, etc.)
4. Use `print()` for debugging output
5. Test happy path + edge cases + errors

**After writing tests:**
1. Run tests: `pytest tests/ -v`
2. Check coverage: `pytest tests/ --cov=src/mcp_excel`
3. Ensure tests are fast (< 1s for integration, < 100ms for unit)
4. Update this README if you added new infrastructure

**Remember:** This testing infrastructure is designed for scalability. Follow the patterns, don't reinvent the wheel.
