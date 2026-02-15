# Testing Infrastructure Documentation

**For AI Agents:** This document describes the complete testing infrastructure for MCP Excel Server. Read this before writing any tests.

---

## Table of Contents

- [Testing Philosophy](#testing-philosophy)
- [Test Organization Strategy](#test-organization-strategy)
- [Test Fixtures](#test-fixtures)
- [Writing Tests](#writing-tests)
- [Running Tests](#running-tests)
- [Adding New Fixtures](#adding-new-fixtures)
- [Future Work](#future-work)
- [Troubleshooting](#troubleshooting)

---

## Testing Philosophy

### Core Principles

Tests must be **deterministic, fast, and use NO real data**. Our testing approach is built on:

**Static Fixtures:** 17 pre-generated Excel files covering all scenarios (basic, messy, edge cases, legacy). These files are committed to git and provide consistent, reproducible test data.

**Fixture Registry:** Central metadata store in `tests/fixtures/registry.py` - NO hardcoded paths or expected values in tests. All test data comes from the registry.

**Pytest Fixtures:** Reusable components in `tests/conftest.py` provide access to core components (FileLoader, HeaderDetector, etc.) and fixture metadata.

**Separation of Concerns:** Unit tests (isolated, fast) vs Integration tests (end-to-end, slower).

### What We Test

**Core Components:**
- FileLoader - loading .xls/.xlsx, caching
- HeaderDetector - auto-detection of header rows
- DateTimeDetector - datetime column detection
- FilterEngine - 12 filter operators
- FormulaGenerator - Excel formula generation

**Operations:** All 23 MCP tools (inspection, data operations, statistics, validation, etc.)

---

## Test Organization Strategy

### Grouping Principle: By FUNCTIONALITY, Not Code Location

Tests are organized by **what they do**, not where the code lives. This prevents file bloat and makes it clear where new tests should go.

**Example:** `test_filtering_and_counting.py` contains ALL tests for `filter_and_count` with ALL 12 operators (~300-500 lines). It doesn't matter that FilterEngine is in `operations/filtering.py` - what matters is that all filtering tests are together.

**Scalability:** When you add a new filter operator, it goes in `test_filtering_and_counting.py`. When you add a new aggregation function, it goes in `test_aggregation.py`. Future agents won't create `test_new_filter_fix.py` because the file name clearly indicates its purpose.

### Unit Tests Structure

Unit tests focus on isolated components. Approximately 7 files organized by component type:

**Core Components** (`tests/unit/core/`):
- `test_file_loader.py` - FileLoader: loading, caching, format detection, datetime conversion
- `test_header_detector.py` - HeaderDetector: auto-detection algorithm, confidence scoring
- `test_datetime_detector.py` - DateTimeDetector: datetime column identification
- `test_datetime_converter.py` - DateTimeConverter: Excel number to datetime conversion

**Operations** (`tests/unit/operations/`):
- `test_filter_engine.py` - FilterEngine: all 12 operators (==, !=, >, <, >=, <=, in, not_in, contains, startswith, endswith, regex, is_null, is_not_null), complex logic (AND/OR)

**Excel Utilities** (`tests/unit/excel/`):
- `test_formula_generator.py` - FormulaGenerator: formula generation for all operators and operations
- `test_tsv_formatter.py` - TSVFormatter: TSV generation for Excel paste

### Integration Tests Structure

Integration tests verify end-to-end functionality. Approximately 9 files organized by feature area:

**File Inspection** (`test_file_inspection.py`):
- Tools: `inspect_file`, `get_sheet_info`, `get_column_names`, `get_data_profile`
- Tests: file metadata, column detection, data profiling, header auto-detection

**Data Retrieval** (`test_data_retrieval.py`):
- Tools: `get_unique_values`, `get_value_counts`, `filter_and_get_rows`
- Tests: unique value extraction, frequency counts, filtered row retrieval with pagination

**Filtering and Counting** (`test_filtering_and_counting.py`):
- Tools: `filter_and_count`
- Tests: ALL 12 filter operators, combined filters (AND/OR), datetime filtering, formula generation

**Aggregation** (`test_aggregation.py`):
- Tools: `aggregate` (8 operations: sum, mean, median, min, max, std, var, count), `group_by`
- Tests: all aggregation operations, filtered aggregation, multi-column grouping, formula generation

**Statistics** (`test_statistics.py`):
- Tools: `get_column_stats`, `correlate`, `detect_outliers`
- Tests: statistical summaries, correlation matrices, outlier detection (IQR and Z-score methods)

**Validation** (`test_validation.py`):
- Tools: `find_duplicates`, `find_nulls`
- Tests: duplicate detection (single and multi-column), null value identification

**Multi-Sheet Operations** (`test_multisheet.py`):
- Tools: `find_column`, `search_across_sheets`, `compare_sheets`
- Tests: cross-sheet search, column location, sheet comparison

**Time Series** (`test_timeseries.py`):
- Tools: `calculate_period_change`, `calculate_running_total`, `calculate_moving_average`
- Tests: period-over-period analysis, cumulative calculations, moving averages

**Advanced Operations** (`test_advanced.py`):
- Tools: `rank_rows`, `calculate_expression`
- Tests: ranking (top-N, grouped), expression evaluation, formula generation

### Why This Structure Works

**Prevents File Bloat:** Each file stays under 600 lines by focusing on a specific feature area, not on mirroring code structure.

**Clear Ownership:** It's immediately obvious where tests for a feature should go. No ambiguity.

**Scalability:** Adding new features doesn't create new test files unless it's a genuinely new feature area.

**Maintainability:** Related tests are together, making it easier to understand feature coverage.

---

## Test Fixtures

### Static Fixtures Overview

We use 17 pre-generated Excel files committed to git. These provide consistent, reproducible test data.

**Categories:**

**Basic Fixtures** (4 files in `tests/fixtures/basic/`):
- `simple.xlsx` - Clean 3-column table with Cyrillic data (10 rows)
- `with_dates.xlsx` - Table with datetime columns (15 rows)
- `numeric_types.xlsx` - Different numeric types, large integers (20 rows)
- `multi_sheet.xlsx` - File with 3 sheets (Products, Clients, Orders) for multi-sheet testing

**Messy Fixtures** (4 files in `tests/fixtures/messy/`):
- `messy_headers.xlsx` - Headers from row 4, junk above (20 rows)
- `merged_cells.xlsx` - Merged cells in headers (5 rows)
- `multilevel_headers.xlsx` - 3-level header hierarchy (10 rows)
- `enterprise_chaos.xlsx` - Worst case: junk + merged + multi-level + formulas (5 rows)

**Edge Cases** (8 files in `tests/fixtures/edge_cases/`):
- `with_nulls.xlsx` - Null/empty values (10 rows)
- `with_duplicates.xlsx` - Duplicate rows (8 rows)
- `wide_table.xlsx` - 50 columns (10 rows)
- `single_column.xlsx` - Minimal structure (10 rows)
- `mixed_languages.xlsx` - Cyrillic, Latin, Chinese, emojis (8 rows)
- `special_chars.xlsx` - Formula injection tests (10 rows)
- `with_formulas.xlsx` - Excel formulas in cells (5 rows)
- `complex_formatting.xlsx` - Various number formats (10 rows)

**Legacy Format** (1 file in `tests/fixtures/legacy/`):
- `simple_legacy.xls` - Legacy format for xlrd engine testing (5 rows)

### Fixture Registry

The registry (`tests/fixtures/registry.py`) provides structured metadata about each fixture. NO hardcoded paths or expected values in tests.

**Usage Example:**
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

### Dynamic Fixtures (Rare)

Use `temp_excel_path` fixture for very specific unit tests where static fixtures don't fit. This should be rare - prefer static fixtures.

---

## Writing Tests

### Using Fixtures from conftest.py

**Example 1: Simple test with fixture metadata**
```python
def test_file_loading(simple_fixture, file_loader):
    """Test basic file loading."""
    df = file_loader.load(simple_fixture.path_str, simple_fixture.sheet_name, header_row=0)
    
    # Use metadata for assertions (no hardcoded values!)
    assert len(df.columns) == len(simple_fixture.columns)
    assert list(df.columns) == simple_fixture.columns
    assert len(df) == simple_fixture.row_count
```

**Example 2: Parametrized test across all basic fixtures**
```python
def test_all_basic_files(basic_fixture_meta, file_loader):
    """This test runs for EACH basic fixture automatically."""
    df = file_loader.load(basic_fixture_meta.path_str, basic_fixture_meta.sheet_name, header_row=basic_fixture_meta.header_row)
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
        filters=[FilterCondition(column="Возраст", operator=">", value=30)],
        logic="AND"
    )
    
    response = ops.filter_and_count(request)
    
    # Use expected values from registry
    assert response.count == simple_fixture.expected.get("age_over_30", 3)
    assert "COUNTIF" in response.excel_output.formula
```

### Available Pytest Fixtures

**Core Components** (from `conftest.py`):
- `file_loader` - FileLoader instance (session-scoped, cached)
- `header_detector` - HeaderDetector instance
- `datetime_detector` - DateTimeDetector instance
- `filter_engine` - FilterEngine instance
- `tsv_formatter` - TSVFormatter instance

**Fixture Metadata** (individual):
- `simple_fixture`, `with_dates_fixture`, `numeric_types_fixture`, `multi_sheet_fixture`
- `messy_headers_fixture`, `merged_cells_fixture`, `multilevel_headers_fixture`, `enterprise_chaos_fixture`
- `with_nulls_fixture`, `with_duplicates_fixture`, `wide_table_fixture`, `single_column_fixture`
- `mixed_languages_fixture`, `special_chars_fixture`, `with_formulas_fixture`, `complex_formatting_fixture`
- `simple_legacy_fixture`

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

Edit `tests/builders/generate_fixtures.py`:

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

Edit `tests/fixtures/registry.py`:

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

Edit `tests/conftest.py`:

```python
@pytest.fixture
def my_new_fixture() -> FixtureMetadata:
    """Description."""
    return get_fixture("my_new_fixture")
```

**Step 6: Update smoke test**

Edit `tests/test_smoke.py` to update fixture count.

**Step 7: Commit fixture file**

```bash
git add tests/fixtures/edge_cases/my_new_fixture.xlsx
git add tests/fixtures/registry.py
git add tests/conftest.py
git add tests/builders/generate_fixtures.py
git add tests/test_smoke.py
git commit -m "test: add my_new_fixture for [feature]"
```

---

## Future Work

### Tests That Need to Be Written

**Priority 1: Core Components (Unit Tests)**
- `tests/unit/core/test_file_loader.py` - ✅ DONE (27 tests)
- `tests/unit/core/test_header_detector.py` - ✅ DONE (24 tests)
- `tests/unit/core/test_datetime_detector.py` - ✅ DONE (16 tests)
- `tests/unit/core/test_datetime_converter.py` - ✅ DONE (14 tests)
- `tests/unit/operations/test_filter_engine.py` - ✅ DONE (46 tests)
- `tests/unit/excel/test_formula_generator.py` - ✅ DONE (40 tests)
- `tests/unit/excel/test_tsv_formatter.py` - ✅ DONE (26 tests)

**Priority 2: Operations (Integration Tests)**
- `tests/integration/test_file_inspection.py` - ✅ DONE (37 tests) - inspect_file, get_sheet_info, get_column_names, get_data_profile
- `tests/integration/test_data_retrieval.py` - ✅ DONE (48 tests) - get_unique_values, get_value_counts, filter_and_get_rows
- `tests/integration/test_filtering_and_counting.py` - ✅ DONE (26 tests) - filter_and_count with ALL 12 operators, combined filters (AND/OR), datetime filtering, edge cases
- `tests/integration/test_aggregation.py` - aggregate (8 operations), group_by
- `tests/integration/test_statistics.py` - get_column_stats, correlate, detect_outliers
- `tests/integration/test_validation.py` - find_duplicates, find_nulls
- `tests/integration/test_multisheet.py` - find_column, search_across_sheets, compare_sheets
- `tests/integration/test_timeseries.py` - calculate_period_change, running_total, moving_average
- `tests/integration/test_advanced.py` - rank_rows, calculate_expression

**Priority 3: Edge Cases & Performance**
- Test all 12 filter operators with all data types
- Test formula generation for all operators
- Test datetime filtering with various formats
- Test merged cells handling
- Test multi-level headers detection
- Test performance with large files (10k+ rows)
- Test cache invalidation
- Test error handling (corrupted files, missing columns, etc.)

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

- Read `docs/ru/ARCHITECTURE.md` for system design
- Check `test_manual.py` for manual testing examples
- Look at existing tests for patterns

---

## Summary for AI Agents

**Before writing tests:**
1. Read this README completely
2. Check `tests/fixtures/registry.py` for available fixtures
3. Check `tests/conftest.py` for available pytest fixtures
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
