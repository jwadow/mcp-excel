# Excel MCP Server

MCP server for Excel file operations using atomic primitives. Enables AI agents to analyze Excel files through composable operations without loading raw data into context.

Made with ❤️ by [Jwadow](https://github.com/jwadow)

## Philosophy

- **Atomic Operations**: Agent combines primitive operations instead of loading entire datasets
- **Stateless Architecture**: No session management - automatic LRU caching based on file path
- **Dynamic Results**: Generates Excel formulas for copy-paste, not static values
- **Read-Only**: Current version doesn't modify files (safe for legacy .xls format)
- **Universal**: Works with any tabular data without domain-specific hardcoding

## Features

- ✅ Automatic header detection for messy Excel files
- ✅ LRU caching with memory management
- ✅ Support for both .xls and .xlsx formats
- ✅ Excel formula generation for dynamic calculations
- ✅ TSV output for easy copy-paste into Excel
- ✅ Comprehensive filtering system
- ✅ Performance metrics for every operation

## Installation

### Prerequisites

- Python 3.10 or higher
- Poetry (recommended) or pip

### Using Poetry (Recommended)

```bash
# Install dependencies
poetry install

# Activate virtual environment
poetry shell
```

### Using pip

```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On Linux/Mac:
source venv/bin/activate

# Install dependencies
pip install -e .
```

## Manual Testing

Test the core functionality without running the MCP server.

### Option 1: Direct execution (no installation needed)

```bash
# Install only dependencies
pip install pandas pydantic xlrd openpyxl psutil

# Run test with your Excel file
python test_manual.py C:/Users/YourName/Documents/data.xlsx
```

### Option 2: With Poetry

```bash
# Install project
poetry install

# Run test
poetry run python test_manual.py C:/Users/YourName/Documents/data.xlsx
```

**Important:** Replace `C:/Users/YourName/Documents/data.xlsx` with the actual path to your Excel file.

### What the manual test does:

1. **FileLoader Test**: Loads file, shows structure, demonstrates caching
2. **HeaderDetector Test**: Automatically detects header row in messy files
3. **InspectionOperations Test**: Shows file inspection and sheet analysis
4. **DataOperations Test**: Tests filtering, unique values, and data retrieval (5 tests)
5. **AggregationOperations Test**: Tests aggregation and group-by operations (6 tests)

### Example Output:

```
================================================================================
  Testing FileLoader
================================================================================

📁 File Information:
{
  "format": "xlsx",
  "size_mb": 2.45,
  "sheet_count": 3,
  "sheet_names": ["Sales", "Inventory", "Archive"]
}

📋 Sheet Names:
  1. Sales
  2. Inventory
  3. Archive

📊 Loading first sheet: Sales
  Rows: 1523
  Columns: 12
  Column names: ['Date', 'Customer', 'Product', 'Quantity', 'Price', ...]

💾 Cache Statistics:
{
  "size": 1,
  "max_size": 5,
  "memory_mb": 145.2,
  "idle_seconds": 0.5
}
```

## Running as MCP Server

### Configuration for Claude Desktop

Add to your Claude Desktop config (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "/path/to/mcp-excel"
    }
  }
}
```

### Configuration for Other MCP Clients

The server uses STDIO transport by default. Start it with:

```bash
poetry run python -m mcp_excel.main
```

## Available Tools

### 1. `inspect_file`

Get basic information about Excel file structure.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx"
}
```

**Output:**
- File format and size
- List of all sheets
- Row/column counts for each sheet
- Performance metrics

### 2. `get_sheet_info`

Get detailed information about a specific sheet.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "header_row": null  // Optional: auto-detected if not provided
}
```

**Output:**
- Column names and types
- Row count
- Sample data (first 3 rows)
- Header detection info (if auto-detected)
- Performance metrics

### 3. `get_column_names`

Quick operation to get just the column names.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales"
}
```

**Output:**
- List of column names
- Column count

### 4. `get_unique_values`

Get unique values from a column (useful for building filters).

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "column": "Customer",
  "limit": 100
}
```

**Output:**
- List of unique values
- Count of unique values
- Truncated flag if limit exceeded

### 5. `get_value_counts`

Get frequency counts for values in a column (top N most common).

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "column": "Status",
  "top_n": 10
}
```

**Output:**
- Dictionary of value -> count
- Total number of values
- TSV output for Excel

### 6. `filter_and_count`

Count rows matching filter conditions.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "filters": [
    {"column": "Customer", "operator": "==", "value": "Acme Corp"},
    {"column": "Amount", "operator": ">", "value": 1000}
  ],
  "logic": "AND"
}
```

**Output:**
- Count of matching rows
- Excel formula (e.g., `=COUNTIFS(...)`)
- Applied filters

**Supported operators:** `==`, `!=`, `>`, `<`, `>=`, `<=`, `in`, `not_in`, `contains`, `startswith`, `endswith`, `regex`, `is_null`, `is_not_null`

### 7. `filter_and_get_rows`

Get rows matching filter conditions with pagination.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "filters": [
    {"column": "Status", "operator": "==", "value": "Active"}
  ],
  "columns": ["Customer", "Amount", "Date"],
  "limit": 50,
  "offset": 0,
  "logic": "AND"
}
```

**Output:**
- Filtered rows as list of dictionaries
- Total matches count
- Truncated flag
- TSV output for Excel

### 8. `aggregate`

Perform aggregation (sum, mean, count, etc.) on a column with optional filters.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "operation": "sum",
  "target_column": "Amount",
  "filters": [
    {"column": "Customer", "operator": "==", "value": "Acme Corp"}
  ]
}
```

**Output:**
- Aggregated value
- Excel formula (e.g., `=SUMIF(...)`)
- Applied filters

**Supported operations:** `sum`, `mean`, `median`, `min`, `max`, `std`, `var`, `count`

**Special feature:** Automatically converts text-stored numbers to numeric (common Excel issue).

### 9. `group_by`

Group data by columns and perform aggregation (like Excel Pivot Table).

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "group_columns": ["Customer", "Month"],
  "agg_column": "Amount",
  "agg_operation": "sum"
}
```

**Output:**
- Grouped data with aggregated values
- TSV output for Excel
- Supports multiple grouping columns

### 10. `find_column`

Find a column across all sheets or in a specific sheet.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "column_name": "Customer",
  "search_all_sheets": true
}
```

**Output:**
- List of sheets where column was found
- Column index and row count for each match
- Case-insensitive search

### 11. `search_across_sheets`

Search for a specific value across all sheets in the file.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "column_name": "Customer",
  "value": "Acme Corp"
}
```

**Output:**
- List of sheets with matches
- Match count per sheet
- Total matches across all sheets
- Supports both numeric and string values

### 12. `compare_sheets`

Compare data between two sheets using a key column.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet1": "January",
  "sheet2": "February",
  "key_column": "Customer",
  "compare_columns": ["Amount", "Quantity"]
}
```

**Output:**
- Rows with differences
- Status for each row (only_in_sheet1, only_in_sheet2, different_values)
- Side-by-side comparison of values
- TSV output for Excel

### 13. `get_column_stats`

Get statistical summary of a column.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "column": "Amount"
}
```

**Output:**
- Count, mean, median, std, min, max
- Quartiles (25th, 75th percentile)
- Null count
- TSV output for Excel

### 14. `correlate`

Calculate correlation matrix between multiple columns.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "columns": ["Amount", "Quantity", "Discount"],
  "method": "pearson"
}
```

**Output:**
- Correlation matrix
- Supports pearson, spearman, kendall methods
- Works with 2+ columns
- TSV output for Excel

### 15. `detect_outliers`

Detect outliers in a column using IQR or Z-score method.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "column": "Amount",
  "method": "iqr",
  "threshold": 1.5
}
```

**Output:**
- List of rows with outliers
- Outlier count
- Method and threshold used
- TSV output for Excel

### 16. `find_duplicates`

Find duplicate rows based on specified columns.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "columns": ["Customer", "Date"]
}
```

**Output:**
- List of duplicate rows (all occurrences including first)
- Duplicate count
- Columns checked
- TSV output for Excel
- Row indices for each duplicate

**Note:** Uses `duplicated(keep=False)` to mark all duplicates including first occurrence.

### 17. `find_nulls`

Find null/empty values in specified columns with detailed statistics.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "columns": ["Customer", "Amount", "Date"]
}
```

**Output:**
- Null statistics per column (count, percentage, indices)
- Total null count across all checked columns
- TSV output for Excel
- First 100 null indices per column

## Architecture

```
mcp-excel/
├── src/mcp_excel/
│   ├── core/              # Core functionality
│   │   ├── cache.py       # LRU cache with memory management
│   │   ├── file_loader.py # File loading with format detection
│   │   └── header_detector.py # Intelligent header detection
│   ├── models/            # Pydantic schemas
│   │   ├── requests.py    # Request models
│   │   └── responses.py   # Response models
│   ├── operations/        # Business logic
│   │   ├── inspection.py     # File/sheet inspection
│   │   ├── data_operations.py # Data filtering and aggregation
│   │   └── filtering.py      # Filter engine
│   ├── excel/             # Excel-specific functionality
│   │   ├── formula_generator.py # Excel formula generation
│   │   └── tsv_formatter.py     # TSV formatting
│   └── main.py            # MCP server entry point
├── tests/                 # Test suite
├── test_manual.py         # Manual testing script
└── pyproject.toml         # Dependencies and config
```

## Development

### Running Tests

```bash
# Run all tests
poetry run pytest

# Run with coverage
poetry run pytest --cov=mcp_excel

# Run specific test file
poetry run pytest tests/test_operations.py
```

### Code Quality

```bash
# Format code
poetry run black src/ tests/

# Lint code
poetry run ruff check src/ tests/

# Type checking
poetry run mypy src/
```

## Roadmap

### Phase 1: Core Operations ✅ COMPLETED
- ✅ File inspection
- ✅ Sheet analysis
- ✅ Column operations
- ✅ Filtering and counting
- ✅ Basic aggregations
- ✅ Data retrieval with pagination
- ✅ Unique values and frequency analysis

### Phase 2: Advanced Analytics ✅ COMPLETED
- ✅ Group-by operations
- ✅ Statistical analysis (get_column_stats)
- ✅ Correlation analysis (correlate)
- ✅ Outlier detection (detect_outliers)

### Phase 3: Multi-Sheet Operations ✅ COMPLETED
- ✅ Cross-sheet column search (find_column)
- ✅ Cross-sheet value search (search_across_sheets)
- ✅ Sheet comparison (compare_sheets)

### Phase 4: Data Validation ✅ COMPLETED
- ✅ Find duplicates (find_duplicates)
- ✅ Find null values (find_nulls)

### Phase 5: Future Enhancements
- ⏳ Write operations (xlsx only)
- ⏳ CSV support
- ⏳ SSE transport mode
- ⏳ Advanced formula generation

## License

This project is licensed under the GNU Affero General Public License v3.0 (AGPL-3.0).

See [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please ensure:

1. All dependencies are AGPL-compatible
2. Code follows the existing style (black, ruff, mypy)
3. Tests are included for new features
4. Documentation is updated

## Support

For issues, questions, or contributions, please open an issue on GitHub.
