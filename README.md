# MCP Excel Server

MCP server for Excel file operations using atomic primitives. Enables AI agents to analyze Excel files through composable operations without loading raw data into context.

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
│   │   ├── inspection.py  # File/sheet inspection
│   │   └── filtering.py   # Filter engine
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

### Phase 1: Core Operations (Current)
- ✅ File inspection
- ✅ Sheet analysis
- ✅ Column operations
- ⏳ Filtering and counting
- ⏳ Basic aggregations

### Phase 2: Advanced Analytics
- ⏳ Group-by operations
- ⏳ Statistical analysis
- ⏳ Correlation analysis
- ⏳ Outlier detection

### Phase 3: Multi-Sheet Operations
- ⏳ Cross-sheet search
- ⏳ Sheet comparison
- ⏳ Data validation

### Phase 4: Future Enhancements
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
