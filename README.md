<div align="center">

# 📊 Excel MCP Server

**MCP server for Excel file operations using atomic primitives**

Made with ❤️ by [@Jwadow](https://github.com/jwadow)

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)
[![Tools](https://img.shields.io/badge/tools-22-orange.svg)](#available-tools)

Enables AI agents to analyze Excel files through composable operations without loading raw data into context

*Works with Claude Code, OpenCode, Codex app, Cursor, Cline, Roo Code, Kilo Code and other MCP-compatible AI agents*

[What You Can Do](#what-you-can-do) • [Installation](#installation) • [Configuration](#configuration) • [Available Tools](#available-tools) • [💖 Sponsor](#-support-the-project)

</div>

---

## What You Can Do

- 📊 **Analyze any Excel file** (.xls and .xlsx) without opening Excel
- 🔍 **Filter and search** data with 12 operators (==, !=, >, <, in, contains, regex, etc.)
- 📈 **Aggregate and group** data (sum, average, count, pivot tables)
- 📉 **Statistical analysis** (correlations, outliers, distributions)
- 📅 **Time series analysis** (period-over-period growth, moving averages, running totals)
- 🏆 **Rank and sort** (top-N, bottom-N, percentiles)
- ✅ **Validate data** (find duplicates, null values)
- 🔄 **Compare sheets** (find differences between versions)
- 📋 **Copy results to Excel** - generates formulas and TSV for instant paste
- 🤖 **Works with any AI agent** - Claude Code, Cline, Roo Code, Cursor, and more

## Prerequisites

- Python 3.10 or higher
- Poetry (recommended) or pip

## Installation

```bash
# Clone the repository
git clone https://github.com/jwadow/mcp-excel.git
cd mcp-excel
```

Then install dependencies using one of these methods:

**Option A: Using Poetry (Recommended)**
```bash
poetry install
```

**Option B: Using pip**
```bash
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Linux/Mac
pip install -e .
```

## Configuration

⚠️ **Important:** This is an MCP server. It runs automatically when your AI agent needs it. Do not run it manually in terminal.

### Supported AI Agents

Works with any MCP-compatible AI agent: *Claude Code, OpenCode, Codex app, Cursor, Cline, Roo Code, Kilo Code*

### Configuration Steps

1. Open your AI agent's MCP settings
2. Add new MCP server with these parameters:
   - **Command:** `python`
   - **Args:** `["-m", "mcp_excel.main"]`
   - **Working Directory:** `/path/to/mcp-excel` (replace with actual path)

**Example JSON configuration** (if your agent uses JSON config):
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

## Usage

After configuration, restart your AI agent and ask it to analyze Excel files:

```
"Analyze the Excel file at C:/Users/YourName/Documents/sales.xls"
"Show me top 10 customers by revenue from sales.xlsx"
"Find duplicates in column 'Email' in contacts.xlsx"
"Calculate month-over-month growth from revenue.xls"
```

## Manual Testing (Optional)

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

### 18. `calculate_period_change`

Calculate period-over-period change (month/quarter/year growth).

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "date_column": "Date",
  "value_column": "Revenue",
  "period_type": "month"
}
```

**Output:**
- Periods with values and changes (absolute and percentage)
- Excel formula for percentage change
- TSV output for Excel

**Period types:** `month`, `quarter`, `year`

**Use case:** "Show month-over-month revenue growth"

### 19. `calculate_running_total`

Calculate running total (cumulative sum) ordered by a column.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "order_column": "Date",
  "value_column": "Revenue",
  "group_by_columns": ["Region"]
}
```

**Output:**
- Rows with running totals
- Excel formula (e.g., `=SUM($B$2:B2)`)
- TSV output for Excel
- Supports grouping (running total within groups)

**Use case:** "Calculate cumulative revenue by date"

### 20. `calculate_moving_average`

Calculate moving average with specified window size.

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "order_column": "Date",
  "value_column": "Revenue",
  "window_size": 7
}
```

**Output:**
- Rows with moving averages
- Excel formula (e.g., `=AVERAGE(B1:B7)`)
- TSV output for Excel

**Use case:** "7-day moving average of daily sales"

### 21. `rank_rows`

Rank rows by column value (ascending or descending).

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "rank_column": "Revenue",
  "direction": "desc",
  "top_n": 10,
  "group_by_columns": ["Region"]
}
```

**Output:**
- Ranked rows with rank numbers
- Excel formula (e.g., `=RANK(B2,$B$2:$B$100,0)`)
- TSV output for Excel
- Supports top-N filtering
- Supports ranking within groups

**Use case:** "Top 10 customers by revenue in each region"

### 22. `calculate_expression`

Calculate expression between columns (arithmetic operations).

**Input:**
```json
{
  "file_path": "/path/to/file.xlsx",
  "sheet_name": "Sales",
  "expression": "Price * Quantity",
  "output_column_name": "Total"
}
```

**Output:**
- Rows with calculated values
- Excel formula (e.g., `=A2*B2`)
- TSV output for Excel

**Supported operations:** `+`, `-`, `*`, `/`, parentheses

**Use cases:**
- "Calculate revenue = Price * Quantity"
- "Calculate margin = (Revenue - Cost) / Revenue"
- "Calculate average speed = Distance / Time"

## Roadmap

- Write operations (xlsx only)
- CSV support
- SSE transport mode
- Advanced formula generation

---

## 📜 License

This project is licensed under the **GNU Affero General Public License v3.0 (AGPL-3.0)**.

This means:
- ✅ You can use, modify, and distribute this software
- ✅ You can use it for commercial purposes
- ⚠️ **You must disclose source code** when you distribute the software
- ⚠️ **Network use is distribution** — if you run a modified version on a server and let others interact with it, you must make the source code available
- ⚠️ Modifications must be released under the same license

See the [LICENSE](LICENSE) file for the full license text.

### Why AGPL-3.0?

AGPL-3.0 ensures that improvements to this software benefit the entire community. If you modify this server and deploy it as a service, you must share your improvements with your users.

---

## 💖 Support the Project

<div align="center">

<img src="https://raw.githubusercontent.com/Tarikul-Islam-Anik/Animated-Fluent-Emojis/master/Emojis/Smilies/Smiling%20Face%20with%20Hearts.png" alt="Love" width="80" />

**If this project saved you time or money, consider supporting it!**

Every contribution helps keep this project alive and growing

<br>

### 🤑 Donate

[**☕ One-time Donation**](https://app.lava.top/jwadow?tabId=donate) &nbsp;•&nbsp; [**💎 Monthly Support**](https://app.lava.top/jwadow?tabId=subscriptions)

<br>

### 🪙 Or send crypto

| Currency | Network | Address |
|:--------:|:-------:|:--------|
| **USDT** | TRC20 | `TSVtgRc9pkC1UgcbVeijBHjFmpkYHDRu26` |
| **BTC** | Bitcoin | `12GZqxqpcBsqJ4Vf1YreLqwoMGvzBPgJq6` |
| **ETH** | Ethereum | `0xc86eab3bba3bbaf4eb5b5fff8586f1460f1fd395` |
| **SOL** | Solana | `9amykF7KibZmdaw66a1oqYJyi75fRqgdsqnG66AK3jvh` |
| **TON** | TON | `UQBVh8T1H3GI7gd7b-_PPNnxHYYxptrcCVf3qQk5v41h3QTM` |

</div>

---

## Contributing

Contributions are welcome! Please ensure:

1. All dependencies are AGPL-compatible
2. Code follows the existing style
3. Tests are included for new features
4. Documentation is updated

For issues, questions, or contributions, please open an issue on GitHub.

---

<div align="center">

**[⬆ Back to Top](#-excel-mcp-server)**

</div>
