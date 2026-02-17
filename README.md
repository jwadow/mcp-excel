<div align="center">

# 📊 Excel MCP Server

**Fast and efficient spreadsheet analysis through atomic operations, built specifically for AI agents**

Made with ❤️ by [@Jwadow](https://github.com/jwadow)

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)
[![Sponsor](https://img.shields.io/badge/💖_Sponsor-Support_Development-ff69b4)](#-support-the-project)

**Analyze Excel spreadsheets with your AI agent through atomic operations — no data dumping into AI context**

*Works with OpenCode, Claude Code, Codex app, Cursor, Cline, Roo Code, Kilo Code and other MCP-compatible AI agents*

[Why This Exists](#-why-this-exists) • [What Your Agent Can Do](#-what-your-agent-can-do) • [Installation & Configuration](#%EF%B8%8F-installation--configuration) • [Available Tools](#%EF%B8%8F-available-tools) • [💖 Donate](#-support-the-project)

</div>

---

## 🤨 Why This Exists

**The Problem:** Most Excel tools for AI dump raw spreadsheet data into the agent's context. This floods the context window, slows everything down, and the AI can still miscalculate or get confused in large datasets.

**This Project:** Think SQL for Excel. Your AI agent composes atomic operations (`filter_and_count`, `aggregate`, `group_by`) and gets back precise results — not thousands of rows.

The agent analyzes data **without seeing it**. Results come as numbers, formulas, and insights.

> *"This is like working with a database through SQL, not dragging everything into memory."*
> — AI Agent after analyzing a production spreadsheet

### 🔌 What is MCP?

[Model Context Protocol](https://modelcontextprotocol.io) is an open standard that lets AI agents use external tools.

This server is such a tool. When you connect it to your AI agent (Claude Desktop, Cline, Roo Code, Cursor, etc.), your agent gets 25 new commands for working with Excel files — filtering, counting, aggregating, analyzing.

**The key benefit:** Your AI doesn't load thousands of spreadsheet rows into its memory. Instead, it asks specific questions and gets precise answers. Faster, more accurate, no context overflow.

---

## 💬 What AI Agents Say About It

Real feedback from AI agents that used this MCP server in production:

> *"Analyzed 34,211 rows without loading data into context. Every operation returns just the result — count, sum, average. Context stays clean. Operations execute in 25-45ms regardless of file size."*

> *"This is SQL for Excel. Query, filter, aggregate—without dumping data into context. Solid tool for analytical tasks."*

> *"The filter system handles complex logic well. Nested AND/OR groups, 12 operators, unlimited conditions. Built a multi-category classification without writing code."*

> *"Batch operations are efficient. One `filter_and_count_batch` call instead of multiple separate requests. File loads once, all filters apply, results come back together."*

*Yes, agents write reviews now. These are actual reflections from AI agents analyzing real-world spreadsheet data. Welcome to 2026.*

---

## 🚀 What Your Agent Can Do

Once connected, your AI agent gets a lot of specialized tools for analyzing spreadsheet data. The agent receives only precise queries and reliable results.

### 📊 Data Exploration
- **Inspect files** - structure, sheets, columns, data types (auto-detects messy headers)
- **Profile columns** - statistics, null counts, top values, data quality in one call
- **Find data** - search across multiple sheets, locate columns anywhere

### 🔍 Filtering & Querying
- **12 filter operators** - `==`, `!=`, `>`, `<`, `>=`, `<=`, `in`, `not_in`, `contains`, `startswith`, `endswith`, `regex`
- **Complex logic** - nested AND/OR groups, NOT operator, unlimited conditions
- **Batch operations** - classify data into multiple categories in one request (6x faster)
- **Overlap analysis** - Venn diagrams, intersection counts, set operations

### 📈 Aggregation & Analysis
- **8 aggregation functions** - sum, mean, median, min, max, std, var, count
- **Group by** - pivot tables with multiple grouping columns
- **Statistical analysis** - correlations (Pearson/Spearman/Kendall), outlier detection (IQR/Z-score)
- **Time series** - period-over-period growth, moving averages, running totals

### 🏆 Advanced Operations
- **Ranking** - top-N, bottom-N, percentile ranking (with grouping support)
- **Calculated columns** - arithmetic expressions between columns
- **Data validation** - find duplicates, null values, data quality checks
- **Sheet comparison** - diff between versions, find changes

### ⚡ Performance Features
- **Atomic operations** - results in 20-50ms, no matter the file size
- **Smart caching** - file loaded once, reused for all operations
- **Sample rows** - preview filtered data without full retrieval
- **Context protection** - smart limits prevent AI context overflow

### 📋 Excel Integration
- **Formula generation** - every result includes Excel formula for dynamic updates
- **TSV output** - copy-paste results directly into Excel
- **Legacy support** - works with old .xls files (Excel 97-2003)
- **Multi-sheet** - analyze across multiple sheets in one file

**Example queries your agent can now handle:**
- *"Show me top 10 customers by revenue"*
- *"Find all orders from Q4 where amount > $1000"*
- *"Calculate month-over-month growth for each product category"*
- *"Which customers are both VIP and active? (overlap analysis)"*
- *"Find duplicates in the email column"*

## ⚙️ Installation & Configuration

### Prerequisites

**Python 3.10 or higher** — [Download here](https://www.python.org/downloads/)

### Step 1: Clone Repository

```bash
git clone https://github.com/jwadow/mcp-excel.git
cd mcp-excel
```

*No Git? Click "Code" → "Download ZIP" at the top of this repository page, extract, and open terminal in that folder.*

### Step 2: Choose Installation Method

<details>
<summary><b>🎯 Option A: Poetry (Recommended)</b></summary>

*Poetry is a modern Python dependency manager (replaces pip+venv+requirements.txt). [Install it](https://python-poetry.org/docs/#installation): `pip install poetry` or `pipx install poetry`*

**Install dependencies:**
```bash
poetry install
```

**Configure your AI agent:**

Add this to your MCP settings (JSON config):
```json
{
  "mcpServers": {
    "excel": {
      "command": "poetry",
      "args": ["run", "python", "-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

**Important:** Replace `C:/path/to/mcp-excel` with actual path to the cloned repository.

</details>

<details>
<summary><b>📦 Option B: pip with virtual environment</b></summary>

**Install dependencies:**
```bash
# Windows
python -m venv venv
venv\Scripts\activate
pip install -e .

# Linux/Mac
python -m venv venv
source venv/bin/activate
pip install -e .
```

**Find Python path in venv:**
```bash
# Windows
where python

# Linux/Mac
which python
```

**Configure your AI agent:**

Add this to your MCP settings (JSON config):
```json
{
  "mcpServers": {
    "excel": {
      "command": "C:/path/to/mcp-excel/venv/Scripts/python.exe",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

**Important:**
- Replace `C:/path/to/mcp-excel/venv/Scripts/python.exe` with actual path from `where python` command
- On Linux/Mac use path from `which python` (e.g., `/path/to/mcp-excel/venv/bin/python`)

</details>

<details>
<summary><b>⚠️ Option C: System Python (Not Recommended)</b></summary>

**Install dependencies globally:**
```bash
pip install "mcp>=1.1.0" "pandas>=2.2.0" "pydantic>=2.10.0" "xlrd>=2.0.1" "openpyxl>=3.1.0" "psutil>=6.1.0" "python-dateutil>=2.9.0"
```

**Configure your AI agent:**
```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.main"],
      "cwd": "C:/path/to/mcp-excel"
    }
  }
}
```

⚠️ **Warning:** This pollutes your global Python environment. Use Poetry or venv instead.

</details>

### Step 3: Verify Installation

Restart your AI agent and test:
```
"Analyze the Excel file at C:/Users/YourName/Documents/test.xlsx"
```

If it works - you're done! If not, check:
- Path to repository is correct in `cwd`
- Python path is correct in `command` (for pip method)
- All dependencies are installed

### Supported AI Agents

Works with any MCP-compatible AI agent.

⚠️ **Important:** This is an MCP server. It runs automatically when your AI agent needs it. Do not run it manually in terminal.

## 💡 Usage

After configuration, restart your AI agent and ask it to analyze Excel files:

```
"Analyze the Excel file at C:/Users/YourName/Documents/sales.xls"
"Show me top 10 customers by revenue from sales.xlsx"
"Find duplicates in column 'Email' in contacts.xlsx"
"Calculate month-over-month growth from revenue.xls"
```

## 🛠️ Available Tools

<details>
<summary><b>📋 Complete Tool Reference (25 tools) - Click to expand</b></summary>

### 📊 File Inspection (5 tools)

#### `inspect_file`
Get file structure overview - sheets, dimensions, format.
**Use for:** Initial file exploration, sheet discovery, format validation
**Returns:** Sheet list, row/column counts, file metadata

#### `get_sheet_info`
Detailed sheet analysis with auto-header detection.
**Use for:** Understanding data structure, column types, sample preview
**Returns:** Column names/types, row count, sample data (3 rows), header detection info

#### `get_column_names`
Quick column enumeration without loading full data.
**Use for:** Schema validation, filter building, column availability checks
**Returns:** Column name list, column count

#### `get_data_profile`
Comprehensive column profiling - types, stats, nulls, top values.
**Use for:** Initial data exploration, quality assessment, distribution analysis
**Returns:** Per-column: type, null %, unique count, stats (numeric), top N values
**Efficiency:** Replaces 10+ separate calls (get_column_stats + get_value_counts + find_nulls)

#### `find_column`
Locate column across multiple sheets.
**Use for:** Multi-sheet navigation, data discovery, cross-sheet analysis
**Returns:** Sheet list with column locations, indices, row counts (case-insensitive)

---

### 📥 Data Retrieval (3 tools)

#### `get_unique_values`
Extract unique values from a column.
**Use for:** Data exploration, filter building, distinct value discovery, data quality checks
**Returns:** Unique value list, count, truncated flag (if limit exceeded)
**Default limit:** 100 values

#### `get_value_counts`
Frequency analysis - top N most common values.
**Use for:** Distribution analysis, identifying dominant categories, data imbalance detection
**Returns:** Value → count dictionary, total count, TSV output
**Default:** Top 10 values

#### `filter_and_get_rows`
Retrieve filtered rows with pagination.
**Use for:** Data extraction, sample inspection, detailed analysis, export
**Returns:** Filtered rows (list of dicts), total count, TSV output
**Pagination:** limit/offset support

---

### 🔍 Filtering & Counting (3 tools)

#### `filter_and_count`
Count rows matching conditions with 14 operators.
**Operators:** `==`, `!=`, `>`, `<`, `>=`, `<=`, `in`, `not_in`, `contains`, `startswith`, `endswith`, `regex`, `is_null`, `is_not_null`
**Logic:** Nested AND/OR groups, NOT operator, unlimited conditions
**Use for:** Classification, segmentation, data validation, category counting
**Returns:** Count + Excel formula (COUNTIFS), optional sample rows

#### `filter_and_count_batch`
Classify data into multiple categories in one call (6x faster).
**Use for:** Multi-category classification, market segmentation, quality control
**Returns:** Count + formula per category, TSV table for Excel
**Efficiency:** Loads file once, applies all filters, returns all results

#### `analyze_overlap`
Venn diagram analysis - intersections, unions, exclusive zones.
**Use for:** Overlap analysis, cross-sell opportunities, data consistency checks
**Returns:** Set counts, pairwise intersections (A ∩ B), union, Venn data (2-3 sets)
**Examples:** VIP AND active customers, product category overlaps, completed orders WITHOUT completion date

---

### 📈 Aggregation & Analysis (2 tools)

#### `aggregate`
Perform aggregation with optional filters (8 operations).
**Operations:** `sum`, `mean`, `median`, `min`, `max`, `std`, `var`, `count`
**Use for:** Totals, averages, statistical summaries, conditional aggregations, KPIs
**Returns:** Aggregated value + Excel formula (SUMIF, AVERAGEIF, etc.)
**Special:** Auto-converts text-stored numbers to numeric

#### `group_by`
Pivot table with multi-column grouping.
**Use for:** Category analysis, hierarchical grouping, sales by region/product
**Returns:** Grouped data with aggregated values, TSV output
**Supports:** Multiple grouping columns, all 8 aggregation operations

---

### 📊 Statistics (3 tools)

#### `get_column_stats`
Statistical summary - count, mean, median, std, quartiles.
**Use for:** Distribution analysis, data profiling, outlier detection prep
**Returns:** Full stats (min, max, mean, median, std, Q1, Q3), null count, TSV output

#### `correlate`
Correlation matrix between 2+ columns.
**Methods:** Pearson (linear), Spearman (rank-based), Kendall (rank-based)
**Use for:** Relationship analysis, variable dependency, feature selection
**Returns:** Correlation matrix (-1 to 1), TSV output

#### `detect_outliers`
Anomaly detection using IQR or Z-score.
**Methods:** IQR (robust), Z-score (assumes normal distribution)
**Use for:** Fraud detection, sensor errors, data quality, unusual value identification
**Returns:** Outlier rows with indices, count, method/threshold used

---

### ✅ Data Validation (2 tools)

#### `find_duplicates`
Detect duplicate rows by specified columns.
**Use for:** Data quality, deduplication planning, integrity checks
**Returns:** All duplicate rows (including first occurrence), count, indices
**Note:** Uses `duplicated(keep=False)` to mark all duplicates

#### `find_nulls`
Find null/empty values with detailed statistics.
**Use for:** Completeness checks, missing value analysis, data cleaning
**Returns:** Per-column: null count, percentage, indices (first 100)
**Note:** Placeholders (".", "-") are NOT null - use `==` or `in` operators

---

### 🔄 Multi-Sheet Operations (2 tools)

#### `search_across_sheets`
Search value across all sheets.
**Use for:** Cross-sheet search, value tracking, data location
**Returns:** Sheet list with match counts, total matches
**Supports:** Numeric and string values

#### `compare_sheets`
Diff between two sheets using key column.
**Use for:** Version comparison, change detection, reconciliation, audit trails
**Returns:** Rows with differences, status (only_in_sheet1/sheet2/different_values), side-by-side comparison

---

### 📅 Time Series (3 tools)

#### `calculate_period_change`
Period-over-period growth analysis.
**Periods:** month, quarter, year
**Use for:** Trend analysis, growth tracking, seasonal comparison, YoY analysis
**Returns:** Periods with values, absolute/percentage changes, Excel formula

#### `calculate_running_total`
Cumulative sum with optional grouping.
**Use for:** Cumulative analysis, progress tracking, balance calculations, cash flow
**Returns:** Rows with running totals, Excel formula (SUM($B$2:B2))
**Supports:** Grouping (running total resets per group)

#### `calculate_moving_average`
Smoothing with specified window size.
**Use for:** Trend detection, noise reduction, pattern identification
**Returns:** Rows with moving averages, Excel formula (AVERAGE(B1:B7))
**Examples:** 7-day moving average, 30-day stock price smoothing

---

### 🏆 Advanced Operations (2 tools)

#### `rank_rows`
Rank by column value with top-N filtering.
**Directions:** desc (highest first), asc (lowest first)
**Use for:** Leaderboards, top/bottom analysis, percentile ranking
**Returns:** Ranked rows with rank numbers, Excel formula (RANK)
**Supports:** Top-N filtering, ranking within groups

#### `calculate_expression`
Arithmetic expressions between columns.
**Operations:** `+`, `-`, `*`, `/`, parentheses
**Use for:** Derived metrics, financial calculations, ratio analysis, KPIs
**Returns:** Calculated values, Excel formula (e.g., =A2*B2)
**Examples:** Revenue = Price * Quantity, Margin = (Revenue - Cost) / Revenue

</details>

## 🗺️ Roadmap

### 📁 File Format Support

**Currently Supported:**
- ✅ **XLS** - Excel 97-2003 (read-only)
- ✅ **XLSX** - Excel 2007+ (read-only)

**Planned:**
- 🔜 **XLSM** - Excel with macros support
- 🔜 **CSV** - Comma-separated values
- 🔜 **TSV** - Tab-separated values
- 🔜 **ODS** - OpenDocument Spreadsheet
- 🔜 **Parquet** - Columnar storage format

### 🚀 Features

- **Write operations** - Modify spreadsheets files (create calculated columns, update values)
- **SSE transport mode** - Server-Sent Events for remote access
- **Advanced formula generation** - More complex Excel formulas with nested functions
- **Data export** - Export filtered/aggregated results to new files

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

## 🤝 Contributing

Contributions are welcome! Please ensure:

1. All dependencies are AGPL-compatible
2. Code follows the existing style
3. Tests are included for new features
4. Documentation is updated

For issues, questions, or contributions, please open an issue on GitHub.

---

## 💬 Need Help?

Got questions? Found a bug? Have a feature idea? We're here to help!

**👉 [Open an Issue on GitHub](https://github.com/jwadow/mcp-excel/issues/new)**

Whether you're stuck with installation, found something broken, or just want to suggest an improvement — GitHub Issues is the place. Don't worry if you're new to GitHub, just click the link above and describe your situation. We'll figure it out together.

---

<div align="center">

**[⬆ Back to Top](#-excel-mcp-server)**

</div>
