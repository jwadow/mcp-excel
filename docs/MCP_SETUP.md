# MCP Server Configuration Guide

## For Claude Desktop

Add this to your Claude Desktop config file:

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`
**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": [
        "-m",
        "mcp_excel.main"
      ],
      "cwd": "D:/Projects/Python/mcp-excel",
      "env": {
        "PYTHONPATH": "D:/Projects/Python/mcp-excel/src"
      }
    }
  }
}
```

**Important:** Replace `D:/Projects/Python/mcp-excel` with your actual project path.

## For Other MCP Clients

If using a different MCP client, start the server manually:

```bash
cd D:/Projects/Python/mcp-excel
python -m mcp_excel.main
```

The server will communicate via STDIO (stdin/stdout).

## Currently Available Tools

### ✅ Implemented (Ready to Use)

1. **inspect_file** - Get file structure
   - File format and size
   - List of all sheets
   - Row/column counts

2. **get_sheet_info** - Detailed sheet analysis
   - Column names and types
   - Auto-detected headers
   - Sample data (first 3 rows)
   - Data quality metrics

3. **get_column_names** - Quick column list
   - Just the column names
   - Fast operation

### ⏳ Not Yet Implemented

These are in the architecture but not coded yet:
- Filtering and counting
- Aggregations (sum, mean, etc.)
- Correlation analysis
- Statistical analysis
- Multi-sheet operations
- Data validation

## Example Tasks for Agent

### What You CAN Ask Now:

1. **File Inspection:**
   ```
   "Analyze the structure of my Excel file at D:/data/report.xlsx"
   ```

2. **Sheet Analysis:**
   ```
   "Show me the columns and data types in the 'Sales' sheet"
   ```

3. **Header Detection:**
   ```
   "This file has messy headers - can you find where the actual data starts?"
   ```

4. **Multi-file Comparison:**
   ```
   "Compare the structure of these two Excel files and tell me if they're compatible"
   ```

### What You CANNOT Ask Yet:

❌ "Calculate correlation between Price and Quantity"
❌ "Filter rows where Status = 'Active'"
❌ "Sum all values in the Amount column"
❌ "Find duplicates in Customer column"

These require additional operations that aren't implemented yet.

## Testing the Connection

After adding the config:

1. Restart Claude Desktop
2. Start a new conversation
3. Ask: "Can you inspect this Excel file: [your file path]"
4. The agent should use the `inspect_file` tool

## Troubleshooting

### Server Not Starting

Check logs in Claude Desktop:
- Windows: `%APPDATA%\Claude\logs\`
- macOS: `~/Library/Logs/Claude/`

### Common Issues

1. **Python not found:**
   - Use full path: `"command": "C:/Python310/python.exe"`

2. **Module not found:**
   - Check `PYTHONPATH` in config
   - Or install: `pip install -e .` in project directory

3. **Permission denied:**
   - Make sure Python has access to the file paths

## Next Steps

To add more functionality (filtering, aggregation, etc.), we need to:
1. Implement the operations in `operations/` directory
2. Register them as tools in `main.py`
3. Restart the server

The architecture is ready - we just need to code the remaining operations.
