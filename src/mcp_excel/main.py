# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""MCP Excel Server - Main entry point."""

import asyncio
import logging
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from .core.file_loader import FileLoader
from .models.requests import (
    AggregateRequest,
    CalculateExpressionRequest,
    CalculateMovingAverageRequest,
    CalculatePeriodChangeRequest,
    CalculateRunningTotalRequest,
    CompareSheetsRequest,
    CorrelateRequest,
    DetectOutliersRequest,
    FilterAndCountRequest,
    FilterAndGetRowsRequest,
    FindColumnRequest,
    FindDuplicatesRequest,
    FindNullsRequest,
    GetColumnNamesRequest,
    GetColumnStatsRequest,
    GetDataProfileRequest,
    GetSheetInfoRequest,
    GetUniqueValuesRequest,
    GetValueCountsRequest,
    GroupByRequest,
    InspectFileRequest,
    RankRowsRequest,
    SearchAcrossSheetsRequest,
)
from .operations.advanced import AdvancedOperations
from .operations.data_operations import DataOperations
from .operations.inspection import InspectionOperations
from .operations.statistics import StatisticsOperations
from .operations.timeseries import TimeSeriesOperations
from .operations.validation import ValidationOperations

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("mcp_excel")


class MCPExcelServer:
    """MCP server for Excel operations."""

    def __init__(self) -> None:
        """Initialize MCP Excel server."""
        self.server = Server("mcp-excel")
        self.file_loader = FileLoader()
        self.inspection_ops = InspectionOperations(self.file_loader)
        self.data_ops = DataOperations(self.file_loader)
        self.stats_ops = StatisticsOperations(self.file_loader)
        self.validation_ops = ValidationOperations(self.file_loader)
        self.timeseries_ops = TimeSeriesOperations(self.file_loader)
        self.advanced_ops = AdvancedOperations(self.file_loader)

        # Register handlers
        self._register_handlers()

    def _register_handlers(self) -> None:
        """Register MCP handlers."""

        @self.server.list_tools()
        async def list_tools() -> list[Tool]:
            """List available tools."""
            return [
                Tool(
                    name="inspect_file",
                    description="Inspect Excel file structure and get basic information about all sheets",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            }
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="get_sheet_info",
                    description="Get detailed information about a specific sheet including column names, types, and sample data",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet to inspect",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name"],
                    },
                ),
                Tool(
                    name="get_column_names",
                    description="Get list of column names from a sheet",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name"],
                    },
                ),
                Tool(
                    name="get_data_profile",
                    description="Get comprehensive data profile for columns including type, statistics, null counts, and top values. Combines multiple operations (get_column_stats, get_value_counts, find_nulls) into a single efficient call. Ideal for initial data exploration and understanding column characteristics.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to profile (optional, profiles all columns if not specified)",
                            },
                            "top_n": {
                                "type": "integer",
                                "description": "Number of top values to return per column (default: 5)",
                                "default": 5,
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name"],
                    },
                ),
                Tool(
                    name="find_column",
                    description="Find a column across all sheets or in the first sheet. Returns list of sheets where the column was found with metadata.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "column_name": {
                                "type": "string",
                                "description": "Column name to search for (case-insensitive)",
                            },
                            "search_all_sheets": {
                                "type": "boolean",
                                "description": "Search in all sheets (true) or just first sheet (false). Default: true",
                                "default": True,
                            },
                        },
                        "required": ["file_path", "column_name"],
                    },
                ),
                Tool(
                    name="get_unique_values",
                    description="Get unique values from a column (useful for building filters)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "column": {
                                "type": "string",
                                "description": "Column name to get unique values from",
                            },
                            "limit": {
                                "type": "integer",
                                "description": "Maximum number of unique values to return (default: 100)",
                                "default": 100,
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "column"],
                    },
                ),
                Tool(
                    name="get_value_counts",
                    description="Get frequency counts for values in a column (top N most common values)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "column": {
                                "type": "string",
                                "description": "Column name to analyze",
                            },
                            "top_n": {
                                "type": "integer",
                                "description": "Number of top values to return (default: 10)",
                                "default": 10,
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "column"],
                    },
                ),
                Tool(
                    name="filter_and_count",
                    description="Count rows matching filter conditions",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "filters": {
                                "type": "array",
                                "description": "List of filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "filters"],
                    },
                ),
                Tool(
                    name="filter_and_get_rows",
                    description="Get rows matching filter conditions with pagination support",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "filters": {
                                "type": "array",
                                "description": "List of filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to return (optional, returns all if not specified)",
                            },
                            "limit": {
                                "type": "integer",
                                "description": "Maximum number of rows to return (default: 50)",
                                "default": 50,
                            },
                            "offset": {
                                "type": "integer",
                                "description": "Number of rows to skip (default: 0)",
                                "default": 0,
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "filters"],
                    },
                ),
                Tool(
                    name="aggregate",
                    description="Perform aggregation (sum, mean, count, etc.) on a column with optional filters",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "operation": {
                                "type": "string",
                                "enum": ["sum", "mean", "median", "min", "max", "std", "var", "count"],
                                "description": "Aggregation operation to perform",
                            },
                            "target_column": {
                                "type": "string",
                                "description": "Column to aggregate",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "operation", "target_column"],
                    },
                ),
                Tool(
                    name="group_by",
                    description="Group data by columns and perform aggregation (like Excel Pivot Table)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "group_columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to group by (can be multiple)",
                            },
                            "agg_column": {
                                "type": "string",
                                "description": "Column to aggregate",
                            },
                            "agg_operation": {
                                "type": "string",
                                "enum": ["sum", "mean", "median", "min", "max", "std", "var", "count"],
                                "description": "Aggregation operation",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "group_columns", "agg_column", "agg_operation"],
                    },
                ),
                Tool(
                    name="get_column_stats",
                    description="Get statistical summary of a column (count, mean, median, std, min, max, quartiles)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "column": {
                                "type": "string",
                                "description": "Column name to analyze",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "column"],
                    },
                ),
                Tool(
                    name="correlate",
                    description="Calculate correlation matrix between multiple columns (supports 2+ columns)",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to correlate (minimum 2 columns)",
                            },
                            "method": {
                                "type": "string",
                                "enum": ["pearson", "spearman", "kendall"],
                                "description": "Correlation method (default: pearson)",
                                "default": "pearson",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "columns"],
                    },
                ),
                Tool(
                    name="detect_outliers",
                    description="Detect outliers in a column using IQR or Z-score method",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "column": {
                                "type": "string",
                                "description": "Column name to analyze",
                            },
                            "method": {
                                "type": "string",
                                "enum": ["iqr", "zscore"],
                                "description": "Outlier detection method (default: iqr)",
                                "default": "iqr",
                            },
                            "threshold": {
                                "type": "number",
                                "description": "Threshold for outlier detection (IQR multiplier or Z-score, default: 1.5)",
                                "default": 1.5,
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "column"],
                    },
                ),
                Tool(
                    name="search_across_sheets",
                    description="Search for a value across all sheets in the file. Returns list of sheets with match counts.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "column_name": {
                                "type": "string",
                                "description": "Column name to search in (case-insensitive)",
                            },
                            "value": {
                                "description": "Value to search for (supports numbers and strings)",
                            },
                        },
                        "required": ["file_path", "column_name", "value"],
                    },
                ),
                Tool(
                    name="compare_sheets",
                    description="Compare data between two sheets using a key column. Returns rows with differences.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet1": {
                                "type": "string",
                                "description": "First sheet name to compare",
                            },
                            "sheet2": {
                                "type": "string",
                                "description": "Second sheet name to compare",
                            },
                            "key_column": {
                                "type": "string",
                                "description": "Column to use as key for matching rows between sheets",
                            },
                            "compare_columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to compare for differences",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet1", "sheet2", "key_column", "compare_columns"],
                    },
                ),
                Tool(
                    name="find_duplicates",
                    description="Find duplicate rows based on specified columns. Returns all duplicate rows including first occurrence.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to check for duplicates (checks combination of these columns)",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "columns"],
                    },
                ),
                Tool(
                    name="find_nulls",
                    description="Find null/empty values in specified columns. Returns statistics and row indices for each column.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to check for null values",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "columns"],
                    },
                ),
                Tool(
                    name="calculate_period_change",
                    description="Calculate period-over-period change (month/quarter/year). Returns periods with values and percentage changes.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "date_column": {
                                "type": "string",
                                "description": "Column containing dates",
                            },
                            "value_column": {
                                "type": "string",
                                "description": "Column containing values to analyze",
                            },
                            "period_type": {
                                "type": "string",
                                "enum": ["month", "quarter", "year"],
                                "description": "Period type for grouping",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "date_column", "value_column", "period_type"],
                    },
                ),
                Tool(
                    name="calculate_running_total",
                    description="Calculate running total (cumulative sum) ordered by a column. Supports grouping.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "order_column": {
                                "type": "string",
                                "description": "Column to order by (typically date)",
                            },
                            "value_column": {
                                "type": "string",
                                "description": "Column containing values to sum",
                            },
                            "group_by_columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Optional columns to group by (running total within groups)",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "order_column", "value_column"],
                    },
                ),
                Tool(
                    name="calculate_moving_average",
                    description="Calculate moving average with specified window size. Useful for trend analysis.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "order_column": {
                                "type": "string",
                                "description": "Column to order by (typically date)",
                            },
                            "value_column": {
                                "type": "string",
                                "description": "Column containing values to average",
                            },
                            "window_size": {
                                "type": "integer",
                                "description": "Number of periods for moving average window",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "order_column", "value_column", "window_size"],
                    },
                ),
                Tool(
                    name="rank_rows",
                    description="Rank rows by column value (ascending or descending). Supports top-N and grouping.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "rank_column": {
                                "type": "string",
                                "description": "Column to rank by",
                            },
                            "direction": {
                                "type": "string",
                                "enum": ["asc", "desc"],
                                "description": "Ranking direction (desc = highest first, default: desc)",
                                "default": "desc",
                            },
                            "top_n": {
                                "type": "integer",
                                "description": "Return only top N rows (optional, returns all if not specified)",
                            },
                            "group_by_columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Optional columns to group by (ranking within groups)",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "rank_column"],
                    },
                ),
                Tool(
                    name="calculate_expression",
                    description="Calculate expression between columns (e.g., 'Price * Quantity'). Supports arithmetic operations.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet",
                            },
                            "expression": {
                                "type": "string",
                                "description": "Expression to calculate (e.g., 'Price * Quantity', 'Revenue / Cost')",
                            },
                            "output_column_name": {
                                "type": "string",
                                "description": "Name for the calculated column",
                            },
                            "filters": {
                                "type": "array",
                                "description": "Optional filter conditions",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "column": {"type": "string"},
                                        "operator": {
                                            "type": "string",
                                            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"]
                                        },
                                        "value": {"description": "Value for single-value operators"},
                                        "values": {
                                            "type": "array",
                                            "description": "Values for 'in' and 'not_in' operators"
                                        }
                                    },
                                    "required": ["column", "operator"]
                                }
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": "Logic operator for combining multiple filters. 'AND' means all filters must match (intersection). 'OR' means at least one filter must match (union). Default: 'AND'. Note: Complex nested logic like '(A AND B) OR C' is not supported in a single call - use multiple calls and combine results in your analysis.",
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "required": ["file_path", "sheet_name", "expression", "output_column_name"],
                    },
                ),
            ]

        @self.server.call_tool()
        async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
            """Handle tool calls."""
            try:
                logger.info(f"Tool called: {name} with arguments: {arguments}")

                if name == "inspect_file":
                    request = InspectFileRequest(**arguments)
                    response = self.inspection_ops.inspect_file(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "get_sheet_info":
                    request = GetSheetInfoRequest(**arguments)
                    response = self.inspection_ops.get_sheet_info(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "get_column_names":
                    request = GetColumnNamesRequest(**arguments)
                    response = self.inspection_ops.get_column_names(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "get_data_profile":
                    request = GetDataProfileRequest(**arguments)
                    response = self.inspection_ops.get_data_profile(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "find_column":
                    request = FindColumnRequest(**arguments)
                    response = self.inspection_ops.find_column(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "get_unique_values":
                    request = GetUniqueValuesRequest(**arguments)
                    response = self.data_ops.get_unique_values(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "get_value_counts":
                    request = GetValueCountsRequest(**arguments)
                    response = self.data_ops.get_value_counts(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "filter_and_count":
                    request = FilterAndCountRequest(**arguments)
                    response = self.data_ops.filter_and_count(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "filter_and_get_rows":
                    request = FilterAndGetRowsRequest(**arguments)
                    response = self.data_ops.filter_and_get_rows(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "aggregate":
                    request = AggregateRequest(**arguments)
                    response = self.data_ops.aggregate(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "group_by":
                    request = GroupByRequest(**arguments)
                    response = self.data_ops.group_by(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "get_column_stats":
                    request = GetColumnStatsRequest(**arguments)
                    response = self.stats_ops.get_column_stats(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "correlate":
                    request = CorrelateRequest(**arguments)
                    response = self.stats_ops.correlate(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "detect_outliers":
                    request = DetectOutliersRequest(**arguments)
                    response = self.stats_ops.detect_outliers(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "search_across_sheets":
                    request = SearchAcrossSheetsRequest(**arguments)
                    response = self.inspection_ops.search_across_sheets(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "compare_sheets":
                    request = CompareSheetsRequest(**arguments)
                    response = self.inspection_ops.compare_sheets(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "find_duplicates":
                    request = FindDuplicatesRequest(**arguments)
                    response = self.validation_ops.find_duplicates(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "find_nulls":
                    request = FindNullsRequest(**arguments)
                    response = self.validation_ops.find_nulls(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "calculate_period_change":
                    request = CalculatePeriodChangeRequest(**arguments)
                    response = self.timeseries_ops.calculate_period_change(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "calculate_running_total":
                    request = CalculateRunningTotalRequest(**arguments)
                    response = self.timeseries_ops.calculate_running_total(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "calculate_moving_average":
                    request = CalculateMovingAverageRequest(**arguments)
                    response = self.timeseries_ops.calculate_moving_average(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "rank_rows":
                    request = RankRowsRequest(**arguments)
                    response = self.advanced_ops.rank_rows(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "calculate_expression":
                    request = CalculateExpressionRequest(**arguments)
                    response = self.advanced_ops.calculate_expression(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                else:
                    raise ValueError(f"Unknown tool: {name}")

            except Exception as e:
                logger.error(f"Error executing tool {name}: {e}", exc_info=True)
                error_response = {
                    "error": type(e).__name__,
                    "message": str(e),
                    "recoverable": True,
                }
                return [TextContent(type="text", text=str(error_response))]

    async def run(self) -> None:
        """Run the MCP server."""
        logger.info("Starting MCP Excel Server...")
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options(),
            )


def main() -> None:
    """Main entry point."""
    server = MCPExcelServer()
    asyncio.run(server.run())


if __name__ == "__main__":
    main()
