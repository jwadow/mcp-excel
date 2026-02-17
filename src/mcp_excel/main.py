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
    AnalyzeOverlapRequest,
    CalculateExpressionRequest,
    CalculateMovingAverageRequest,
    CalculatePeriodChangeRequest,
    CalculateRunningTotalRequest,
    CompareSheetsRequest,
    CorrelateRequest,
    DetectOutliersRequest,
    FilterAndCountBatchRequest,
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

# ============================================================================
# SHARED DESCRIPTIONS (DRY principle - reusable across multiple tools)
# ============================================================================

OPERATOR_DESCRIPTION = """Comparison operator for filtering data.

OPERATORS: ==, !=, >, <, >=, <= (comparison) | in, not_in (set) | contains, startswith, endswith, regex (string) | is_null, is_not_null (null check)

NEGATION: Use 'negate: true' to invert any condition (NOT operator). Example: {"column": "Status", "operator": "==", "value": "Active", "negate": true} means Status != "Active"

CRITICAL - NULL vs PLACEHOLDER DISTINCTION:
- is_null: detects ONLY real empty cells (NaN/None in pandas)
- is_not_null: matches everything that's not empty, INCLUDING placeholders
- Placeholders (".", "-", spaces, etc.) are REGULAR STRINGS, NOT null
- To filter placeholders: use {"operator": "==", "value": "."} or {"operator": "in", "values": [".", "-"]}
- To exclude placeholders: use {"operator": "not_in", "values": [".", "-"]}

USAGE PATTERNS:
- Single value: {"column": "Age", "operator": ">", "value": 30}
- Multiple values: {"column": "Status", "operator": "in", "values": ["Active", "Pending"]}
- Null check: {"column": "Email", "operator": "is_null"}"""

LOGIC_DESCRIPTION = """Logic operator for combining multiple filters.

AND: ALL conditions must be true (intersection)
OR: AT LEAST ONE condition must be true (union)

NESTED GROUPS: You can create complex logical expressions using nested groups:
- (A AND B) OR C: Use a group with AND logic for A,B, then combine with C using OR
- A AND (B OR C): Use a group with OR logic for B,C, then combine with A using AND
- ((A OR B) AND C) OR D: Nest groups within groups for complex logic

NUMERICAL EXAMPLE:
Dataset: 100 rows | Filter A: 30 rows | Filter B: 20 rows | Overlap: 5 rows
→ AND returns 5 rows (intersection) | OR returns 45 rows (30+20-5, union)

COMMON PATTERNS:
- Classification: Use multiple filter_and_count calls with different filters
- Union of conditions: Use OR logic (e.g., Category A OR Category B)
- Intersection of conditions: Use AND logic (e.g., Age > 30 AND City = "Moscow")
- Exclude records: Use != or not_in operators
- Complex logic: Use nested groups (e.g., (Status=Active AND Amount>1000) OR (Status=VIP))

Default: AND"""

# Define FilterCondition schema (atomic filter)
FILTER_CONDITION_SCHEMA = {
    "type": "object",
    "properties": {
        "column": {"type": "string", "description": "Column name to filter on"},
        "operator": {
            "type": "string",
            "enum": ["==", "!=", ">", "<", ">=", "<=", "in", "not_in", "contains", "startswith", "endswith", "regex", "is_null", "is_not_null"],
            "description": OPERATOR_DESCRIPTION
        },
        "value": {"description": "Value for single-value operators (==, !=, >, <, >=, <=, contains, startswith, endswith, regex, is_null, is_not_null)"},
        "values": {
            "type": "array",
            "description": "List of values for set operators (in, not_in). Example: ['Active', 'Pending']"
        },
        "negate": {
            "type": "boolean",
            "default": False,
            "description": "Negate the condition (NOT operator). Inverts the result of the filter."
        }
    },
    "required": ["column", "operator"]
}

# Define FilterGroup schema (nested group of filters)
# Note: This is defined as a separate variable to enable recursive reference
FILTER_GROUP_SCHEMA = {
    "type": "object",
    "properties": {
        "filters": {
            "type": "array",
            "description": "List of filter conditions or nested groups",
            "items": {
                "oneOf": [
                    FILTER_CONDITION_SCHEMA,
                    {"$ref": "#/definitions/FilterGroup"}
                ]
            }
        },
        "logic": {
            "type": "string",
            "enum": ["AND", "OR"],
            "description": LOGIC_DESCRIPTION,
            "default": "AND"
        },
        "negate": {
            "type": "boolean",
            "default": False,
            "description": "Negate the entire group (NOT operator). Inverts the result of all filters in this group."
        }
    },
    "required": ["filters"]
}

# Filter property schema - for use in inputSchema properties (without definitions)
# This is the structure for the "filters" property itself
FILTER_PROPERTY_SCHEMA = {
    "type": "array",
    "description": "List of filter conditions or nested groups. Supports complex logical expressions like (A AND B) OR C.",
    "items": {
        "oneOf": [
            FILTER_CONDITION_SCHEMA,
            {"$ref": "#/definitions/FilterGroup"}
        ]
    }
}

# Filter definitions - must be placed at root level of inputSchema
# This enables recursive $ref resolution for nested FilterGroup structures
FILTER_DEFINITIONS = {
    "FilterGroup": FILTER_GROUP_SCHEMA
}

# Legacy: Keep FILTER_SCHEMA for backward compatibility (if needed elsewhere)
FILTER_SCHEMA = {
    "type": "array",
    "description": "List of filter conditions or nested groups. Supports complex logical expressions like (A AND B) OR C.",
    "items": {
        "oneOf": [
            FILTER_CONDITION_SCHEMA,
            FILTER_GROUP_SCHEMA
        ]
    },
    "definitions": {
        "FilterGroup": FILTER_GROUP_SCHEMA
    }
}


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
                    description="Inspect Excel file structure and get basic information about all sheets. Returns sheet names, dimensions, and file metadata. Use for: file overview, sheet discovery, file validation, structure understanding, initial file assessment. EXAMPLES: Get all sheet names in workbook, Check file size and format, Verify file structure before processing, Discover available data sheets.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file (.xls or .xlsx)",
                            }
                        },
                        "required": ["file_path"],
                    },
                ),
                Tool(
                    name="get_sheet_info",
                    description="Get detailed information about a specific sheet including column names, types, sample data, and row count. Returns comprehensive sheet metadata. Use for: sheet exploration, data structure understanding, type inference, header detection, schema discovery. EXAMPLES: Understand data types before filtering, Check column names for filter building, Verify data structure matches expectations, Get sample rows for validation.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet to inspect (case-sensitive)",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided). Use if headers are not in row 1.",
                            },
                        },
                        "required": ["file_path", "sheet_name"],
                    },
                ),
                Tool(
                    name="get_column_names",
                    description="Get list of column names from a sheet. Quick operation for schema discovery without loading full data. Use for: column enumeration, schema validation, filter building, data structure verification, column availability checks. EXAMPLES: List all available columns, Verify column exists before filtering, Get column names for dynamic filter building, Check schema consistency.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet (case-sensitive)",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided). Use if headers are not in row 1.",
                            },
                        },
                        "required": ["file_path", "sheet_name"],
                    },
                ),
                Tool(
                    name="get_data_profile",
                    description="Get comprehensive data profile for columns including type, statistics, null counts, and top values. Combines multiple operations (get_column_stats, get_value_counts, find_nulls) into a single efficient call. Use for: initial data exploration, column profiling, data quality overview, schema understanding, one-shot data assessment. EXAMPLES: Profile all columns in dataset, Understand data quality issues, Get distribution of values, Assess completeness and types.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet (case-sensitive)",
                            },
                            "columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to profile (optional, profiles all columns if not specified). Use to focus on specific columns.",
                            },
                            "top_n": {
                                "type": "integer",
                                "description": "Number of top values to return per column (default: 5). Shows most frequent values.",
                                "default": 5,
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided). Use if headers are not in row 1.",
                            },
                        },
                        "required": ["file_path", "sheet_name"],
                    },
                ),
                Tool(
                    name="find_column",
                    description="Find a column across all sheets or in the first sheet. Returns list of sheets where the column was found with metadata. Use for: multi-sheet navigation, column discovery, data structure understanding, cross-sheet analysis, locating data. EXAMPLES: Find which sheets contain 'CustomerID', Locate 'Revenue' column across monthly sheets, Discover data distribution across sheets.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "column_name": {
                                "type": "string",
                                "description": "Column name to search for (case-insensitive). Searches by exact name match.",
                            },
                            "search_all_sheets": {
                                "type": "boolean",
                                "description": "Search in all sheets (true, default) or just first sheet (false). Use false for faster search in single-sheet files.",
                                "default": True,
                            },
                        },
                        "required": ["file_path", "column_name"],
                    },
                ),
                Tool(
                    name="get_unique_values",
                    description="Get unique values from a column. Essential for understanding data distribution and building accurate filters. Use for: data exploration, filter validation, distinct value discovery, data quality checks, filter building. EXAMPLES: Get all customer names for filter, Find unique product categories, Discover all status values, Validate data consistency.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet (case-sensitive)",
                            },
                            "column": {
                                "type": "string",
                                "description": "Column name to get unique values from",
                            },
                            "limit": {
                                "type": "integer",
                                "description": "Maximum number of unique values to return (default: 100). Use to limit output for high-cardinality columns.",
                                "default": 100,
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided). Use if headers are not in row 1.",
                            },
                        },
                        "required": ["file_path", "sheet_name", "column"],
                    },
                ),
                Tool(
                    name="get_value_counts",
                    description="Get frequency counts for values in a column (top N most common values). Shows distribution and prevalence of values. Use for: frequency analysis, data distribution, identifying dominant categories, quality assessment, understanding data patterns. EXAMPLES: Top 10 customers by order count, Most common product categories, Frequency of status values, Identifying data imbalance.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet_name": {
                                "type": "string",
                                "description": "Name of the sheet (case-sensitive)",
                            },
                            "column": {
                                "type": "string",
                                "description": "Column name to analyze",
                            },
                            "top_n": {
                                "type": "integer",
                                "description": "Number of top values to return (default: 10). Shows most frequent values with their counts.",
                                "default": 10,
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided). Use if headers are not in row 1.",
                            },
                        },
                        "required": ["file_path", "sheet_name", "column"],
                    },
                ),
                Tool(
                    name="filter_and_count",
                    description="Count rows matching filter conditions. Returns count and Excel formula for dynamic updates. Use for: classification, segmentation, data validation, counting specific categories, multi-condition analysis. EXAMPLES: Count active users (filter: Status=Active), Count high-value orders (filter: Amount>1000), Count items in multiple categories (filter: Category in [A,B,C]).",
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
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                            "sample_rows": {
                                "type": "integer",
                                "description": "Number of sample rows to return (optional). Shows examples of rows matching filters.",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "filters"],
                    },
                ),
                Tool(
                    name="filter_and_count_batch",
                    description="Count rows for multiple filter sets in a single call. Optimized for classification, segmentation, and multi-category analysis. Much faster than multiple filter_and_count calls (loads file once, applies all filters). Use for: data classification into categories, market segmentation, quality control checks, multi-condition validation, inventory classification. EXAMPLES: Classify orders (Pending, Processing, Shipped, Delivered), Segment customers (VIP, Regular, Inactive), Quality checks (Pass, Fail, Rework).",
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
                            "filter_sets": {
                                "type": "array",
                                "description": "List of filter sets to evaluate independently. Each set is processed separately and returns its own count. Each filter set can use AND/OR logic internally.",
                                "minItems": 1,
                                "maxItems": 50,
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "label": {
                                            "type": "string",
                                            "description": "Optional label for this filter set (e.g., 'Category A', 'Active items'). If not provided, will be labeled as 'Set 1', 'Set 2', etc."
                                        },
                                        "filters": FILTER_PROPERTY_SCHEMA,
                                        "logic": {
                                            "type": "string",
                                            "enum": ["AND", "OR"],
                                            "description": LOGIC_DESCRIPTION,
                                            "default": "AND"
                                        },
                                        "sample_rows": {
                                            "type": "integer",
                                            "description": "Number of sample rows to return for this filter set (optional).",
                                        }
                                    },
                                    "required": ["filters"]
                                }
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "filter_sets"],
                    },
                ),
                Tool(
                    name="analyze_overlap",
                    description="Analyze overlap between multiple filter sets (Venn diagram analysis). Returns intersection counts, union, and exclusive zones. Optimized for classification, segmentation, and cross-sell analysis. Much faster than multiple separate calls (loads file once, applies all filters). Use for: overlap analysis, Venn diagrams, classification into categories, market segmentation, cross-sell analysis, data consistency checks. EXAMPLES: Find customers who are both VIP AND active, Analyze overlap between product categories, Segment users by multiple criteria, Check data consistency (orders completed BUT no completion date).",
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
                            "filter_sets": {
                                "type": "array",
                                "description": "List of filter sets to analyze overlap between (2-10 sets). Each set is evaluated independently, then overlaps are calculated.",
                                "minItems": 2,
                                "maxItems": 10,
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "label": {
                                            "type": "string",
                                            "description": "Optional label for this filter set (e.g., 'VIP customers', 'Active users'). If not provided, will be labeled as 'Set 1', 'Set 2', etc."
                                        },
                                        "filters": FILTER_PROPERTY_SCHEMA,
                                        "logic": {
                                            "type": "string",
                                            "enum": ["AND", "OR"],
                                            "description": LOGIC_DESCRIPTION,
                                            "default": "AND"
                                        }
                                    },
                                    "required": ["filters"]
                                }
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "filter_sets"],
                    },
                ),
                Tool(
                    name="filter_and_get_rows",
                    description="Get rows matching filter conditions with pagination support. Returns filtered data for analysis, export, or further processing. Use for: data extraction, sample inspection, detailed analysis, data export, verification of filter results. EXAMPLES: Get all orders from customer X, Extract rows with errors for review, Get sample of high-value transactions, Export filtered data to another system.",
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
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to return (optional, returns all if not specified). Use to reduce output size and focus on relevant fields.",
                            },
                            "limit": {
                                "type": "integer",
                                "description": "Maximum number of rows to return (default: 50). Use for pagination or limiting large result sets.",
                                "default": 50,
                            },
                            "offset": {
                                "type": "integer",
                                "description": "Number of rows to skip (default: 0). Use for pagination (e.g., offset=50 with limit=50 gets rows 51-100).",
                                "default": 0,
                            },
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "filters"],
                    },
                ),
                Tool(
                    name="aggregate",
                    description="Perform aggregation (sum, mean, count, etc.) on a column with optional filters. Returns aggregated value and Excel formula for dynamic updates. Use for: totals, averages, min/max values, statistical summaries, conditional aggregations, KPI calculations. EXAMPLES: Total revenue for Q4, Average order value for VIP customers, Maximum temperature in July, Count of defective items, Standard deviation of prices.",
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
                                "description": "Aggregation operation: sum (total), mean (average), median (middle value), min (minimum), max (maximum), std (standard deviation), var (variance), count (row count)",
                            },
                            "target_column": {
                                "type": "string",
                                "description": "Column to aggregate (must be numeric for sum/mean/median/min/max/std/var)",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                            "sample_rows": {
                                "type": "integer",
                                "description": "Number of sample rows to return (optional). Shows examples of rows used in aggregation.",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "operation", "target_column"],
                    },
                ),
                Tool(
                    name="group_by",
                    description="Group data by columns and perform aggregation (like Excel Pivot Table). Returns grouped results with aggregated values. Supports multiple grouping columns for hierarchical analysis. Use for: pivot tables, data summarization, category analysis, hierarchical grouping, sales by region/product, performance by team/month. EXAMPLES: Revenue by product category, Average salary by department and job level, Count of orders by customer and month.",
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
                                "description": "Columns to group by (can be multiple for hierarchical grouping, e.g., [Region, Product] groups by region first, then product within each region)",
                            },
                            "agg_column": {
                                "type": "string",
                                "description": "Column to aggregate (must be numeric for sum/mean/median/min/max/std/var)",
                            },
                            "agg_operation": {
                                "type": "string",
                                "enum": ["sum", "mean", "median", "min", "max", "std", "var", "count"],
                                "description": "Aggregation operation: sum (total), mean (average), median (middle value), min (minimum), max (maximum), std (standard deviation), var (variance), count (row count)",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "group_columns", "agg_column", "agg_operation"],
                    },
                ),
                Tool(
                    name="get_column_stats",
                    description="Get statistical summary of a column (count, mean, median, std, min, max, quartiles, null count). Returns comprehensive statistics for numeric columns. Use for: statistical analysis, data profiling, distribution understanding, outlier detection preparation, data quality assessment. EXAMPLES: Analyze salary distribution, Check price range and variance, Understand age demographics, Assess data completeness (null count).",
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
                                "description": "Column name to analyze (should be numeric for meaningful statistics)",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                            "sample_rows": {
                                "type": "integer",
                                "description": "Number of sample rows to return (optional). Shows examples of data being analyzed.",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "column"],
                    },
                ),
                Tool(
                    name="correlate",
                    description="Calculate correlation matrix between multiple columns (supports 2+ columns). Supports pearson (linear), spearman (rank-based), kendall (rank-based) methods. Returns correlation coefficients (-1 to 1). Use for: relationship analysis, variable dependency, feature selection, multivariate analysis, identifying correlated metrics. EXAMPLES: Correlation between price and sales volume, Relationship between temperature and energy consumption, Dependencies between financial metrics.",
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
                                "description": "Columns to correlate (minimum 2 columns, should be numeric)",
                            },
                            "method": {
                                "type": "string",
                                "enum": ["pearson", "spearman", "kendall"],
                                "description": "Correlation method: pearson (linear relationships), spearman (rank-based, handles outliers), kendall (rank-based, more robust). Default: pearson",
                                "default": "pearson",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "columns"],
                    },
                ),
                Tool(
                    name="detect_outliers",
                    description="Detect outliers in a column using IQR or Z-score method. Returns outlier rows with indices and method used. Use for: anomaly detection, data quality, statistical analysis, unusual value identification, fraud detection, sensor error detection. EXAMPLES: Find unusually high transactions, Detect equipment failures (abnormal readings), Identify data entry errors, Find suspicious user behavior.",
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
                                "description": "Column name to analyze (should be numeric)",
                            },
                            "method": {
                                "type": "string",
                                "enum": ["iqr", "zscore"],
                                "description": "Outlier detection method: iqr (Interquartile Range, robust to extreme outliers), zscore (Z-score, assumes normal distribution). Default: iqr",
                                "default": "iqr",
                            },
                            "threshold": {
                                "type": "number",
                                "description": "Threshold for outlier detection: IQR multiplier (1.5=standard, 3.0=extreme) or Z-score (2.0=95%, 3.0=99.7%). Default: 1.5",
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
                    description="Search for a value across all sheets in the file. Returns list of sheets with match counts and locations. Use for: data location, cross-sheet search, value tracking, data consistency checks, duplicate detection across sheets. EXAMPLES: Find customer ID across all monthly sheets, Locate product code in inventory and sales sheets, Track order number across processing stages.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "column_name": {
                                "type": "string",
                                "description": "Column name to search in (case-insensitive). Searches this column across all sheets.",
                            },
                            "value": {
                                "description": "Value to search for (supports numbers and strings). Exact match search.",
                            },
                        },
                        "required": ["file_path", "column_name", "value"],
                    },
                ),
                Tool(
                    name="compare_sheets",
                    description="Compare data between two sheets using a key column. Returns rows with differences and status (only_in_sheet1, only_in_sheet2, different_values). Use for: version comparison, data reconciliation, change detection, audit trails, before/after analysis. EXAMPLES: Compare current vs previous month data, Detect changes between versions, Reconcile data from two sources, Find missing records.",
                    inputSchema={
                        "type": "object",
                        "properties": {
                            "file_path": {
                                "type": "string",
                                "description": "Absolute path to the Excel file",
                            },
                            "sheet1": {
                                "type": "string",
                                "description": "First sheet name to compare (e.g., 'January' or 'Version1')",
                            },
                            "sheet2": {
                                "type": "string",
                                "description": "Second sheet name to compare (e.g., 'February' or 'Version2')",
                            },
                            "key_column": {
                                "type": "string",
                                "description": "Column to use as key for matching rows between sheets (e.g., 'ID', 'OrderNumber'). Rows with same key are compared.",
                            },
                            "compare_columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Columns to compare for differences (e.g., ['Amount', 'Status']). Only differences in these columns are reported.",
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
                    description="Find duplicate rows based on specified columns. Returns all duplicate rows including first occurrence with row indices. Use for: data quality, duplicate detection, deduplication planning, data integrity checks, identifying redundant records. EXAMPLES: Find duplicate customer records, Detect duplicate orders, Identify duplicate email addresses, Find repeated transactions.",
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
                                "description": "Columns to check for duplicates (checks combination of these columns). Example: [Email] finds duplicate emails, [FirstName, LastName] finds duplicate name combinations.",
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
                    description="Find null/empty values in specified columns. Returns statistics (count, percentage) and row indices for each column. Use for: data quality assessment, missing value analysis, completeness checks, data cleaning, identifying incomplete records. EXAMPLES: Find missing email addresses, Identify incomplete customer records, Check for missing prices, Locate unfilled required fields.",
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
                                "description": "Columns to check for null values (empty cells, NaN, None). Note: Placeholders like '.', '-' are NOT null.",
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
                    description="Calculate period-over-period change (month/quarter/year). Returns periods with values, absolute changes, and percentage changes. Use for: trend analysis, growth tracking, seasonal comparison, performance monitoring, YoY analysis. EXAMPLES: Month-over-month revenue growth, Quarter-over-quarter user growth, Year-over-year sales comparison, Seasonal trend analysis.",
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
                                "description": "Column containing dates (should be datetime format)",
                            },
                            "value_column": {
                                "type": "string",
                                "description": "Column containing values to analyze (should be numeric)",
                            },
                            "period_type": {
                                "type": "string",
                                "enum": ["month", "quarter", "year"],
                                "description": "Period type for grouping: month (monthly comparison), quarter (quarterly comparison), year (yearly comparison)",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "date_column", "value_column", "period_type"],
                    },
                ),
                Tool(
                    name="calculate_running_total",
                    description="Calculate running total (cumulative sum) ordered by a column. Supports grouping for per-category totals. Returns rows with running totals and Excel formulas. Use for: cumulative analysis, progress tracking, balance calculations, hierarchical totals, cash flow tracking. EXAMPLES: Cumulative revenue by date, Running balance by account, Cumulative units sold by product, Progress tracking (cumulative completion).",
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
                                "description": "Column to order by (typically date). Determines the sequence for cumulative calculation.",
                            },
                            "value_column": {
                                "type": "string",
                                "description": "Column containing values to sum (should be numeric)",
                            },
                            "group_by_columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Optional columns to group by (running total resets within each group). Example: [Region] calculates running total per region.",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "order_column", "value_column"],
                    },
                ),
                Tool(
                    name="calculate_moving_average",
                    description="Calculate moving average with specified window size. Useful for trend analysis and smoothing data. Returns rows with moving averages and Excel formulas. Use for: trend detection, noise reduction, smoothing time series, pattern identification, signal processing. EXAMPLES: 7-day moving average of daily sales, 30-day moving average of stock price, 3-month moving average of temperature, Smoothing sensor readings.",
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
                                "description": "Column to order by (typically date). Determines the sequence for moving average calculation.",
                            },
                            "value_column": {
                                "type": "string",
                                "description": "Column containing values to average (should be numeric)",
                            },
                            "window_size": {
                                "type": "integer",
                                "description": "Number of periods for moving average window (e.g., 7 for 7-day average, 30 for 30-day average)",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "order_column", "value_column", "window_size"],
                    },
                ),
                Tool(
                    name="rank_rows",
                    description="Rank rows by column value (ascending or descending). Supports top-N filtering and ranking within groups. Returns ranked rows with rank numbers and Excel formulas. Use for: top/bottom analysis, leaderboards, percentile ranking, category-wise ranking, performance comparison. EXAMPLES: Top 10 customers by revenue, Bottom 5 products by sales, Top salesperson per region, Ranking students by score.",
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
                                "description": "Column to rank by (should be numeric)",
                            },
                            "direction": {
                                "type": "string",
                                "enum": ["asc", "desc"],
                                "description": "Ranking direction: desc (highest first, default), asc (lowest first)",
                                "default": "desc",
                            },
                            "top_n": {
                                "type": "integer",
                                "description": "Return only top N rows (optional, returns all if not specified). Example: top_n=10 returns top 10 ranked rows.",
                            },
                            "group_by_columns": {
                                "type": "array",
                                "items": {"type": "string"},
                                "description": "Optional columns to group by (ranking resets within each group). Example: [Region] ranks within each region separately.",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
                        "required": ["file_path", "sheet_name", "rank_column"],
                    },
                ),
                Tool(
                    name="calculate_expression",
                    description="Calculate expression between columns (e.g., 'Price * Quantity'). Supports arithmetic operations (+, -, *, /, parentheses). Returns calculated values and Excel formulas. Use for: derived metrics, financial calculations, ratio analysis, custom computations, KPI calculations. EXAMPLES: Total = Price * Quantity, Margin = (Revenue - Cost) / Revenue, Discount = Price * 0.1, Efficiency = Output / Input.",
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
                                "description": "Expression to calculate using column names (e.g., 'Price * Quantity', '(Revenue - Cost) / Revenue'). Supports +, -, *, /, parentheses.",
                            },
                            "output_column_name": {
                                "type": "string",
                                "description": "Name for the calculated column (e.g., 'Total', 'Margin', 'Profit')",
                            },
                            "filters": FILTER_PROPERTY_SCHEMA,
                            "logic": {
                                "type": "string",
                                "enum": ["AND", "OR"],
                                "description": LOGIC_DESCRIPTION,
                                "default": "AND",
                            },
                            "header_row": {
                                "type": "integer",
                                "description": "Row index for headers (optional, auto-detected if not provided)",
                            },
                        },
                        "definitions": FILTER_DEFINITIONS,
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

                elif name == "filter_and_count_batch":
                    request = FilterAndCountBatchRequest(**arguments)
                    response = self.data_ops.filter_and_count_batch(request)
                    return [TextContent(type="text", text=response.model_dump_json(indent=2))]

                elif name == "analyze_overlap":
                    request = AnalyzeOverlapRequest(**arguments)
                    response = self.data_ops.analyze_overlap(request)
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
