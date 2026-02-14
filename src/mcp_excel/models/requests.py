# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Pydantic models for API requests."""

from typing import Any, Literal, Optional, Union

from pydantic import BaseModel, Field


class FilterCondition(BaseModel):
    """Single filter condition."""

    column: str = Field(description="Column name to filter on")
    operator: Literal[
        "==", "!=", ">", "<", ">=", "<=",
        "in", "not_in",
        "contains", "startswith", "endswith", "regex",
        "is_null", "is_not_null"
    ] = Field(description="Comparison operator")
    value: Optional[Any] = Field(default=None, description="Value to compare against (for single-value operators)")
    values: Optional[list[Any]] = Field(default=None, description="List of values (for 'in' and 'not_in' operators)")


class FilterGroup(BaseModel):
    """Group of filter conditions with logic operator."""

    filters: list[Union[FilterCondition, "FilterGroup"]] = Field(description="List of filter conditions or nested groups")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for combining filters")


class InspectFileRequest(BaseModel):
    """Request to inspect Excel file structure."""

    file_path: str = Field(description="Absolute path to the Excel file")


class GetSheetInfoRequest(BaseModel):
    """Request to get detailed sheet information."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class GetColumnNamesRequest(BaseModel):
    """Request to get column names from a sheet."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class FindColumnRequest(BaseModel):
    """Request to find column across sheets."""

    file_path: str = Field(description="Absolute path to the Excel file")
    column_name: str = Field(description="Column name to search for")
    search_all_sheets: bool = Field(default=True, description="Search in all sheets or just first one")


class GetUniqueValuesRequest(BaseModel):
    """Request to get unique values from a column."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    column: str = Field(description="Column name")
    limit: int = Field(default=100, description="Maximum number of unique values to return")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class GetValueCountsRequest(BaseModel):
    """Request to get value counts (frequency) from a column."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    column: str = Field(description="Column name")
    top_n: int = Field(default=10, description="Number of top values to return")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class FilterAndCountRequest(BaseModel):
    """Request to count rows matching filter conditions."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    filters: list[FilterCondition] = Field(description="List of filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for combining filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class FilterAndGetRowsRequest(BaseModel):
    """Request to get rows matching filter conditions."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    filters: list[FilterCondition] = Field(description="List of filter conditions")
    columns: Optional[list[str]] = Field(default=None, description="Columns to return (None = all columns)")
    limit: int = Field(default=50, description="Maximum number of rows to return")
    offset: int = Field(default=0, description="Number of rows to skip")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for combining filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class AggregateRequest(BaseModel):
    """Request to perform aggregation on a column."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    operation: Literal["sum", "mean", "median", "min", "max", "std", "var", "count"] = Field(
        description="Aggregation operation"
    )
    target_column: str = Field(description="Column to aggregate")
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class GroupByRequest(BaseModel):
    """Request to perform group-by aggregation."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    group_columns: list[str] = Field(description="Columns to group by")
    agg_column: str = Field(description="Column to aggregate")
    agg_operation: Literal["sum", "mean", "median", "min", "max", "std", "var", "count"] = Field(
        description="Aggregation operation"
    )
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class CorrelateRequest(BaseModel):
    """Request to calculate correlation between columns."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    columns: list[str] = Field(description="Columns to correlate (minimum 2)")
    method: Literal["pearson", "spearman", "kendall"] = Field(default="pearson", description="Correlation method")
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class GetColumnStatsRequest(BaseModel):
    """Request to get statistical summary of a column."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    column: str = Field(description="Column name")
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class DetectOutliersRequest(BaseModel):
    """Request to detect outliers in a column."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    column: str = Field(description="Column name")
    method: Literal["iqr", "zscore"] = Field(default="iqr", description="Outlier detection method")
    threshold: float = Field(default=1.5, description="Threshold for outlier detection (IQR multiplier or Z-score)")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class CompareSheetsRequest(BaseModel):
    """Request to compare data between two sheets."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet1: str = Field(description="First sheet name")
    sheet2: str = Field(description="Second sheet name")
    key_column: str = Field(description="Column to use as key for comparison")
    compare_columns: list[str] = Field(description="Columns to compare")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class SearchAcrossSheetsRequest(BaseModel):
    """Request to search for a value across all sheets."""

    file_path: str = Field(description="Absolute path to the Excel file")
    column_name: str = Field(description="Column name to search in")
    value: Any = Field(description="Value to search for")


class FindDuplicatesRequest(BaseModel):
    """Request to find duplicate rows."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    columns: list[str] = Field(description="Columns to check for duplicates")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class FindNullsRequest(BaseModel):
    """Request to find null/empty values."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    columns: list[str] = Field(description="Columns to check for nulls")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class CalculatePeriodChangeRequest(BaseModel):
    """Request to calculate period-over-period change."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    date_column: str = Field(description="Column containing dates")
    value_column: str = Field(description="Column containing values to analyze")
    period_type: Literal["month", "quarter", "year"] = Field(description="Period type for grouping")
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class CalculateRunningTotalRequest(BaseModel):
    """Request to calculate running total (cumulative sum)."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    order_column: str = Field(description="Column to order by (typically date)")
    value_column: str = Field(description="Column containing values to sum")
    group_by_columns: Optional[list[str]] = Field(default=None, description="Optional columns to group by")
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class CalculateMovingAverageRequest(BaseModel):
    """Request to calculate moving average."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    order_column: str = Field(description="Column to order by (typically date)")
    value_column: str = Field(description="Column containing values to average")
    window_size: int = Field(description="Number of periods for moving average window")
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class RankRowsRequest(BaseModel):
    """Request to rank rows by a column value."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    rank_column: str = Field(description="Column to rank by")
    direction: Literal["asc", "desc"] = Field(default="desc", description="Ranking direction (desc = highest first)")
    top_n: Optional[int] = Field(default=None, description="Return only top N rows (None = all rows)")
    group_by_columns: Optional[list[str]] = Field(default=None, description="Optional columns to group by for ranking within groups")
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")


class CalculateExpressionRequest(BaseModel):
    """Request to calculate expression between columns."""

    file_path: str = Field(description="Absolute path to the Excel file")
    sheet_name: str = Field(description="Name of the sheet")
    expression: str = Field(description="Expression to calculate (e.g., 'Price * Quantity')")
    output_column_name: str = Field(description="Name for the calculated column")
    filters: list[FilterCondition] = Field(default_factory=list, description="Optional filter conditions")
    logic: Literal["AND", "OR"] = Field(default="AND", description="Logic operator for filters")
    header_row: Optional[int] = Field(default=None, description="Row index for headers (None = auto-detect)")
