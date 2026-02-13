"""Pydantic models for API responses."""

from typing import Any, Optional

from pydantic import BaseModel, Field


class ExcelOutput(BaseModel):
    """Excel-specific output for copy-paste functionality."""

    tsv: Optional[str] = Field(default=None, description="TSV-formatted data ready for Ctrl+V in Excel")
    formula: Optional[str] = Field(default=None, description="Excel formula for dynamic calculation")
    references: Optional[dict[str, Any]] = Field(default=None, description="Cell/range references used in formula")


class PerformanceMetrics(BaseModel):
    """Performance metrics for the operation."""

    execution_time_ms: float = Field(description="Execution time in milliseconds")
    rows_processed: int = Field(description="Number of rows processed")
    cache_hit: bool = Field(description="Whether result was served from cache")
    memory_used_mb: float = Field(description="Memory used by operation in MB")


class FileMetadata(BaseModel):
    """Metadata about the Excel file."""

    file_format: str = Field(description="File format")
    sheet_name: Optional[str] = Field(default=None, description="Sheet name if applicable")
    rows_total: Optional[int] = Field(default=None, description="Total rows in sheet")
    columns_total: Optional[int] = Field(default=None, description="Total columns in sheet")


class HeaderDetectionInfo(BaseModel):
    """Information about header detection."""

    header_row: int = Field(description="Detected header row index")
    confidence: float = Field(description="Confidence score (0.0 to 1.0)")
    candidates: Optional[list[dict[str, Any]]] = Field(default=None, description="Alternative candidates if confidence is low")


class InspectFileResponse(BaseModel):
    """Response for file inspection."""

    format: str = Field(description="File format")
    size_bytes: int = Field(description="File size in bytes")
    size_mb: float = Field(description="File size in megabytes")
    sheet_count: int = Field(description="Number of sheets")
    sheet_names: list[str] = Field(description="List of sheet names")
    sheets_info: list[dict[str, Any]] = Field(description="Basic info for each sheet")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class GetSheetInfoResponse(BaseModel):
    """Response for sheet information."""

    sheet_name: str = Field(description="Sheet name")
    column_names: list[str] = Field(description="List of column names")
    column_count: int = Field(description="Number of columns")
    column_types: dict[str, str] = Field(description="Data types for each column")
    row_count: int = Field(description="Number of data rows")
    data_start_row: int = Field(description="Row index where data starts")
    sample_rows: list[dict[str, Any]] = Field(description="First 3 rows as sample")
    header_detection: Optional[HeaderDetectionInfo] = Field(default=None, description="Header detection info if auto-detected")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class GetColumnNamesResponse(BaseModel):
    """Response for column names."""

    column_names: list[str] = Field(description="List of column names")
    column_count: int = Field(description="Number of columns")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class FindColumnResponse(BaseModel):
    """Response for column search."""

    found_in: list[dict[str, Any]] = Field(description="List of sheets where column was found")
    total_matches: int = Field(description="Total number of matches")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class GetUniqueValuesResponse(BaseModel):
    """Response for unique values."""

    values: list[Any] = Field(description="List of unique values")
    count: int = Field(description="Number of unique values")
    truncated: bool = Field(description="Whether result was truncated due to limit")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class GetValueCountsResponse(BaseModel):
    """Response for value counts."""

    value_counts: dict[str, int] = Field(description="Dictionary of value -> count")
    total_values: int = Field(description="Total number of values")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class FilterAndCountResponse(BaseModel):
    """Response for filtered count."""

    count: int = Field(description="Number of rows matching filters")
    filters_applied: list[dict[str, Any]] = Field(description="Filters that were applied")
    excel_output: ExcelOutput = Field(description="Excel-formatted output with formula")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class FilterAndGetRowsResponse(BaseModel):
    """Response for filtered rows."""

    rows: list[dict[str, Any]] = Field(description="Filtered rows as list of dictionaries")
    count: int = Field(description="Number of rows returned")
    total_matches: int = Field(description="Total number of matching rows (before limit/offset)")
    truncated: bool = Field(description="Whether result was truncated")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class AggregateResponse(BaseModel):
    """Response for aggregation."""

    value: float = Field(description="Aggregated value")
    operation: str = Field(description="Operation performed")
    target_column: str = Field(description="Column that was aggregated")
    filters_applied: list[dict[str, Any]] = Field(description="Filters that were applied")
    excel_output: ExcelOutput = Field(description="Excel-formatted output with formula")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class GroupByResponse(BaseModel):
    """Response for group-by aggregation."""

    groups: list[dict[str, Any]] = Field(description="Grouped data with aggregated values")
    group_columns: list[str] = Field(description="Columns used for grouping")
    agg_column: str = Field(description="Column that was aggregated")
    agg_operation: str = Field(description="Aggregation operation")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class CorrelateResponse(BaseModel):
    """Response for correlation analysis."""

    correlation_matrix: dict[str, dict[str, float]] = Field(description="Correlation matrix")
    method: str = Field(description="Correlation method used")
    columns: list[str] = Field(description="Columns analyzed")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class ColumnStats(BaseModel):
    """Statistical summary of a column."""

    count: int = Field(description="Number of non-null values")
    mean: Optional[float] = Field(default=None, description="Mean value")
    median: Optional[float] = Field(default=None, description="Median value")
    std: Optional[float] = Field(default=None, description="Standard deviation")
    min: Optional[float] = Field(default=None, description="Minimum value")
    max: Optional[float] = Field(default=None, description="Maximum value")
    q25: Optional[float] = Field(default=None, description="25th percentile")
    q75: Optional[float] = Field(default=None, description="75th percentile")
    null_count: int = Field(description="Number of null values")


class GetColumnStatsResponse(BaseModel):
    """Response for column statistics."""

    column: str = Field(description="Column name")
    stats: ColumnStats = Field(description="Statistical summary")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class DetectOutliersResponse(BaseModel):
    """Response for outlier detection."""

    outliers: list[dict[str, Any]] = Field(description="Rows containing outliers")
    outlier_count: int = Field(description="Number of outliers detected")
    method: str = Field(description="Detection method used")
    threshold: float = Field(description="Threshold used")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class CompareSheetsResponse(BaseModel):
    """Response for sheet comparison."""

    differences: list[dict[str, Any]] = Field(description="Rows with differences")
    difference_count: int = Field(description="Number of differences found")
    key_column: str = Field(description="Column used as key")
    compare_columns: list[str] = Field(description="Columns compared")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class SearchAcrossSheetsResponse(BaseModel):
    """Response for cross-sheet search."""

    matches: list[dict[str, Any]] = Field(description="Sheets with matches and counts")
    total_matches: int = Field(description="Total number of matches across all sheets")
    column_name: str = Field(description="Column searched")
    value: Any = Field(description="Value searched for")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class FindDuplicatesResponse(BaseModel):
    """Response for duplicate detection."""

    duplicates: list[dict[str, Any]] = Field(description="Duplicate rows")
    duplicate_count: int = Field(description="Number of duplicate rows")
    columns_checked: list[str] = Field(description="Columns checked for duplicates")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class FindNullsResponse(BaseModel):
    """Response for null detection."""

    null_info: dict[str, dict[str, Any]] = Field(description="Null information per column")
    total_nulls: int = Field(description="Total number of null values")
    columns_checked: list[str] = Field(description="Columns checked for nulls")
    excel_output: ExcelOutput = Field(description="Excel-formatted output")
    metadata: FileMetadata = Field(description="File metadata")
    performance: PerformanceMetrics = Field(description="Performance metrics")


class ErrorResponse(BaseModel):
    """Error response."""

    error: str = Field(description="Error type")
    message: str = Field(description="Error message")
    details: Optional[dict[str, Any]] = Field(default=None, description="Additional error details")
    suggestion: Optional[str] = Field(default=None, description="Suggestion for fixing the error")
    recoverable: bool = Field(default=True, description="Whether error is recoverable")
