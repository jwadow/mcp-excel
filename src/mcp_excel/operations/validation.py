# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Data validation operations for Excel files."""

import time
from typing import Any

import pandas as pd
import psutil

from ..core.file_loader import FileLoader
from ..core.header_detector import HeaderDetector
from ..excel.tsv_formatter import TSVFormatter
from ..models.requests import (
    FindDuplicatesRequest,
    FindNullsRequest,
)
from ..models.responses import (
    ExcelOutput,
    FileMetadata,
    FindDuplicatesResponse,
    FindNullsResponse,
    PerformanceMetrics,
)


class ValidationOperations:
    """Data validation operations for Excel data."""

    def __init__(self, file_loader: FileLoader) -> None:
        """Initialize validation operations.

        Args:
            file_loader: FileLoader instance for loading files
        """
        self._loader = file_loader
        self._header_detector = HeaderDetector()
        self._tsv_formatter = TSVFormatter()

    def _format_value(self, value: Any) -> Any:
        """Format value for natural display to agent/user.
        
        Converts values to JSON-serializable types:
        - Floats without decimal parts -> ints
        - Datetime values -> ISO 8601 strings (per DATE_TIME_ARCHITECTURE.md)
        - NaN/NaT -> None

        Args:
            value: Value to format

        Returns:
            Formatted value (JSON-serializable)
        """
        if pd.isna(value):
            return None
        elif isinstance(value, (pd.Timestamp, pd.DatetimeTZDtype)):
            # Convert datetime to ISO 8601 string for agent
            return value.isoformat()
        elif pd.api.types.is_datetime64_any_dtype(type(value)):
            # Handle numpy datetime64
            return pd.Timestamp(value).isoformat()
        elif isinstance(value, float) and value.is_integer():
            return int(value)
        else:
            return value

    def _get_performance_metrics(
        self, start_time: float, rows_processed: int, cache_hit: bool
    ) -> PerformanceMetrics:
        """Create performance metrics.

        Args:
            start_time: Operation start time
            rows_processed: Number of rows processed
            cache_hit: Whether cache was used

        Returns:
            PerformanceMetrics object
        """
        execution_time = (time.time() - start_time) * 1000
        process = psutil.Process()
        memory_mb = process.memory_info().rss / 1024 / 1024

        return PerformanceMetrics(
            execution_time_ms=round(execution_time, 2),
            rows_processed=rows_processed,
            cache_hit=cache_hit,
            memory_used_mb=round(memory_mb, 2),
        )

    def _get_file_metadata(
        self, file_path: str, sheet_name: str | None = None
    ) -> FileMetadata:
        """Get file metadata.

        Args:
            file_path: Path to file
            sheet_name: Optional sheet name

        Returns:
            FileMetadata object
        """
        file_info = self._loader.get_file_info(file_path)
        return FileMetadata(
            file_format=file_info["format"],
            sheet_name=sheet_name,
            rows_total=None,
            columns_total=None,
        )

    def _load_with_header_detection(
        self, file_path: str, sheet_name: str, header_row: int | None
    ) -> tuple[pd.DataFrame, int]:
        """Load DataFrame with header detection.

        Args:
            file_path: Path to file
            sheet_name: Sheet name
            header_row: Optional header row index

        Returns:
            Tuple of (DataFrame, header_row_used)
        """
        if header_row is not None:
            df = self._loader.load(file_path, sheet_name, header_row=header_row, use_cache=True)
            detected_row = header_row
        else:
            df_preview = self._loader.load(file_path, sheet_name, header_row=None, use_cache=True)
            detection_result = self._header_detector.detect(df_preview)
            
            # Always trust the detector - it picks the best candidate from first 20 rows
            detected_row = detection_result.header_row
            df = self._loader.load(file_path, sheet_name, header_row=detected_row, use_cache=True)

        # Normalize column names to strings
        df.columns = [str(col) for col in df.columns]

        return df, detected_row

    def find_duplicates(self, request: FindDuplicatesRequest) -> FindDuplicatesResponse:
        """Find duplicate rows based on specified columns.

        Args:
            request: FindDuplicatesRequest with parameters

        Returns:
            FindDuplicatesResponse with duplicate rows

        Raises:
            ValueError: If columns don't exist
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate columns exist
        missing_cols = [col for col in request.columns if col not in df.columns]
        if missing_cols:
            available = ", ".join(df.columns.tolist())
            raise ValueError(
                f"Columns not found: {', '.join(missing_cols)}. "
                f"Available columns: {available}"
            )

        # Find duplicates
        # duplicated() marks all duplicates except the first occurrence
        # keep=False marks all duplicates including first occurrence
        duplicate_mask = df.duplicated(subset=request.columns, keep=False)
        duplicate_df = df[duplicate_mask]

        # Convert to list of dicts
        duplicates = []
        for idx in range(len(duplicate_df)):
            row_dict = duplicate_df.iloc[idx].to_dict()
            # Format values and add row index
            formatted_dict = {str(k): self._format_value(v) for k, v in row_dict.items()}
            formatted_dict["_row_index"] = int(duplicate_df.index[idx])
            duplicates.append(formatted_dict)

        # Generate TSV output
        if duplicates:
            # Include all columns plus row index
            headers = ["_row_index"] + df.columns.tolist()
            rows = []

            for dup in duplicates:
                tsv_row = [dup.get(col) for col in headers]
                rows.append(tsv_row)

            tsv = self._tsv_formatter.format_table(headers, rows)
        else:
            tsv = "No duplicates found"

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        # Create response
        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), False)

        return FindDuplicatesResponse(
            duplicates=duplicates,
            duplicate_count=len(duplicates),
            columns_checked=request.columns,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

    def find_nulls(self, request: FindNullsRequest) -> FindNullsResponse:
        """Find null/empty values in specified columns.

        Args:
            request: FindNullsRequest with parameters

        Returns:
            FindNullsResponse with null statistics per column

        Raises:
            ValueError: If columns don't exist
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate columns exist
        missing_cols = [col for col in request.columns if col not in df.columns]
        if missing_cols:
            available = ", ".join(df.columns.tolist())
            raise ValueError(
                f"Columns not found: {', '.join(missing_cols)}. "
                f"Available columns: {available}"
            )

        # Analyze nulls for each column
        null_info = {}
        total_nulls = 0

        for col in request.columns:
            null_mask = df[col].isna()
            null_count = int(null_mask.sum())
            total_nulls += null_count

            # Get indices of null rows
            null_indices = df[null_mask].index.tolist()

            # Calculate percentage
            null_percentage = (null_count / len(df) * 100) if len(df) > 0 else 0

            null_info[col] = {
                "null_count": null_count,
                "null_percentage": round(null_percentage, 2),
                "total_rows": len(df),
                "null_indices": [int(idx) for idx in null_indices[:100]],  # Limit to first 100
                "truncated": len(null_indices) > 100,
            }

        # Generate TSV output
        headers = ["Column", "Null Count", "Percentage", "Total Rows"]
        rows = []

        for col in request.columns:
            info = null_info[col]
            rows.append([
                col,
                info["null_count"],
                f"{info['null_percentage']}%",
                info["total_rows"]
            ])

        tsv = self._tsv_formatter.format_table(headers, rows)

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        # Create response
        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), False)

        return FindNullsResponse(
            null_info=null_info,
            total_nulls=total_nulls,
            columns_checked=request.columns,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )
