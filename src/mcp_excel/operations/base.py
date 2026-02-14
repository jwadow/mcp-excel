# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Base class for all operations with common functionality."""

import time
from typing import Any

import pandas as pd
import psutil
from pydantic import BaseModel

from ..core.file_loader import FileLoader
from ..core.header_detector import HeaderDetector
from ..models.responses import FileMetadata, PerformanceMetrics

# Response size limits to prevent agent context overflow
DEFAULT_COLUMN_LIMIT = 5      # Columns returned when not specified
DEFAULT_ROW_LIMIT = 50        # Default limit for rows
MAX_ROW_LIMIT = 1000          # Maximum rows per request (hard cap)
MAX_RESPONSE_CHARS = 50_000   # ~20k tokens for Cyrillic text
MAX_DIFFERENCES = 500         # Maximum differences in compare_sheets


class BaseOperations:
    """Base class for all operations with common functionality."""

    def __init__(self, file_loader: FileLoader) -> None:
        """Initialize base operations.

        Args:
            file_loader: FileLoader instance for loading files
        """
        self._loader = file_loader
        self._header_detector = HeaderDetector()

    def _format_value(self, value: Any) -> Any:
        """Format value for natural display to agent/user.
        
        Converts values to JSON-serializable types:
        - Floats without decimal parts -> ints
        - Datetime values -> ISO 8601 strings
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

    def _validate_response_size(
        self, 
        response: BaseModel, 
        rows_count: int | None = None, 
        columns_count: int | None = None,
        request_limit: int | None = None
    ) -> None:
        """Validate response size to prevent agent context overflow.
        
        Args:
            response: Response object to validate
            rows_count: Number of rows in response (for better error message)
            columns_count: Number of columns in response (for better error message)
            request_limit: Original limit from request (for suggestions)
            
        Raises:
            ValueError: If response exceeds safe character limit
        """
        response_json = response.model_dump_json()
        char_count = len(response_json)
        
        if char_count > MAX_RESPONSE_CHARS:
            # Build detailed error message following MCP philosophy
            error_parts = [
                f"Response too large: {char_count:,} characters (limit: {MAX_RESPONSE_CHARS:,})."
            ]
            
            if rows_count is not None and columns_count is not None:
                error_parts.append(
                    f"Current request returned: {rows_count} rows × {columns_count} columns."
                )
            
            # Explain MCP philosophy: agent should not receive raw data
            error_parts.append(
                "\nMCP Philosophy: Agent should analyze data using atomic operations, not load raw data into context."
            )
            
            # Provide actionable suggestions
            error_parts.append("\nHow to fix:")
            
            suggestion_num = 1
            if request_limit is not None and request_limit > DEFAULT_ROW_LIMIT:
                suggested_limit = max(DEFAULT_ROW_LIMIT, request_limit // 2)
                error_parts.append(
                    f"{suggestion_num}) Reduce 'limit' parameter: current={request_limit}, try={suggested_limit}"
                )
                suggestion_num += 1
            
            if columns_count is not None and columns_count > DEFAULT_COLUMN_LIMIT:
                error_parts.append(
                    f"{suggestion_num}) Specify fewer columns: current={columns_count}, default={DEFAULT_COLUMN_LIMIT}"
                )
                suggestion_num += 1
            
            error_parts.append(
                f"{suggestion_num}) Use MCP atomic operations for analysis instead of retrieving rows"
            )
            
            raise ValueError(" ".join(error_parts))

    def _apply_column_limit(
        self, 
        df: pd.DataFrame, 
        requested_columns: list[str] | None
    ) -> tuple[pd.DataFrame, list[str]]:
        """Apply smart column limit to prevent context overflow.
        
        If no columns specified, returns only first DEFAULT_COLUMN_LIMIT columns.
        
        Args:
            df: DataFrame to limit
            requested_columns: Requested columns (None = apply default limit)
            
        Returns:
            Tuple of (limited_df, actual_columns_used)
        """
        if requested_columns is None or len(requested_columns) == 0:
            # Apply smart default: return only first N columns
            actual_columns = list(df.columns[:DEFAULT_COLUMN_LIMIT])
            return df[actual_columns], actual_columns
        else:
            # Use requested columns
            return df[requested_columns], requested_columns

    def _enforce_row_limit(self, limit: int) -> int:
        """Enforce maximum row limit to prevent context overflow.
        
        Args:
            limit: Requested limit
            
        Returns:
            Enforced limit (capped at MAX_ROW_LIMIT)
        """
        if limit > MAX_ROW_LIMIT:
            return MAX_ROW_LIMIT
        return limit
