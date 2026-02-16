# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Base class for all operations with common functionality."""

import time
import unicodedata
from difflib import get_close_matches
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
MAX_RESPONSE_CHARS = 10_000   # ~4k tokens for text
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
        - NumPy scalar types (int64, float64, etc.) -> Python native types
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
        # Handle numpy scalar types (int8, int16, int32, int64, float32, float64, etc.)
        # All numpy scalars have .item() method to convert to Python native types
        elif hasattr(value, 'item'):
            return value.item()
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

    def _normalize_column_name(self, name: str) -> str:
        """Normalize column name for robust matching.
        
        Handles:
        - Non-string column names (pandas can use integers: 0, 1, 2...)
        - Unicode normalization (NFC - composed form)
        - Non-breaking spaces (U+00A0) → regular spaces (U+0020)
        - Leading/trailing whitespace
        - Multiple consecutive spaces → single space
        
        Args:
            name: Column name to normalize (accepts str or int)
        
        Returns:
            Normalized column name (always str)
        
        Examples:
            >>> _normalize_column_name("café")  # NFC or NFD
            "café"  # Always NFC
            >>> _normalize_column_name("Нетто,\u00A0кг")  # Non-breaking space
            "Нетто, кг"  # Regular space
            >>> _normalize_column_name("  Name  ")
            "Name"
            >>> _normalize_column_name(0)  # Integer column name
            "0"
        """
        # 0. Handle non-string column names (defensive programming)
        # Pandas can use integers (0, 1, 2...) when headers are not detected
        # This makes BaseOperations robust for all usage scenarios
        if not isinstance(name, str):
            name = str(name)
        
        # 1. Unicode normalization (NFC - composed form)
        # This ensures "café" (NFC) and "café" (NFD) are treated as identical
        name = unicodedata.normalize('NFC', name)
        
        # 2. Replace non-breaking spaces (U+00A0) with regular spaces (U+0020)
        # Excel often uses non-breaking spaces, which look identical but compare differently
        name = name.replace('\u00A0', ' ')
        
        # 3. Strip leading/trailing whitespace
        name = name.strip()
        
        # 4. Collapse multiple consecutive spaces into one
        # "Name  Value" → "Name Value"
        name = ' '.join(name.split())
        
        return name

    def _find_column(
        self,
        df: pd.DataFrame,
        column_name: str,
        context: str = "operation"
    ) -> str:
        """Find column in DataFrame using normalized matching.
        
        Uses Unicode normalization to handle:
        - NFC/NFD forms (composed vs decomposed)
        - Non-breaking spaces (U+00A0)
        - Leading/trailing/multiple whitespace
        
        Args:
            df: DataFrame to search in
            column_name: Column name to find (will be normalized)
            context: Context for error message (e.g., "filter", "aggregation")
        
        Returns:
            Original column name from DataFrame (not normalized)
        
        Raises:
            ValueError: If column not found (with fuzzy suggestions)
        
        Example:
            >>> df = pd.DataFrame({"café": [1, 2, 3]})  # NFC in DataFrame
            >>> _find_column(df, "café")  # NFD in request
            "café"  # Returns original NFC name from DataFrame
        """
        # Normalize requested column name
        normalized_request = self._normalize_column_name(column_name)
        
        # Build mapping: normalized name → original DataFrame column name
        normalized_to_original = {
            self._normalize_column_name(col): col
            for col in df.columns
        }
        
        # Find column using normalized comparison
        if normalized_request in normalized_to_original:
            # Return original column name from DataFrame
            return normalized_to_original[normalized_request]
        
        # Column not found - provide helpful error with fuzzy matching
        suggestions = get_close_matches(
            normalized_request,
            normalized_to_original.keys(),
            n=3,
            cutoff=0.6
        )
        
        available = ", ".join(str(col) for col in df.columns)
        suggestion_text = ""
        if suggestions:
            # Map back to original names for suggestions
            original_suggestions = [
                normalized_to_original[s] for s in suggestions
            ]
            suggestion_text = f" Did you mean: {', '.join(repr(s) for s in original_suggestions)}?"
        
        raise ValueError(
            f"Column '{column_name}' not found in {context}.{suggestion_text} "
            f"Available columns: {available}"
        )

    def _find_columns(
        self,
        df: pd.DataFrame,
        column_names: list[str],
        context: str = "operation"
    ) -> list[str]:
        """Find multiple columns in DataFrame using normalized matching.
        
        Args:
            df: DataFrame to search in
            column_names: List of column names to find
            context: Context for error message
        
        Returns:
            List of original column names from DataFrame
        
        Raises:
            ValueError: If any column not found
        """
        return [self._find_column(df, col, context) for col in column_names]

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

    def _ensure_numeric_column(
        self,
        col_data: pd.Series,
        col_name: str,
        min_conversion_rate: float = 0.5
    ) -> pd.Series:
        """Convert column to numeric or raise informative error.
        
        Attempts to convert object/string columns to numeric type.
        If conversion fails for too many values, raises a clear error message.
        
        Args:
            col_data: Column data to convert
            col_name: Column name for error messages
            min_conversion_rate: Minimum % of values that must convert (default 50%)
            
        Returns:
            Numeric Series (float64 dtype)
            
        Raises:
            ValueError: If column is not numeric and cannot be converted
        """
        # If already numeric, return as-is
        if pd.api.types.is_numeric_dtype(col_data):
            return col_data
        
        # Try to convert object/string columns to numeric
        if col_data.dtype == 'object' or col_data.dtype.name == 'string':
            col_numeric = pd.to_numeric(col_data, errors='coerce')
            non_null_original = col_data.notna().sum()
            non_null_converted = col_numeric.notna().sum()
            
            # Check if conversion was successful (at least min_conversion_rate % converted)
            if non_null_converted >= non_null_original * min_conversion_rate:
                return col_numeric
            else:
                raise ValueError(
                    f"Column '{col_name}' is not numeric. "
                    f"Only {non_null_converted}/{non_null_original} values could be converted to numbers."
                )
        
        # If not numeric and not object/string, raise error
        raise ValueError(
            f"Column '{col_name}' must be numeric for this operation. "
            f"Current type: {col_data.dtype}"
        )
