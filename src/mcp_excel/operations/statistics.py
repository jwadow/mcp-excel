# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Statistical operations for Excel data analysis."""

import time
from typing import Any

import pandas as pd
import psutil

from ..core.file_loader import FileLoader
from ..core.header_detector import HeaderDetector
from ..excel.tsv_formatter import TSVFormatter
from ..models.requests import (
    CorrelateRequest,
    DetectOutliersRequest,
    GetColumnStatsRequest,
)
from ..models.responses import (
    ColumnStats,
    CorrelateResponse,
    DetectOutliersResponse,
    ExcelOutput,
    FileMetadata,
    GetColumnStatsResponse,
    PerformanceMetrics,
)
from ..operations.filtering import FilterEngine


class StatisticsOperations:
    """Statistical analysis operations for Excel data."""

    def __init__(self, file_loader: FileLoader) -> None:
        """Initialize statistics operations.

        Args:
            file_loader: FileLoader instance for loading files
        """
        self._loader = file_loader
        self._header_detector = HeaderDetector()
        self._filter_engine = FilterEngine()
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

    def get_column_stats(self, request: GetColumnStatsRequest) -> GetColumnStatsResponse:
        """Get statistical summary of a column.

        Args:
            request: GetColumnStatsRequest with parameters

        Returns:
            GetColumnStatsResponse with statistics

        Raises:
            ValueError: If column doesn't exist or is not numeric
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate column exists
        if request.column not in df.columns:
            available = ", ".join(df.columns.tolist())
            raise ValueError(
                f"Column '{request.column}' not found. Available columns: {available}"
            )

        # Apply filters if provided
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Get column data
        col_data = df[request.column]

        # Try to convert to numeric if it's object/string type
        if col_data.dtype == 'object' or col_data.dtype.name == 'string':
            col_data_numeric = pd.to_numeric(col_data, errors='coerce')
            non_null_original = col_data.notna().sum()
            non_null_converted = col_data_numeric.notna().sum()

            # If we didn't lose too much data (>50%), use numeric version
            if non_null_converted >= non_null_original * 0.5:
                col_data = col_data_numeric
            else:
                raise ValueError(
                    f"Column '{request.column}' is not numeric. "
                    f"Only {non_null_converted}/{non_null_original} values could be converted to numbers."
                )

        # Check if column is numeric
        if not pd.api.types.is_numeric_dtype(col_data):
            raise ValueError(
                f"Column '{request.column}' must be numeric for statistical analysis. "
                f"Current type: {col_data.dtype}"
            )

        # Calculate statistics
        null_count = int(col_data.isna().sum())
        non_null_data = col_data.dropna()

        if len(non_null_data) == 0:
            raise ValueError(f"Column '{request.column}' has no non-null numeric values")

        stats = ColumnStats(
            count=int(len(non_null_data)),
            mean=float(non_null_data.mean()),
            median=float(non_null_data.median()),
            std=float(non_null_data.std()) if len(non_null_data) > 1 else None,
            min=self._format_value(non_null_data.min()),
            max=self._format_value(non_null_data.max()),
            q25=float(non_null_data.quantile(0.25)),
            q75=float(non_null_data.quantile(0.75)),
            null_count=null_count,
        )

        # Generate TSV output
        headers = ["Statistic", "Value"]
        rows = [
            ["Count", stats.count],
            ["Mean", round(stats.mean, 2)],
            ["Median", round(stats.median, 2)],
            ["Std Dev", round(stats.std, 2) if stats.std else "N/A"],
            ["Min", stats.min],
            ["Max", stats.max],
            ["25th Percentile", round(stats.q25, 2)],
            ["75th Percentile", round(stats.q75, 2)],
            ["Null Count", stats.null_count],
        ]
        tsv = self._tsv_formatter.format_table(headers, rows)

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        # Create response
        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), False)

        return GetColumnStatsResponse(
            column=request.column,
            stats=stats,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

    def correlate(self, request: CorrelateRequest) -> CorrelateResponse:
        """Calculate correlation between columns.

        Args:
            request: CorrelateRequest with parameters

        Returns:
            CorrelateResponse with correlation matrix

        Raises:
            ValueError: If columns don't exist or are not numeric
        """
        start_time = time.time()

        # Validate minimum columns
        if len(request.columns) < 2:
            raise ValueError("At least 2 columns are required for correlation analysis")

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate all columns exist
        missing_cols = [col for col in request.columns if col not in df.columns]
        if missing_cols:
            available = ", ".join(df.columns.tolist())
            raise ValueError(
                f"Columns not found: {', '.join(missing_cols)}. "
                f"Available columns: {available}"
            )

        # Apply filters if provided
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Select only requested columns
        df_subset = df[request.columns].copy()

        # Convert to numeric, handling text-stored numbers
        for col in request.columns:
            if df_subset[col].dtype == 'object' or df_subset[col].dtype.name == 'string':
                df_subset[col] = pd.to_numeric(df_subset[col], errors='coerce')

        # Check if all columns are numeric
        non_numeric = [col for col in request.columns if not pd.api.types.is_numeric_dtype(df_subset[col])]
        if non_numeric:
            raise ValueError(
                f"All columns must be numeric for correlation. "
                f"Non-numeric columns: {', '.join(non_numeric)}"
            )

        # Drop rows with any NaN values
        df_clean = df_subset.dropna()

        if len(df_clean) < 2:
            raise ValueError(
                f"Not enough data for correlation analysis. "
                f"Only {len(df_clean)} rows remain after removing nulls (minimum 2 required)"
            )

        # Calculate correlation matrix
        corr_matrix = df_clean.corr(method=request.method)

        # Convert to nested dict format
        correlation_dict = {}
        for col1 in request.columns:
            correlation_dict[col1] = {}
            for col2 in request.columns:
                correlation_dict[col1][col2] = round(float(corr_matrix.loc[col1, col2]), 4)

        # Generate TSV output (correlation matrix)
        headers = [""] + request.columns
        rows = []
        for col1 in request.columns:
            row = [col1]
            for col2 in request.columns:
                row.append(round(correlation_dict[col1][col2], 4))
            rows.append(row)

        tsv = self._tsv_formatter.format_table(headers, rows)

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        # Create response
        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df_clean), False)

        return CorrelateResponse(
            correlation_matrix=correlation_dict,
            method=request.method,
            columns=request.columns,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

    def detect_outliers(self, request: DetectOutliersRequest) -> DetectOutliersResponse:
        """Detect outliers in a column.

        Args:
            request: DetectOutliersRequest with parameters

        Returns:
            DetectOutliersResponse with outlier information

        Raises:
            ValueError: If column doesn't exist or is not numeric
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate column exists
        if request.column not in df.columns:
            available = ", ".join(df.columns.tolist())
            raise ValueError(
                f"Column '{request.column}' not found. Available columns: {available}"
            )

        # Get column data
        col_data = df[request.column].copy()

        # Try to convert to numeric if it's object/string type
        if col_data.dtype == 'object' or col_data.dtype.name == 'string':
            col_data = pd.to_numeric(col_data, errors='coerce')

        # Check if column is numeric
        if not pd.api.types.is_numeric_dtype(col_data):
            raise ValueError(
                f"Column '{request.column}' must be numeric for outlier detection. "
                f"Current type: {col_data.dtype}"
            )

        # Remove NaN values for calculation
        non_null_data = col_data.dropna()

        if len(non_null_data) < 4:
            raise ValueError(
                f"Not enough data for outlier detection. "
                f"Only {len(non_null_data)} non-null values (minimum 4 required)"
            )

        # Detect outliers based on method
        if request.method == "iqr":
            # IQR method
            q1 = non_null_data.quantile(0.25)
            q3 = non_null_data.quantile(0.75)
            iqr = q3 - q1
            lower_bound = q1 - request.threshold * iqr
            upper_bound = q3 + request.threshold * iqr

            outlier_mask = (col_data < lower_bound) | (col_data > upper_bound)

        elif request.method == "zscore":
            # Z-score method
            mean = non_null_data.mean()
            std = non_null_data.std()

            if std == 0:
                raise ValueError("Standard deviation is 0, cannot use Z-score method")

            z_scores = (col_data - mean) / std
            outlier_mask = z_scores.abs() > request.threshold

        else:
            raise ValueError(f"Unknown outlier detection method: {request.method}")

        # Get outlier rows
        outlier_indices = df[outlier_mask].index.tolist()
        outlier_rows = []

        for idx in outlier_indices:
            row_dict = df.loc[idx].to_dict()
            # Format values
            row_dict = {k: self._format_value(v) for k, v in row_dict.items()}
            row_dict["_row_index"] = int(idx)
            outlier_rows.append(row_dict)

        # Generate TSV output
        if outlier_rows:
            # Include all columns plus row index
            headers = ["_row_index"] + df.columns.tolist()
            rows = []

            for row in outlier_rows:
                tsv_row = [row.get(col) for col in headers]
                rows.append(tsv_row)

            tsv = self._tsv_formatter.format_table(headers, rows)
        else:
            tsv = "No outliers detected"

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        # Create response
        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), False)

        return DetectOutliersResponse(
            outliers=outlier_rows,
            outlier_count=len(outlier_rows),
            method=request.method,
            threshold=request.threshold,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )
