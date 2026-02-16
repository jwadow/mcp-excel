# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Statistical operations for Excel data analysis."""

import time

import pandas as pd

from ..core.file_loader import FileLoader
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
    GetColumnStatsResponse,
)
from ..operations.base import BaseOperations
from ..operations.filtering import FilterEngine


class StatisticsOperations(BaseOperations):
    """Statistical analysis operations for Excel data."""

    def __init__(self, file_loader: FileLoader) -> None:
        """Initialize statistics operations.

        Args:
            file_loader: FileLoader instance for loading files
        """
        super().__init__(file_loader)
        self._filter_engine = FilterEngine()
        self._tsv_formatter = TSVFormatter()

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

        # Find column using normalized matching
        actual_column = self._find_column(df, request.column, context="get_column_stats")

        # Apply filters if provided
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Get column data
        col_data = df[actual_column]

        # Ensure column is numeric (with auto-conversion for text-stored numbers)
        col_data = self._ensure_numeric_column(col_data, request.column)

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

        # Find all columns using normalized matching
        actual_columns = self._find_columns(df, request.columns, context="correlate")

        # Apply filters if provided
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Select only requested columns
        df_subset = df[actual_columns].copy()

        # Ensure all columns are numeric (with auto-conversion for text-stored numbers)
        for col in actual_columns:
            df_subset[col] = self._ensure_numeric_column(df_subset[col], col)

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

        # Find column using normalized matching
        actual_column = self._find_column(df, request.column, context="detect_outliers")

        # Get column data
        col_data = df[actual_column].copy()

        # Ensure column is numeric (with auto-conversion for text-stored numbers)
        col_data = self._ensure_numeric_column(col_data, request.column)

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

        response = DetectOutliersResponse(
            outliers=outlier_rows,
            outlier_count=len(outlier_rows),
            method=request.method,
            threshold=request.threshold,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

        # CONTEXT OVERFLOW PROTECTION: Validate response size
        self._validate_response_size(
            response,
            rows_count=len(outlier_rows),
            columns_count=len(df.columns)
        )

        return response
