# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Time series operations for Excel data."""

import time
from typing import Any, Optional

import pandas as pd
import psutil

from ..core.file_loader import FileLoader
from ..core.header_detector import HeaderDetector
from ..excel.formula_generator import FormulaGenerator
from ..excel.tsv_formatter import TSVFormatter
from ..models.requests import (
    CalculateMovingAverageRequest,
    CalculatePeriodChangeRequest,
    CalculateRunningTotalRequest,
)
from ..models.responses import (
    CalculateMovingAverageResponse,
    CalculatePeriodChangeResponse,
    CalculateRunningTotalResponse,
    ExcelOutput,
    FileMetadata,
    PerformanceMetrics,
)
from .filtering import FilterEngine


class TimeSeriesOperations:
    """Operations for time series analysis."""

    def __init__(self, file_loader: FileLoader):
        """Initialize time series operations.

        Args:
            file_loader: FileLoader instance for loading Excel files
        """
        self._loader = file_loader
        self._header_detector = HeaderDetector()
        self._filter_engine = FilterEngine()
        self._tsv_formatter = TSVFormatter()

    def _column_letter(self, col_index: int) -> str:
        """Convert column index to Excel letter (supports AA, AB, etc).

        Args:
            col_index: Zero-based column index

        Returns:
            Excel column letter (A, B, ..., Z, AA, AB, ...)
        """
        result = ""
        col_index += 1  # Excel is 1-based
        while col_index > 0:
            col_index -= 1
            result = chr(65 + (col_index % 26)) + result
            col_index //= 26
        return result

    def _format_value(self, value: Any) -> Any:
        """Format value for output (convert float to int if whole number).

        Args:
            value: Value to format

        Returns:
            Formatted value
        """
        if pd.isna(value):
            return None
        elif isinstance(value, float) and value.is_integer():
            return int(value)
        else:
            return value

    def _get_performance_metrics(
        self, start_time: float, rows_processed: int, cache_hit: bool
    ) -> PerformanceMetrics:
        """Get performance metrics for operation.

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
        self, file_path: str, sheet_name: str, df: pd.DataFrame
    ) -> FileMetadata:
        """Get file metadata.

        Args:
            file_path: Path to file
            sheet_name: Sheet name
            df: DataFrame

        Returns:
            FileMetadata object
        """
        file_info = self._loader.get_file_info(file_path)

        return FileMetadata(
            file_format=file_info["format"],
            sheet_name=sheet_name,
            rows_total=len(df),
            columns_total=len(df.columns),
        )

    def _load_with_header_detection(
        self, file_path: str, sheet_name: str, header_row: Optional[int]
    ) -> tuple[pd.DataFrame, int]:
        """Load file with automatic header detection.

        Args:
            file_path: Path to file
            sheet_name: Sheet name
            header_row: Optional header row index

        Returns:
            Tuple of (DataFrame, header_row_index)
        """
        if header_row is not None:
            df = self._loader.load(file_path, sheet_name, header_row=header_row, use_cache=True)
            return df, header_row

        # Auto-detect header
        df_raw = self._loader.load(file_path, sheet_name, header_row=None, use_cache=True)
        detection_result = self._header_detector.detect(df_raw)
        df = self._loader.load(
            file_path, sheet_name, header_row=detection_result.header_row, use_cache=True
        )

        return df, detection_result.header_row

    def calculate_period_change(
        self, request: CalculatePeriodChangeRequest
    ) -> CalculatePeriodChangeResponse:
        """Calculate period-over-period change.

        Args:
            request: Request parameters

        Returns:
            CalculatePeriodChangeResponse with period changes
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate columns
        if request.date_column not in df.columns:
            raise ValueError(
                f"Date column '{request.date_column}' not found. "
                f"Available columns: {list(df.columns)}"
            )
        if request.value_column not in df.columns:
            raise ValueError(
                f"Value column '{request.value_column}' not found. "
                f"Available columns: {list(df.columns)}"
            )

        # Apply filters
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Ensure date column is datetime
        if not pd.api.types.is_datetime64_any_dtype(df[request.date_column]):
            df[request.date_column] = pd.to_datetime(df[request.date_column], errors='coerce')

        # Convert value column to numeric
        value_col = pd.to_numeric(df[request.value_column], errors='coerce')

        # Group by period
        if request.period_type == "month":
            df['period'] = df[request.date_column].dt.to_period('M')
        elif request.period_type == "quarter":
            df['period'] = df[request.date_column].dt.to_period('Q')
        elif request.period_type == "year":
            df['period'] = df[request.date_column].dt.to_period('Y')

        # Aggregate by period
        period_data = df.groupby('period')[request.value_column].agg(
            lambda x: pd.to_numeric(x, errors='coerce').sum()
        ).reset_index()
        period_data.columns = ['period', 'value']

        # Calculate changes
        period_data['change_absolute'] = period_data['value'].diff()
        period_data['change_percent'] = (
            period_data['value'].pct_change() * 100
        )

        # Format results
        periods = []
        for _, row in period_data.iterrows():
            periods.append({
                'period': str(row['period']),
                'value': self._format_value(row['value']),
                'change_absolute': self._format_value(row['change_absolute']),
                'change_percent': self._format_value(row['change_percent']),
            })

        # Generate TSV
        headers = ['Period', 'Value', 'Change (Absolute)', 'Change (%)']
        rows = [
            [p['period'], p['value'], p['change_absolute'], p['change_percent']]
            for p in periods
        ]
        tsv = self._tsv_formatter.format_table(headers, rows)

        # Generate Excel formula for percent change
        formula = "=(B2-B1)/B1*100"

        return CalculatePeriodChangeResponse(
            periods=periods,
            period_type=request.period_type,
            value_column=request.value_column,
            excel_output=ExcelOutput(tsv=tsv, formula=formula),
            metadata=self._get_file_metadata(request.file_path, request.sheet_name, df),
            performance=self._get_performance_metrics(start_time, len(df), False),
        )

    def calculate_running_total(
        self, request: CalculateRunningTotalRequest
    ) -> CalculateRunningTotalResponse:
        """Calculate running total (cumulative sum).

        Args:
            request: Request parameters

        Returns:
            CalculateRunningTotalResponse with running totals
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate columns
        if request.order_column not in df.columns:
            raise ValueError(
                f"Order column '{request.order_column}' not found. "
                f"Available columns: {list(df.columns)}"
            )
        if request.value_column not in df.columns:
            raise ValueError(
                f"Value column '{request.value_column}' not found. "
                f"Available columns: {list(df.columns)}"
            )

        # Apply filters
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Convert value column to numeric
        df[request.value_column] = pd.to_numeric(df[request.value_column], errors='coerce')

        # Sort by order column
        df = df.sort_values(by=request.order_column)

        # Calculate running total
        if request.group_by_columns:
            # Running total within groups
            df['running_total'] = df.groupby(request.group_by_columns)[request.value_column].cumsum()
        else:
            # Overall running total
            df['running_total'] = df[request.value_column].cumsum()

        # Format results
        result_columns = [request.order_column, request.value_column, 'running_total']
        if request.group_by_columns:
            result_columns = request.group_by_columns + result_columns

        rows = []
        for _, row in df.iterrows():
            row_dict = {}
            for col in result_columns:
                row_dict[col] = self._format_value(row[col])
            rows.append(row_dict)

        # Generate TSV
        headers = result_columns
        tsv_rows = [[row[col] for col in result_columns] for row in rows]
        tsv = self._tsv_formatter.format_table(headers, tsv_rows)

        # Generate Excel formula
        # Find value_column index in result_columns
        value_col_idx = result_columns.index(request.value_column)
        value_col_letter = self._column_letter(value_col_idx)
        formula = f"=SUM(${value_col_letter}$2:{value_col_letter}2)"

        return CalculateRunningTotalResponse(
            rows=rows,
            order_column=request.order_column,
            value_column=request.value_column,
            group_by_columns=request.group_by_columns,
            excel_output=ExcelOutput(tsv=tsv, formula=formula),
            metadata=self._get_file_metadata(request.file_path, request.sheet_name, df),
            performance=self._get_performance_metrics(start_time, len(df), False),
        )

    def calculate_moving_average(
        self, request: CalculateMovingAverageRequest
    ) -> CalculateMovingAverageResponse:
        """Calculate moving average.

        Args:
            request: Request parameters

        Returns:
            CalculateMovingAverageResponse with moving averages
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate columns
        if request.order_column not in df.columns:
            raise ValueError(
                f"Order column '{request.order_column}' not found. "
                f"Available columns: {list(df.columns)}"
            )
        if request.value_column not in df.columns:
            raise ValueError(
                f"Value column '{request.value_column}' not found. "
                f"Available columns: {list(df.columns)}"
            )

        # Apply filters
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Convert value column to numeric
        df[request.value_column] = pd.to_numeric(df[request.value_column], errors='coerce')

        # Sort by order column
        df = df.sort_values(by=request.order_column)

        # Calculate moving average
        df['moving_average'] = df[request.value_column].rolling(
            window=request.window_size, min_periods=1
        ).mean()

        # Format results
        result_columns = [request.order_column, request.value_column, 'moving_average']
        rows = []
        for _, row in df.iterrows():
            row_dict = {}
            for col in result_columns:
                row_dict[col] = self._format_value(row[col])
            rows.append(row_dict)

        # Generate TSV
        headers = result_columns
        tsv_rows = [[row[col] for col in result_columns] for row in rows]
        tsv = self._tsv_formatter.format_table(headers, tsv_rows)

        # Generate Excel formula
        # Find value_column index in result_columns
        value_col_idx = result_columns.index(request.value_column)
        value_col_letter = self._column_letter(value_col_idx)
        start_row = max(2, 3 - request.window_size)  # Don't go below row 2 (first data row)
        formula = f"=AVERAGE({value_col_letter}{start_row}:{value_col_letter}2)"

        return CalculateMovingAverageResponse(
            rows=rows,
            order_column=request.order_column,
            value_column=request.value_column,
            window_size=request.window_size,
            excel_output=ExcelOutput(tsv=tsv, formula=formula),
            metadata=self._get_file_metadata(request.file_path, request.sheet_name, df),
            performance=self._get_performance_metrics(start_time, len(df), False),
        )
