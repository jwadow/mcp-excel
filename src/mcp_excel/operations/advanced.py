# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Advanced operations for Excel data."""

import re
import time
from typing import Any, Optional

import pandas as pd
import psutil

from ..core.file_loader import FileLoader
from ..core.header_detector import HeaderDetector
from ..excel.formula_generator import FormulaGenerator
from ..excel.tsv_formatter import TSVFormatter
from ..models.requests import (
    CalculateExpressionRequest,
    RankRowsRequest,
)
from ..models.responses import (
    CalculateExpressionResponse,
    ExcelOutput,
    FileMetadata,
    PerformanceMetrics,
    RankRowsResponse,
)
from .filtering import FilterEngine


class AdvancedOperations:
    """Advanced operations for data analysis."""

    def __init__(self, file_loader: FileLoader):
        """Initialize advanced operations.

        Args:
            file_loader: FileLoader instance for loading Excel files
        """
        self._loader = file_loader
        self._header_detector = HeaderDetector()
        self._filter_engine = FilterEngine()
        self._tsv_formatter = TSVFormatter()

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

    def rank_rows(self, request: RankRowsRequest) -> RankRowsResponse:
        """Rank rows by column value.

        Args:
            request: Request parameters

        Returns:
            RankRowsResponse with ranked rows
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate rank column
        if request.rank_column not in df.columns:
            raise ValueError(
                f"Rank column '{request.rank_column}' not found. "
                f"Available columns: {list(df.columns)}"
            )

        # Apply filters
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Convert rank column to numeric
        df[request.rank_column] = pd.to_numeric(df[request.rank_column], errors='coerce')

        # Calculate ranks
        ascending = request.direction == "asc"
        
        if request.group_by_columns:
            # Rank within groups
            df['rank'] = df.groupby(request.group_by_columns)[request.rank_column].rank(
                ascending=ascending, method='min'
            )
        else:
            # Global ranking
            df['rank'] = df[request.rank_column].rank(ascending=ascending, method='min')

        # Sort by rank
        df = df.sort_values(by='rank')

        # Apply top_n limit if specified
        if request.top_n is not None:
            df = df.head(request.top_n)

        # Format results
        result_columns = ['rank'] + list(df.columns[df.columns != 'rank'])
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
        # RANK(value, range, order) where order: 0=desc, 1=asc
        order = 1 if ascending else 0
        formula = f"=RANK(B2,$B$2:$B$100,{order})"

        return RankRowsResponse(
            rows=rows,
            rank_column=request.rank_column,
            direction=request.direction,
            total_rows=len(rows),
            group_by_columns=request.group_by_columns,
            excel_output=ExcelOutput(tsv=tsv, formula=formula),
            metadata=self._get_file_metadata(request.file_path, request.sheet_name, df),
            performance=self._get_performance_metrics(start_time, len(df), False),
        )

    def calculate_expression(
        self, request: CalculateExpressionRequest
    ) -> CalculateExpressionResponse:
        """Calculate expression between columns.

        Args:
            request: Request parameters

        Returns:
            CalculateExpressionResponse with calculated values
        """
        start_time = time.time()

        # Load data
        df, header_row = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Apply filters
        if request.filters:
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Parse expression to find column names
        # Support column names (including those with spaces)
        column_pattern = r'\b[A-Za-z_][A-Za-z0-9_\s]*\b'
        potential_columns = re.findall(column_pattern, request.expression)
        
        # Filter to only actual column names
        used_columns = [col for col in potential_columns if col in df.columns]
        
        if not used_columns:
            raise ValueError(
                f"No valid column names found in expression '{request.expression}'. "
                f"Available columns: {list(df.columns)}"
            )

        # Validate all used columns exist
        for col in used_columns:
            if col not in df.columns:
                raise ValueError(
                    f"Column '{col}' from expression not found. "
                    f"Available columns: {list(df.columns)}"
                )

        # Convert columns to numeric
        for col in used_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # Build safe expression for pandas eval
        # Replace column names with df['column_name'] syntax
        safe_expr = request.expression
        for col in sorted(used_columns, key=len, reverse=True):  # Sort by length to avoid partial replacements
            safe_expr = safe_expr.replace(col, f"df['{col}']")

        try:
            # Evaluate expression using pandas eval (safe for arithmetic)
            df[request.output_column_name] = eval(safe_expr)
        except Exception as e:
            raise ValueError(
                f"Failed to evaluate expression '{request.expression}': {str(e)}"
            )

        # Format results
        result_columns = list(df.columns)
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
        # Convert expression to Excel formula syntax
        excel_formula = request.expression
        for col in used_columns:
            # Find column index (assuming columns are A, B, C, ...)
            if col in df.columns:
                col_idx = list(df.columns).index(col)
                col_letter = chr(65 + col_idx)  # A=65, B=66, etc.
                excel_formula = excel_formula.replace(col, f"{col_letter}2")
        
        formula = f"={excel_formula}"

        return CalculateExpressionResponse(
            rows=rows,
            expression=request.expression,
            output_column_name=request.output_column_name,
            excel_output=ExcelOutput(tsv=tsv, formula=formula),
            metadata=self._get_file_metadata(request.file_path, request.sheet_name, df),
            performance=self._get_performance_metrics(start_time, len(df), False),
        )
