# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Advanced operations for Excel data."""

import time
from typing import Optional

import pandas as pd

from ..core.file_loader import FileLoader
from ..excel.formula_generator import FormulaGenerator
from ..excel.tsv_formatter import TSVFormatter
from ..models.requests import (
    CalculateExpressionRequest,
    RankRowsRequest,
)
from ..models.responses import (
    CalculateExpressionResponse,
    ExcelOutput,
    RankRowsResponse,
)
from ..operations.base import BaseOperations
from ..operations.filtering import FilterEngine


class AdvancedOperations(BaseOperations):
    """Advanced operations for data analysis."""

    def __init__(self, file_loader: FileLoader):
        """Initialize advanced operations.

        Args:
            file_loader: FileLoader instance for loading Excel files
        """
        super().__init__(file_loader)
        self._filter_engine = FilterEngine()
        self._tsv_formatter = TSVFormatter()

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

        # Generate Excel formula using FormulaGenerator
        formula_gen = FormulaGenerator(request.sheet_name)
        column_indices = {str(col): idx for idx, col in enumerate(df.columns)}
        
        # Find rank column index
        rank_col_idx = column_indices.get(request.rank_column)
        if rank_col_idx is not None:
            col_letter = formula_gen._column_letter(rank_col_idx)
            # RANK(value, range, order) where order: 0=desc, 1=asc
            order = 1 if ascending else 0
            # Use actual row count for range
            last_row = len(df) + 1  # +1 because Excel is 1-based and we have header
            formula = f"=RANK({col_letter}2,${col_letter}$2:${col_letter}${last_row},{order})"
        else:
            formula = None

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        response = RankRowsResponse(
            rows=rows,
            rank_column=request.rank_column,
            direction=request.direction,
            total_rows=len(rows),
            group_by_columns=request.group_by_columns,
            excel_output=ExcelOutput(tsv=tsv, formula=formula),
            metadata=metadata,
            performance=self._get_performance_metrics(start_time, len(df), False),
        )

        # CONTEXT OVERFLOW PROTECTION: Validate response size
        self._validate_response_size(
            response,
            rows_count=len(rows),
            columns_count=len(result_columns)
        )

        return response

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

        # Find which DataFrame columns are used in the expression
        # Sort by length (longest first) to avoid partial matches
        # Example: "Дата прибытия" must be found before "Дата"
        # This approach is language-agnostic and works with any Unicode characters
        sorted_columns = sorted(df.columns, key=len, reverse=True)
        used_columns = [col for col in sorted_columns if col in request.expression]
        
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

        # Build safe expression for pandas.eval()
        # Backtick-quote column names to handle spaces and special chars
        safe_expr = request.expression
        for col in sorted(used_columns, key=len, reverse=True):  # Sort by length to avoid partial replacements
            safe_expr = safe_expr.replace(col, f"`{col}`")

        try:
            # Use pandas.eval() which is safe for arithmetic expressions
            df[request.output_column_name] = df.eval(safe_expr)
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

        # Generate Excel formula using FormulaGenerator
        formula_gen = FormulaGenerator(request.sheet_name)
        column_indices = {str(col): idx for idx, col in enumerate(df.columns)}
        
        # Convert expression to Excel formula syntax
        excel_formula = request.expression
        for col in sorted(used_columns, key=len, reverse=True):
            if col in column_indices:
                col_idx = column_indices[col]
                col_letter = formula_gen._column_letter(col_idx)
                excel_formula = excel_formula.replace(col, f"{col_letter}2")
        
        formula = f"={excel_formula}"

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        response = CalculateExpressionResponse(
            rows=rows,
            expression=request.expression,
            output_column_name=request.output_column_name,
            excel_output=ExcelOutput(tsv=tsv, formula=formula),
            metadata=metadata,
            performance=self._get_performance_metrics(start_time, len(df), False),
        )

        # CONTEXT OVERFLOW PROTECTION: Validate response size
        self._validate_response_size(
            response,
            rows_count=len(rows),
            columns_count=len(result_columns)
        )

        return response
