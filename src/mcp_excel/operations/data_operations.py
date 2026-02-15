# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Data operations for Excel files - filtering, aggregation, and data retrieval."""

import time

import pandas as pd

from ..core.file_loader import FileLoader
from ..excel.formula_generator import FormulaGenerator
from ..excel.tsv_formatter import TSVFormatter
from ..models.requests import (
    AggregateRequest,
    FilterAndCountRequest,
    FilterAndGetRowsRequest,
    GetUniqueValuesRequest,
    GetValueCountsRequest,
    GroupByRequest,
)
from ..models.responses import (
    AggregateResponse,
    ExcelOutput,
    FilterAndCountResponse,
    FilterAndGetRowsResponse,
    GetUniqueValuesResponse,
    GetValueCountsResponse,
    GroupByResponse,
)
from ..operations.base import BaseOperations
from ..operations.filtering import FilterEngine


class DataOperations(BaseOperations):
    """Operations for retrieving and filtering Excel data."""

    def __init__(self, file_loader: FileLoader) -> None:
        """Initialize data operations.

        Args:
            file_loader: FileLoader instance for loading files
        """
        super().__init__(file_loader)
        self._filter_engine = FilterEngine()
        self._tsv_formatter = TSVFormatter()
    
    def _get_column_types(self, df: pd.DataFrame) -> dict[str, str]:
        """Get types of all columns in DataFrame.
        
        Args:
            df: DataFrame to analyze
            
        Returns:
            Dictionary mapping column names to type strings
        """
        column_types = {}
        
        for col in df.columns:
            col_str = str(col)
            
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                column_types[col_str] = "datetime"
            elif pd.api.types.is_integer_dtype(df[col]):
                column_types[col_str] = "integer"
            elif pd.api.types.is_float_dtype(df[col]):
                column_types[col_str] = "float"
            elif pd.api.types.is_bool_dtype(df[col]):
                column_types[col_str] = "boolean"
            else:
                column_types[col_str] = "string"
        
        return column_types

    def get_unique_values(
        self, request: GetUniqueValuesRequest
    ) -> GetUniqueValuesResponse:
        """Get unique values from a column.

        Args:
            request: Unique values request

        Returns:
            Unique values response
        """
        start_time = time.time()

        # Load DataFrame
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Check column exists
        if request.column not in df.columns:
            available = ", ".join(str(col) for col in df.columns)
            raise ValueError(
                f"Column '{request.column}' not found. Available columns: {available}"
            )

        # Get unique values
        unique_vals = df[request.column].dropna().unique()
        
        # Sort for consistency (handle mixed types)
        try:
            unique_vals = sorted(unique_vals)
        except TypeError:
            # Mixed types - convert to string for sorting
            unique_vals = sorted(unique_vals, key=str)

        # Apply limit
        truncated = len(unique_vals) > request.limit
        values = [self._format_value(v) for v in unique_vals[: request.limit]]

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        return GetUniqueValuesResponse(
            values=values,
            count=len(values),
            truncated=truncated,
            metadata=metadata,
            performance=performance,
        )

    def get_value_counts(
        self, request: GetValueCountsRequest
    ) -> GetValueCountsResponse:
        """Get value counts (frequency) from a column.

        Args:
            request: Value counts request

        Returns:
            Value counts response
        """
        start_time = time.time()

        # Load DataFrame
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Check column exists
        if request.column not in df.columns:
            available = ", ".join(str(col) for col in df.columns)
            raise ValueError(
                f"Column '{request.column}' not found. Available columns: {available}"
            )

        # Get value counts
        value_counts = df[request.column].value_counts().head(request.top_n)
        # Format keys to remove .0 from floats, then convert to string
        value_counts_dict = {str(self._format_value(k)): int(v) for k, v in value_counts.items()}
        total_values = int(df[request.column].count())

        # Generate TSV output
        headers = [request.column, "Count"]
        rows = [[k, v] for k, v in value_counts_dict.items()]
        tsv = self._tsv_formatter.format_table(headers, rows)

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        return GetValueCountsResponse(
            value_counts=value_counts_dict,
            total_values=total_values,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

    def filter_and_count(
        self, request: FilterAndCountRequest
    ) -> FilterAndCountResponse:
        """Count rows matching filter conditions.

        Args:
            request: Filter and count request

        Returns:
            Filter and count response
        """
        start_time = time.time()

        # Load DataFrame
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate filters
        is_valid, error_msg = self._filter_engine.validate_filters(df, request.filters)
        if not is_valid:
            raise ValueError(error_msg)

        # Count filtered rows
        count = self._filter_engine.count_filtered(df, request.filters, request.logic)

        # Generate Excel formula
        formula_gen = FormulaGenerator(request.sheet_name)
        
        # Get column types for datetime handling in formulas
        column_types = self._get_column_types(df)
        
        # Build column ranges for formula generation
        column_indices = {str(col): idx for idx, col in enumerate(df.columns)}
        column_ranges = {}
        for filter_cond in request.filters:
            col_idx = column_indices.get(filter_cond.column)
            if col_idx is not None:
                column_ranges[filter_cond.column] = formula_gen._get_column_range(
                    filter_cond.column, col_idx
                )

        # Generate formula (returns None if filters use operators not supported in Excel)
        formula = formula_gen.generate_from_filter(
            operation="count",
            filters=request.filters,
            column_ranges=column_ranges,
            column_types=column_types,
        )

        # Generate TSV output
        filter_summary = self._filter_engine.get_filter_summary(request.filters, request.logic)
        tsv = self._tsv_formatter.format_single_value(
            "Count", count, formula
        )

        excel_output = ExcelOutput(
            tsv=tsv,
            formula=formula,
            references=formula_gen.get_references(
                list(column_ranges.keys()), column_indices
            ) if formula else None,
        )

        # Serialize filters for response
        filters_applied = [
            {
                "column": f.column,
                "operator": f.operator,
                "value": f.value,
                "values": f.values,
            }
            for f in request.filters
        ]

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        return FilterAndCountResponse(
            count=count,
            filters_applied=filters_applied,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

    def filter_and_get_rows(
        self, request: FilterAndGetRowsRequest
    ) -> FilterAndGetRowsResponse:
        """Get rows matching filter conditions.

        Args:
            request: Filter and get rows request

        Returns:
            Filter and get rows response
        """
        start_time = time.time()

        # Load DataFrame
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate filters
        is_valid, error_msg = self._filter_engine.validate_filters(df, request.filters)
        if not is_valid:
            raise ValueError(error_msg)

        # Apply filters
        filtered_df = self._filter_engine.apply_filters(df, request.filters, request.logic)
        total_matches = len(filtered_df)

        # CONTEXT OVERFLOW PROTECTION: Apply smart column limit
        if request.columns:
            # Validate requested columns
            missing_cols = set(request.columns) - set(df.columns)
            if missing_cols:
                available = ", ".join(str(col) for col in df.columns)
                raise ValueError(
                    f"Columns not found: {', '.join(missing_cols)}. Available: {available}"
                )
            filtered_df = filtered_df[request.columns]
            actual_columns = request.columns
        else:
            # Apply default column limit to prevent context overflow
            filtered_df, actual_columns = self._apply_column_limit(filtered_df, None)

        # CONTEXT OVERFLOW PROTECTION: Enforce row limit
        enforced_limit = self._enforce_row_limit(request.limit)
        
        # Apply limit and offset
        result_df = filtered_df.iloc[request.offset : request.offset + enforced_limit]
        truncated = total_matches > (request.offset + enforced_limit)

        # Convert to list of dicts
        rows = []
        for idx in range(len(result_df)):
            row_dict = result_df.iloc[idx].to_dict()
            # Convert to JSON-serializable types with string keys and format values
            row_dict = {str(k): self._format_value(v) for k, v in row_dict.items()}
            rows.append(row_dict)

        # Generate TSV output
        if rows:
            headers = list(rows[0].keys())
            data_rows = [[row[h] for h in headers] for row in rows]
            tsv = self._tsv_formatter.format_table(headers, data_rows)
        else:
            tsv = "No rows match the filters"

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        response = FilterAndGetRowsResponse(
            rows=rows,
            count=len(rows),
            total_matches=total_matches,
            truncated=truncated,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

        # CONTEXT OVERFLOW PROTECTION: Validate response size
        self._validate_response_size(
            response,
            rows_count=len(rows),
            columns_count=len(actual_columns),
            request_limit=request.limit
        )

        return response

    def aggregate(self, request: AggregateRequest) -> AggregateResponse:
        """Perform aggregation on a column.

        Args:
            request: Aggregation request

        Returns:
            Aggregation response
        """
        start_time = time.time()

        # Load DataFrame
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Check target column exists
        if request.target_column not in df.columns:
            available = ", ".join(str(col) for col in df.columns)
            raise ValueError(
                f"Column '{request.target_column}' not found. Available columns: {available}"
            )

        # Apply filters if provided
        if request.filters:
            is_valid, error_msg = self._filter_engine.validate_filters(df, request.filters)
            if not is_valid:
                raise ValueError(error_msg)
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Get column data
        col_data = df[request.target_column]

        # Try to convert to numeric if it's not already numeric
        # This handles Excel's common issue of storing numbers as text
        if col_data.dtype == 'object' or col_data.dtype.name == 'string':
            col_data_numeric = pd.to_numeric(col_data, errors='coerce')
            # Check if conversion was successful for most values
            non_null_original = col_data.notna().sum()
            non_null_converted = col_data_numeric.notna().sum()
            
            # If we lost less than 50% of data in conversion, use numeric version
            if non_null_converted >= non_null_original * 0.5:
                col_data = col_data_numeric
        
        # Drop NaN for aggregation
        col_data = col_data.dropna()

        # Perform aggregation
        operation = request.operation
        try:
            if operation == "sum":
                result = float(col_data.sum())
            elif operation == "mean":
                result = float(col_data.mean())
            elif operation == "median":
                result = float(col_data.median())
            elif operation == "min":
                result = float(col_data.min())
            elif operation == "max":
                result = float(col_data.max())
            elif operation == "std":
                result = float(col_data.std())
            elif operation == "var":
                result = float(col_data.var())
            elif operation == "count":
                result = len(col_data)
            else:
                raise ValueError(f"Unsupported operation: {operation}")
        except (TypeError, ValueError) as e:
            raise ValueError(
                f"Cannot perform '{operation}' on column '{request.target_column}'. "
                f"Column contains non-numeric data that cannot be converted. Error: {e}"
            )

        # Generate Excel formula
        formula_gen = FormulaGenerator(request.sheet_name)
        column_indices = {str(col): idx for idx, col in enumerate(df.columns)}
        
        # Get column types for datetime handling in formulas
        column_types = self._get_column_types(df)
        
        # Build column ranges
        column_ranges = {}
        if request.filters:
            for filter_cond in request.filters:
                col_idx = column_indices.get(filter_cond.column)
                if col_idx is not None:
                    column_ranges[filter_cond.column] = formula_gen._get_column_range(
                        filter_cond.column, col_idx
                    )

        # Get target column range
        target_col_idx = column_indices.get(request.target_column)
        if target_col_idx is not None:
            target_range = formula_gen._get_column_range(request.target_column, target_col_idx)
        else:
            target_range = None

        # Generate formula (returns None if filters use operators not supported in Excel)
        formula = formula_gen.generate_from_filter(
            operation=operation,
            filters=request.filters,
            column_ranges=column_ranges,
            target_range=target_range,
            column_types=column_types,
        )

        # Generate TSV output
        tsv = self._tsv_formatter.format_single_value(
            f"{operation.capitalize()} of {request.target_column}",
            result,
            formula
        )

        excel_output = ExcelOutput(
            tsv=tsv,
            formula=formula,
            references=formula_gen.get_references(
                list(column_ranges.keys()) + [request.target_column], column_indices
            ) if formula else None,
        )

        # Serialize filters for response
        filters_applied = [
            {
                "column": f.column,
                "operator": f.operator,
                "value": f.value,
                "values": f.values,
            }
            for f in request.filters
        ]

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        return AggregateResponse(
            value=self._format_value(result),
            operation=operation,
            target_column=request.target_column,
            filters_applied=filters_applied,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

    def group_by(self, request: GroupByRequest) -> GroupByResponse:
        """Perform group-by aggregation.

        Args:
            request: Group-by request

        Returns:
            Group-by response
        """
        start_time = time.time()

        # Load DataFrame
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate columns
        all_columns = request.group_columns + [request.agg_column]
        missing_cols = set(all_columns) - set(df.columns)
        if missing_cols:
            available = ", ".join(str(col) for col in df.columns)
            raise ValueError(
                f"Columns not found: {', '.join(missing_cols)}. Available: {available}"
            )

        # Apply filters if provided
        if request.filters:
            is_valid, error_msg = self._filter_engine.validate_filters(df, request.filters)
            if not is_valid:
                raise ValueError(error_msg)
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Perform group-by aggregation
        operation = request.agg_operation
        
        # Try to convert aggregation column to numeric if needed (only for numeric operations)
        if operation != "count":
            if df[request.agg_column].dtype == 'object' or df[request.agg_column].dtype.name == 'string':
                df[request.agg_column] = pd.to_numeric(df[request.agg_column], errors='coerce')
        try:
            if operation == "sum":
                grouped = df.groupby(request.group_columns)[request.agg_column].sum()
            elif operation == "mean":
                grouped = df.groupby(request.group_columns)[request.agg_column].mean()
            elif operation == "median":
                grouped = df.groupby(request.group_columns)[request.agg_column].median()
            elif operation == "min":
                grouped = df.groupby(request.group_columns)[request.agg_column].min()
            elif operation == "max":
                grouped = df.groupby(request.group_columns)[request.agg_column].max()
            elif operation == "std":
                grouped = df.groupby(request.group_columns)[request.agg_column].std()
            elif operation == "var":
                grouped = df.groupby(request.group_columns)[request.agg_column].var()
            elif operation == "count":
                grouped = df.groupby(request.group_columns)[request.agg_column].count()
            else:
                raise ValueError(f"Unsupported operation: {operation}")
        except (TypeError, ValueError, KeyError) as e:
            raise ValueError(
                f"Cannot perform '{operation}' on column '{request.agg_column}' "
                f"grouped by {request.group_columns}. Column may contain non-numeric data. Error: {e}"
            )

        # Convert to list of dicts
        # Use reset_index with name parameter to avoid column name conflicts
        if isinstance(grouped, pd.Series):
            # For Series, specify the name for the aggregation column
            agg_col_name = f"{request.agg_column}_{operation}"
            grouped_df = grouped.reset_index(name=agg_col_name)
        else:
            # For DataFrame (shouldn't happen with single agg column, but handle it)
            grouped_df = grouped.reset_index()
        
        groups = []
        for idx in range(len(grouped_df)):
            row_dict = grouped_df.iloc[idx].to_dict()
            # Convert to JSON-serializable types with string keys and format values
            formatted_dict = {str(k): self._format_value(v) for k, v in row_dict.items()}
            groups.append(formatted_dict)

        # Generate TSV output
        if groups:
            headers = list(groups[0].keys())
            data_rows = [[row[h] for h in headers] for row in groups]
            tsv = self._tsv_formatter.format_table(headers, data_rows)
        else:
            tsv = "No groups found"

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        response = GroupByResponse(
            groups=groups,
            group_columns=request.group_columns,
            agg_column=request.agg_column,
            agg_operation=operation,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

        # CONTEXT OVERFLOW PROTECTION: Validate response size
        self._validate_response_size(
            response,
            rows_count=len(groups),
            columns_count=len(grouped_df.columns)
        )

        return response
