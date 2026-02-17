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
    AnalyzeOverlapRequest,
    FilterAndCountBatchRequest,
    FilterAndCountRequest,
    FilterAndGetRowsRequest,
    GetUniqueValuesRequest,
    GetValueCountsRequest,
    GroupByRequest,
    FilterCondition,
    FilterGroup,
)
from ..models.responses import (
    AggregateResponse,
    AnalyzeOverlapResponse,
    ExcelOutput,
    FilterAndCountBatchResponse,
    FilterAndCountResponse,
    FilterAndGetRowsResponse,
    FilterSetResult,
    GetUniqueValuesResponse,
    GetValueCountsResponse,
    GroupByResponse,
    SetInfo,
    VennDiagram2,
    VennDiagram3,
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
    
    def _serialize_filters(self, filters: list[FilterCondition | FilterGroup]) -> list[dict]:
        """Serialize filters for response (recursive).
        
        Handles both FilterCondition and FilterGroup (nested).
        
        Args:
            filters: List of filter conditions or groups
            
        Returns:
            List of serialized filter dictionaries
        """
        result = []
        
        for filter_item in filters:
            if isinstance(filter_item, FilterCondition):
                # Atomic condition
                result.append({
                    "column": filter_item.column,
                    "operator": filter_item.operator,
                    "value": filter_item.value,
                    "values": filter_item.values,
                    "negate": filter_item.negate
                })
            elif isinstance(filter_item, FilterGroup):
                # Nested group - RECURSION
                result.append({
                    "type": "group",
                    "filters": self._serialize_filters(filter_item.filters),
                    "logic": filter_item.logic,
                    "negate": filter_item.negate
                })
            else:
                raise ValueError(f"Unknown filter type: {type(filter_item)}")
        
        return result
    
    def _build_column_ranges_from_filters(
        self,
        filters: list[FilterCondition | FilterGroup],
        formula_gen,
        column_indices: dict[str, int]
    ) -> dict[str, str]:
        """Build column ranges from filters (recursive).
        
        Extracts all columns used in filters (including nested groups)
        and builds Excel ranges for them.
        
        Args:
            filters: List of filter conditions or groups
            formula_gen: FormulaGenerator instance
            column_indices: Mapping of column names to indices
            
        Returns:
            Dictionary mapping column names to Excel ranges
        """
        column_ranges = {}
        
        for filter_item in filters:
            if isinstance(filter_item, FilterCondition):
                # Atomic condition - add its column
                col_idx = column_indices.get(filter_item.column)
                if col_idx is not None:
                    column_ranges[filter_item.column] = formula_gen._get_column_range(
                        filter_item.column, col_idx
                    )
            elif isinstance(filter_item, FilterGroup):
                # Nested group - RECURSION
                nested_ranges = self._build_column_ranges_from_filters(
                    filter_item.filters, formula_gen, column_indices
                )
                column_ranges.update(nested_ranges)
        
        return column_ranges

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

        # Find column using normalized matching
        actual_column = self._find_column(df, request.column, context="get_unique_values")

        # Get unique values
        unique_vals = df[actual_column].dropna().unique()
        
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

        # Find column using normalized matching
        actual_column = self._find_column(df, request.column, context="get_value_counts")

        # Get value counts
        value_counts = df[actual_column].value_counts().head(request.top_n)
        # Format keys to remove .0 from floats, then convert to string
        value_counts_dict = {str(self._format_value(k)): int(v) for k, v in value_counts.items()}
        total_values = int(df[actual_column].count())

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

        # Get filtered DataFrame for sample_rows
        filtered_df = self._filter_engine.apply_filters(df, request.filters, request.logic) if request.filters else df
        sample_rows_data = self._add_sample_rows(filtered_df, request.sample_rows)

        # Generate Excel formula
        formula_gen = FormulaGenerator(request.sheet_name)
        
        # Get column types for datetime handling in formulas
        column_types = self._get_column_types(df)
        
        # Build column ranges for formula generation
        column_indices = {str(col): idx for idx, col in enumerate(df.columns)}
        column_ranges = self._build_column_ranges_from_filters(
            request.filters, formula_gen, column_indices
        )

        # For count with no filters, we need target_range to generate formula like =COUNTA(A:A)
        # Pick first column as reference for counting rows
        target_range = None
        if not request.filters and len(df.columns) > 0:
            first_col = str(df.columns[0])
            first_col_idx = 0
            target_range = formula_gen._get_column_range(first_col, first_col_idx)

        # Generate formula (returns None if filters use operators not supported in Excel)
        formula = formula_gen.generate_from_filter(
            operation="count",
            filters=request.filters,
            column_ranges=column_ranges,
            target_range=target_range,
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
        filters_applied = self._serialize_filters(request.filters)

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        return FilterAndCountResponse(
            count=count,
            filters_applied=filters_applied,
            excel_output=excel_output,
            sample_rows=sample_rows_data,
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
            # Find requested columns using normalized matching
            actual_columns = self._find_columns(df, request.columns, context="filter_and_get_rows")
            filtered_df = filtered_df[actual_columns]
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

        # Find target column using normalized matching
        actual_target_column = self._find_column(df, request.target_column, context="aggregate")

        # Apply filters if provided
        if request.filters:
            is_valid, error_msg = self._filter_engine.validate_filters(df, request.filters)
            if not is_valid:
                raise ValueError(error_msg)
            df = self._filter_engine.apply_filters(df, request.filters, request.logic)

        # Store filtered df for sample_rows before aggregation
        filtered_df_for_samples = df.copy()

        # Get column data
        col_data = df[actual_target_column]

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
            column_ranges = self._build_column_ranges_from_filters(
                request.filters, formula_gen, column_indices
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
        filters_applied = self._serialize_filters(request.filters)

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        sample_rows_data = self._add_sample_rows(filtered_df_for_samples, request.sample_rows)

        return AggregateResponse(
            value=self._format_value(result),
            operation=operation,
            target_column=request.target_column,
            filters_applied=filters_applied,
            excel_output=excel_output,
            sample_rows=sample_rows_data,
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

        # Find all columns using normalized matching
        actual_group_columns = self._find_columns(df, request.group_columns, context="group_by")
        actual_agg_column = self._find_column(df, request.agg_column, context="group_by")

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
            if df[actual_agg_column].dtype == 'object' or df[actual_agg_column].dtype.name == 'string':
                df[actual_agg_column] = pd.to_numeric(df[actual_agg_column], errors='coerce')
        try:
            if operation == "sum":
                grouped = df.groupby(actual_group_columns)[actual_agg_column].sum()
            elif operation == "mean":
                grouped = df.groupby(actual_group_columns)[actual_agg_column].mean()
            elif operation == "median":
                grouped = df.groupby(actual_group_columns)[actual_agg_column].median()
            elif operation == "min":
                grouped = df.groupby(actual_group_columns)[actual_agg_column].min()
            elif operation == "max":
                grouped = df.groupby(actual_group_columns)[actual_agg_column].max()
            elif operation == "std":
                grouped = df.groupby(actual_group_columns)[actual_agg_column].std()
            elif operation == "var":
                grouped = df.groupby(actual_group_columns)[actual_agg_column].var()
            elif operation == "count":
                grouped = df.groupby(actual_group_columns)[actual_agg_column].count()
            else:
                raise ValueError(f"Unsupported operation: {operation}")
        except (TypeError, ValueError, KeyError) as e:
            raise ValueError(
                f"Cannot perform '{operation}' on column '{actual_agg_column}' "
                f"grouped by {actual_group_columns}. Column may contain non-numeric data. Error: {e}"
            )

        # Convert to list of dicts
        # Use reset_index with name parameter to avoid column name conflicts
        if isinstance(grouped, pd.Series):
            # For Series, specify the name for the aggregation column
            agg_col_name = f"{actual_agg_column}_{operation}"
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

    def filter_and_count_batch(
        self, request: FilterAndCountBatchRequest
    ) -> FilterAndCountBatchResponse:
        """Count rows for multiple filter sets in a single call.

        Args:
            request: Batch filter and count request

        Returns:
            Batch filter and count response
        """
        start_time = time.time()

        # Load DataFrame ONCE
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Validate ALL filter sets BEFORE execution
        for i, filter_set in enumerate(request.filter_sets):
            is_valid, error_msg = self._filter_engine.validate_filters(
                df, filter_set.filters
            )
            if not is_valid:
                label = filter_set.label or f"Set {i+1}"
                raise ValueError(f"Filter set '{label}': {error_msg}")

        # Execute all filter sets
        results = []
        for i, filter_set in enumerate(request.filter_sets):
            # Count filtered rows
            count = self._filter_engine.count_filtered(
                df, filter_set.filters, filter_set.logic
            )

            # Get filtered DataFrame for sample_rows
            filtered_df = self._filter_engine.apply_filters(df, filter_set.filters, filter_set.logic)
            sample_rows_data = self._add_sample_rows(filtered_df, filter_set.sample_rows)

            # Serialize filters for response
            filters_applied = self._serialize_filters(filter_set.filters)

            # Generate Excel formula
            formula_gen = FormulaGenerator(request.sheet_name)
            column_types = self._get_column_types(df)
            column_indices = {str(col): idx for idx, col in enumerate(df.columns)}

            column_ranges = self._build_column_ranges_from_filters(
                filter_set.filters, formula_gen, column_indices
            )

            formula = formula_gen.generate_from_filter(
                operation="count",
                filters=filter_set.filters,
                column_ranges=column_ranges,
                column_types=column_types,
            )

            results.append(FilterSetResult(
                label=filter_set.label,
                count=count,
                filters_applied=filters_applied,
                formula=formula,
                sample_rows=sample_rows_data
            ))

        # Generate TSV output
        headers = ["Label", "Count", "Formula"]
        rows = [
            [r.label or f"Set {i+1}", r.count, r.formula or ""]
            for i, r in enumerate(results)
        ]
        tsv = self._tsv_formatter.format_table(headers, rows)

        excel_output = ExcelOutput(
            tsv=tsv,
            formula=None,
            references=None
        )

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        response = FilterAndCountBatchResponse(
            results=results,
            total_filter_sets=len(results),
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

        # CONTEXT OVERFLOW PROTECTION: Validate response size
        self._validate_response_size(
            response,
            rows_count=len(results),
            columns_count=3
        )

        return response

    def analyze_overlap(
        self, request: AnalyzeOverlapRequest
    ) -> AnalyzeOverlapResponse:
        """Analyze overlap between multiple filter sets (Venn diagram analysis).

        Args:
            request: Analyze overlap request

        Returns:
            Analyze overlap response
        """
        start_time = time.time()

        # Load DataFrame ONCE
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        total_rows = len(df)

        # Validate ALL filter sets BEFORE execution
        for i, filter_set in enumerate(request.filter_sets):
            is_valid, error_msg = self._filter_engine.validate_filters(
                df, filter_set.filters
            )
            if not is_valid:
                label = filter_set.label or f"Set {i+1}"
                raise ValueError(f"Filter set '{label}': {error_msg}")

        # Build masks for each filter set
        masks = []
        labels = []
        counts = []

        for i, filter_set in enumerate(request.filter_sets):
            label = filter_set.label or f"Set {i+1}"
            labels.append(label)

            # Build mask using FilterEngine public API
            if filter_set.filters:
                # Apply filters to get filtered DataFrame
                filtered_df = self._filter_engine.apply_filters(df, filter_set.filters, filter_set.logic)
                # Create boolean mask by checking which indices are in filtered DataFrame
                mask = df.index.isin(filtered_df.index)
            else:
                # Empty filter set - all rows
                mask = pd.Series([True] * len(df), index=df.index)

            masks.append(mask)
            counts.append(int(mask.sum()))

        # Calculate union
        union_mask = masks[0]
        for mask in masks[1:]:
            union_mask = union_mask | mask
        union_count = int(union_mask.sum())
        union_percentage = (union_count / total_rows * 100) if total_rows > 0 else 0.0

        # Build set info
        sets_info = {}
        for label, count in zip(labels, counts):
            percentage = (count / total_rows * 100) if total_rows > 0 else 0.0
            sets_info[label] = SetInfo(
                label=label,
                count=count,
                percentage=round(percentage, 2)
            )

        # Calculate pairwise intersections
        pairwise_intersections = {}
        n_sets = len(masks)

        for i in range(n_sets):
            for j in range(i + 1, n_sets):
                intersection_mask = masks[i] & masks[j]
                intersection_count = int(intersection_mask.sum())
                key = f"{labels[i]} ∩ {labels[j]}"
                pairwise_intersections[key] = intersection_count

        # Calculate Venn diagrams for 2 or 3 sets
        venn_diagram_2 = None
        venn_diagram_3 = None

        if n_sets == 2:
            # Full Venn diagram for 2 sets
            intersection_AB = int((masks[0] & masks[1]).sum())
            A_only = counts[0] - intersection_AB
            B_only = counts[1] - intersection_AB

            venn_diagram_2 = VennDiagram2(
                A_only=A_only,
                B_only=B_only,
                A_and_B=intersection_AB
            )

        elif n_sets == 3:
            # Full Venn diagram for 3 sets (7 zones)
            intersection_AB = int((masks[0] & masks[1]).sum())
            intersection_AC = int((masks[0] & masks[2]).sum())
            intersection_BC = int((masks[1] & masks[2]).sum())
            intersection_ABC = int((masks[0] & masks[1] & masks[2]).sum())

            # Calculate exclusive zones
            A_and_B_only = intersection_AB - intersection_ABC
            A_and_C_only = intersection_AC - intersection_ABC
            B_and_C_only = intersection_BC - intersection_ABC

            A_only = counts[0] - intersection_AB - intersection_AC + intersection_ABC
            B_only = counts[1] - intersection_AB - intersection_BC + intersection_ABC
            C_only = counts[2] - intersection_AC - intersection_BC + intersection_ABC

            venn_diagram_3 = VennDiagram3(
                A_only=A_only,
                B_only=B_only,
                C_only=C_only,
                A_and_B_only=A_and_B_only,
                A_and_C_only=A_and_C_only,
                B_and_C_only=B_and_C_only,
                A_and_B_and_C=intersection_ABC
            )

        # Generate TSV output
        if n_sets == 2:
            # Special format for 2 sets with Venn diagram
            headers = ["Set", "Count", "Only", "Intersection"]
            rows = [
                [labels[0], counts[0], venn_diagram_2.A_only, venn_diagram_2.A_and_B],
                [labels[1], counts[1], venn_diagram_2.B_only, venn_diagram_2.A_and_B],
                ["Union", union_count, "", ""]
            ]
        elif n_sets == 3:
            # Special format for 3 sets with Venn diagram
            headers = ["Zone", "Count"]
            rows = [
                [f"{labels[0]} only", venn_diagram_3.A_only],
                [f"{labels[1]} only", venn_diagram_3.B_only],
                [f"{labels[2]} only", venn_diagram_3.C_only],
                [f"{labels[0]} ∩ {labels[1]} only", venn_diagram_3.A_and_B_only],
                [f"{labels[0]} ∩ {labels[2]} only", venn_diagram_3.A_and_C_only],
                [f"{labels[1]} ∩ {labels[2]} only", venn_diagram_3.B_and_C_only],
                [f"{labels[0]} ∩ {labels[1]} ∩ {labels[2]}", venn_diagram_3.A_and_B_and_C],
                ["", ""],
                ["Union", union_count]
            ]
        else:
            # General format for 4+ sets
            headers = ["Set", "Count", "Percentage"]
            rows = [[label, count, f"{sets_info[label].percentage}%"] for label, count in zip(labels, counts)]
            rows.append(["", "", ""])
            rows.append(["Pairwise Intersections", "", ""])
            for key, count in pairwise_intersections.items():
                rows.append([key, count, ""])
            rows.append(["", "", ""])
            rows.append(["Union", union_count, f"{union_percentage:.2f}%"])

        tsv = self._tsv_formatter.format_table(headers, rows)

        excel_output = ExcelOutput(
            tsv=tsv,
            formula=None,
            references=None
        )

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = total_rows
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, total_rows, True)

        return AnalyzeOverlapResponse(
            sets=sets_info,
            pairwise_intersections=pairwise_intersections,
            union_count=union_count,
            union_percentage=round(union_percentage, 2),
            venn_diagram_2=venn_diagram_2,
            venn_diagram_3=venn_diagram_3,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )
