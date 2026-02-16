# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Data validation operations for Excel files."""

import time

from ..core.file_loader import FileLoader
from ..excel.tsv_formatter import TSVFormatter
from ..models.requests import (
    FindDuplicatesRequest,
    FindNullsRequest,
)
from ..models.responses import (
    ExcelOutput,
    FindDuplicatesResponse,
    FindNullsResponse,
)
from ..operations.base import BaseOperations


class ValidationOperations(BaseOperations):
    """Data validation operations for Excel data."""

    def __init__(self, file_loader: FileLoader) -> None:
        """Initialize validation operations.

        Args:
            file_loader: FileLoader instance for loading files
        """
        super().__init__(file_loader)
        self._tsv_formatter = TSVFormatter()

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

        # Find all columns using normalized matching
        actual_columns = self._find_columns(df, request.columns, context="find_duplicates")

        # Find duplicates
        # duplicated() marks all duplicates except the first occurrence
        # keep=False marks all duplicates including first occurrence
        duplicate_mask = df.duplicated(subset=actual_columns, keep=False)
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

        response = FindDuplicatesResponse(
            duplicates=duplicates,
            duplicate_count=len(duplicates),
            columns_checked=request.columns,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

        # CONTEXT OVERFLOW PROTECTION: Validate response size
        self._validate_response_size(
            response,
            rows_count=len(duplicates),
            columns_count=len(df.columns)
        )

        return response

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

        # Find all columns using normalized matching
        actual_columns = self._find_columns(df, request.columns, context="find_nulls")

        # Analyze nulls for each column
        null_info = {}
        total_nulls = 0

        for col in actual_columns:
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

        for col in actual_columns:
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
