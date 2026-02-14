# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Inspection operations for Excel files."""

import time
from typing import Any

import pandas as pd
import psutil

from ..core.file_loader import FileLoader
from ..core.header_detector import HeaderDetector
from ..models.requests import (
    CompareSheetsRequest,
    FindColumnRequest,
    GetColumnNamesRequest,
    GetSheetInfoRequest,
    InspectFileRequest,
    SearchAcrossSheetsRequest,
)
from ..models.responses import (
    CompareSheetsResponse,
    ExcelOutput,
    FileMetadata,
    FindColumnResponse,
    GetColumnNamesResponse,
    GetSheetInfoResponse,
    HeaderDetectionInfo,
    InspectFileResponse,
    PerformanceMetrics,
    SearchAcrossSheetsResponse,
)


class InspectionOperations:
    """Operations for inspecting Excel file structure."""

    def __init__(self, file_loader: FileLoader) -> None:
        """Initialize inspection operations.

        Args:
            file_loader: FileLoader instance for loading files
        """
        self._loader = file_loader
        self._header_detector = HeaderDetector()

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
            file_path: Path to the file
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

    def inspect_file(self, request: InspectFileRequest) -> InspectFileResponse:
        """Inspect Excel file structure.

        Args:
            request: Inspection request

        Returns:
            File inspection response
        """
        start_time = time.time()

        file_info = self._loader.get_file_info(request.file_path)

        # Get basic info for each sheet
        sheets_info = []
        total_rows = 0

        for sheet_name in file_info["sheet_names"]:
            try:
                df = self._loader.load(request.file_path, sheet_name, header_row=None)
                row_count = len(df)
                col_count = len(df.columns)
                total_rows += row_count

                sheets_info.append({
                    "sheet_name": sheet_name,
                    "row_count": row_count,
                    "column_count": col_count,
                })
            except Exception as e:
                sheets_info.append({
                    "sheet_name": sheet_name,
                    "error": str(e),
                })

        metadata = self._get_file_metadata(request.file_path)
        performance = self._get_performance_metrics(start_time, total_rows, False)

        return InspectFileResponse(
            format=file_info["format"],
            size_bytes=file_info["size_bytes"],
            size_mb=file_info["size_mb"],
            sheet_count=file_info["sheet_count"],
            sheet_names=file_info["sheet_names"],
            sheets_info=sheets_info,
            metadata=metadata,
            performance=performance,
        )

    def get_sheet_info(self, request: GetSheetInfoRequest) -> GetSheetInfoResponse:
        """Get detailed information about a sheet.

        Args:
            request: Sheet info request

        Returns:
            Sheet information response
        """
        start_time = time.time()

        # Load without header first for detection
        df_raw = self._loader.load(
            request.file_path, request.sheet_name, header_row=None, use_cache=True
        )

        header_detection_info = None
        header_row = request.header_row

        # Auto-detect header if not specified
        if header_row is None:
            detection_result = self._header_detector.detect(df_raw)
            header_row = detection_result.header_row
            header_detection_info = HeaderDetectionInfo(
                header_row=detection_result.header_row,
                confidence=detection_result.confidence,
                candidates=detection_result.candidates,
            )

        # Reload with detected header
        df = self._loader.load(
            request.file_path, request.sheet_name, header_row=header_row, use_cache=True
        )

        # Get column types
        column_types = {}
        for col in df.columns:
            dtype = df[col].dtype
            if pd.api.types.is_integer_dtype(dtype):
                column_types[str(col)] = "integer"
            elif pd.api.types.is_float_dtype(dtype):
                column_types[str(col)] = "float"
            elif pd.api.types.is_datetime64_any_dtype(dtype):
                column_types[str(col)] = "datetime"
            elif pd.api.types.is_bool_dtype(dtype):
                column_types[str(col)] = "boolean"
            else:
                column_types[str(col)] = "string"

        # Get sample rows
        sample_rows = []
        for idx in range(min(3, len(df))):
            row_dict = df.iloc[idx].to_dict()
            # Convert to JSON-serializable types with string keys
            # CRITICAL: Convert datetime values to ISO 8601 strings for agent
            serialized_row = {}
            for k, v in row_dict.items():
                key = str(k)
                if pd.isna(v):
                    serialized_row[key] = None
                elif pd.api.types.is_datetime64_any_dtype(type(v)) or isinstance(v, pd.Timestamp):
                    # Convert datetime to ISO 8601 string
                    serialized_row[key] = v.isoformat()
                else:
                    serialized_row[key] = v
            sample_rows.append(serialized_row)

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        return GetSheetInfoResponse(
            sheet_name=request.sheet_name,
            column_names=[str(col) for col in df.columns],
            column_count=len(df.columns),
            column_types=column_types,
            row_count=len(df),
            data_start_row=header_row + 1 if header_row is not None else 0,
            sample_rows=sample_rows,
            header_detection=header_detection_info,
            metadata=metadata,
            performance=performance,
        )

    def get_column_names(self, request: GetColumnNamesRequest) -> GetColumnNamesResponse:
        """Get column names from a sheet.

        Args:
            request: Column names request

        Returns:
            Column names response
        """
        start_time = time.time()

        # Load with header detection
        df_raw = self._loader.load(
            request.file_path, request.sheet_name, header_row=None, use_cache=True
        )

        header_row = request.header_row
        if header_row is None:
            detection_result = self._header_detector.detect(df_raw)
            header_row = detection_result.header_row

        df = self._loader.load(
            request.file_path, request.sheet_name, header_row=header_row, use_cache=True
        )

        column_names = [str(col) for col in df.columns]

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        performance = self._get_performance_metrics(start_time, 0, True)

        return GetColumnNamesResponse(
            column_names=column_names,
            column_count=len(column_names),
            metadata=metadata,
            performance=performance,
        )

    def find_column(self, request: FindColumnRequest) -> FindColumnResponse:
        """Find column across sheets.

        Args:
            request: Find column request

        Returns:
            Find column response
        """
        start_time = time.time()

        file_info = self._loader.get_file_info(request.file_path)
        sheets_to_search = (
            file_info["sheet_names"] if request.search_all_sheets else [file_info["sheet_names"][0]]
        )

        found_in = []
        total_rows = 0

        for sheet_name in sheets_to_search:
            try:
                # Load with header detection
                df_raw = self._loader.load(request.file_path, sheet_name, header_row=None, use_cache=True)
                detection_result = self._header_detector.detect(df_raw)
                df = self._loader.load(request.file_path, sheet_name, header_row=detection_result.header_row, use_cache=True)
                
                total_rows += len(df)

                # Check if column exists (case-insensitive)
                column_names = [str(col) for col in df.columns]
                matching_cols = [
                    (idx, col)
                    for idx, col in enumerate(column_names)
                    if col.lower() == request.column_name.lower()
                ]

                if matching_cols:
                    for col_idx, col_name in matching_cols:
                        found_in.append({
                            "sheet": sheet_name,
                            "column_name": col_name,
                            "column_index": col_idx,
                            "row_count": len(df),
                        })

            except Exception:
                continue

        metadata = self._get_file_metadata(request.file_path)
        performance = self._get_performance_metrics(start_time, total_rows, True)

        return FindColumnResponse(
            found_in=found_in,
            total_matches=len(found_in),
            metadata=metadata,
            performance=performance,
        )

    def search_across_sheets(
        self, request: "SearchAcrossSheetsRequest"
    ) -> "SearchAcrossSheetsResponse":
        """Search for a value across all sheets.

        Args:
            request: Search request

        Returns:
            Search results with match counts per sheet
        """
        start_time = time.time()

        file_info = self._loader.get_file_info(request.file_path)
        matches = []
        total_matches = 0
        total_rows = 0

        for sheet_name in file_info["sheet_names"]:
            try:
                # Load with header detection
                df_raw = self._loader.load(request.file_path, sheet_name, header_row=None, use_cache=True)
                detection_result = self._header_detector.detect(df_raw)
                df = self._loader.load(request.file_path, sheet_name, header_row=detection_result.header_row, use_cache=True)
                
                total_rows += len(df)

                # Check if column exists (case-insensitive)
                column_names = [str(col) for col in df.columns]
                matching_col = None
                for col in column_names:
                    if col.lower() == request.column_name.lower():
                        matching_col = col
                        break

                if matching_col:
                    # Count matches in this column
                    if pd.api.types.is_numeric_dtype(df[matching_col]):
                        # Numeric comparison
                        match_count = int((df[matching_col] == request.value).sum())
                    else:
                        # String comparison (case-insensitive)
                        match_count = int(
                            df[matching_col]
                            .astype(str)
                            .str.lower()
                            .eq(str(request.value).lower())
                            .sum()
                        )

                    if match_count > 0:
                        matches.append({
                            "sheet": sheet_name,
                            "column_name": matching_col,
                            "match_count": match_count,
                            "total_rows": len(df),
                        })
                        total_matches += match_count

            except Exception:
                continue

        metadata = self._get_file_metadata(request.file_path)
        performance = self._get_performance_metrics(start_time, total_rows, True)

        return SearchAcrossSheetsResponse(
            matches=matches,
            total_matches=total_matches,
            column_name=request.column_name,
            value=request.value,
            metadata=metadata,
            performance=performance,
        )

    def compare_sheets(
        self, request: "CompareSheetsRequest"
    ) -> "CompareSheetsResponse":
        """Compare data between two sheets.

        Args:
            request: Comparison request

        Returns:
            Differences between sheets
        """
        start_time = time.time()

        # Load both sheets with header detection
        df1_raw = self._loader.load(
            request.file_path, request.sheet1, header_row=None, use_cache=True
        )
        df2_raw = self._loader.load(
            request.file_path, request.sheet2, header_row=None, use_cache=True
        )

        header_row = request.header_row
        if header_row is None:
            detection_result = self._header_detector.detect(df1_raw)
            header_row = detection_result.header_row

        df1 = self._loader.load(
            request.file_path, request.sheet1, header_row=header_row, use_cache=True
        )
        df2 = self._loader.load(
            request.file_path, request.sheet2, header_row=header_row, use_cache=True
        )

        # Normalize column names
        df1.columns = [str(col) for col in df1.columns]
        df2.columns = [str(col) for col in df2.columns]

        # Validate key column exists in both sheets
        if request.key_column not in df1.columns:
            raise ValueError(
                f"Key column '{request.key_column}' not found in sheet '{request.sheet1}'. "
                f"Available columns: {list(df1.columns)}"
            )
        if request.key_column not in df2.columns:
            raise ValueError(
                f"Key column '{request.key_column}' not found in sheet '{request.sheet2}'. "
                f"Available columns: {list(df2.columns)}"
            )

        # Validate compare columns exist in both sheets
        for col in request.compare_columns:
            if col not in df1.columns:
                raise ValueError(
                    f"Compare column '{col}' not found in sheet '{request.sheet1}'"
                )
            if col not in df2.columns:
                raise ValueError(
                    f"Compare column '{col}' not found in sheet '{request.sheet2}'"
                )

        # Merge on key column
        merged = df1.merge(
            df2,
            on=request.key_column,
            how="outer",
            suffixes=("_sheet1", "_sheet2"),
            indicator=True,
        )

        # Find differences
        differences = []
        for _, row in merged.iterrows():
            diff_entry = {request.key_column: row[request.key_column]}

            # Check if row exists in both sheets
            if row["_merge"] == "left_only":
                diff_entry["status"] = "only_in_sheet1"
                for col in request.compare_columns:
                    diff_entry[f"{col}_sheet1"] = row.get(f"{col}_sheet1")
                    diff_entry[f"{col}_sheet2"] = None
                differences.append(diff_entry)
            elif row["_merge"] == "right_only":
                diff_entry["status"] = "only_in_sheet2"
                for col in request.compare_columns:
                    diff_entry[f"{col}_sheet1"] = None
                    diff_entry[f"{col}_sheet2"] = row.get(f"{col}_sheet2")
                differences.append(diff_entry)
            else:
                # Row exists in both - check for value differences
                has_diff = False
                for col in request.compare_columns:
                    val1 = row.get(f"{col}_sheet1")
                    val2 = row.get(f"{col}_sheet2")

                    # Handle NaN comparison
                    if pd.isna(val1) and pd.isna(val2):
                        continue
                    if pd.isna(val1) or pd.isna(val2) or val1 != val2:
                        has_diff = True
                        diff_entry[f"{col}_sheet1"] = None if pd.isna(val1) else val1
                        diff_entry[f"{col}_sheet2"] = None if pd.isna(val2) else val2

                if has_diff:
                    diff_entry["status"] = "different_values"
                    differences.append(diff_entry)

        # Generate TSV output
        from ..excel.tsv_formatter import TSVFormatter

        tsv_formatter = TSVFormatter()
        if differences:
            headers = [request.key_column, "status"]
            for col in request.compare_columns:
                headers.extend([f"{col}_sheet1", f"{col}_sheet2"])

            rows = []
            for diff in differences:
                row = [diff.get(h) for h in headers]
                rows.append(row)

            tsv = tsv_formatter.format_table(headers, rows)
        else:
            tsv = "No differences found"

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        metadata = self._get_file_metadata(request.file_path)
        performance = self._get_performance_metrics(
            start_time, len(df1) + len(df2), True
        )

        return CompareSheetsResponse(
            differences=differences,
            difference_count=len(differences),
            key_column=request.key_column,
            compare_columns=request.compare_columns,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )
