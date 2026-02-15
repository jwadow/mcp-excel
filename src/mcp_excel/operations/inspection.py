# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Inspection operations for Excel files."""

import time

import pandas as pd

from ..core.file_loader import FileLoader
from ..models.requests import (
    CompareSheetsRequest,
    FindColumnRequest,
    GetColumnNamesRequest,
    GetDataProfileRequest,
    GetSheetInfoRequest,
    InspectFileRequest,
    SearchAcrossSheetsRequest,
)
from ..models.responses import (
    ColumnProfile,
    ColumnStats,
    CompareSheetsResponse,
    ExcelOutput,
    FindColumnResponse,
    GetColumnNamesResponse,
    GetDataProfileResponse,
    GetSheetInfoResponse,
    HeaderDetectionInfo,
    InspectFileResponse,
    SearchAcrossSheetsResponse,
)
from ..operations.base import BaseOperations, MAX_DIFFERENCES


class InspectionOperations(BaseOperations):
    """Operations for inspecting Excel file structure."""

    def __init__(self, file_loader: FileLoader) -> None:
        """Initialize inspection operations.

        Args:
            file_loader: FileLoader instance for loading files
        """
        super().__init__(file_loader)

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
            # Convert to JSON-serializable types with string keys and format values
            # Use _format_value() to handle numpy types, datetime, and other conversions
            serialized_row = {str(k): self._format_value(v) for k, v in row_dict.items()}
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

        # Find differences (with limit to prevent context overflow)
        differences = []
        truncated = False
        
        for _, row in merged.iterrows():
            # CONTEXT OVERFLOW PROTECTION: Stop if we hit the limit
            if len(differences) >= MAX_DIFFERENCES:
                truncated = True
                break
                
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
            
            # Add truncation warning if needed
            if truncated:
                tsv += f"\n\n[TRUNCATED: Showing first {MAX_DIFFERENCES} differences]"
        else:
            tsv = "No differences found"

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        metadata = self._get_file_metadata(request.file_path)
        performance = self._get_performance_metrics(
            start_time, len(df1) + len(df2), True
        )

        response = CompareSheetsResponse(
            differences=differences,
            difference_count=len(differences),
            truncated=truncated,
            key_column=request.key_column,
            compare_columns=request.compare_columns,
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )

        # CONTEXT OVERFLOW PROTECTION: Validate response size
        self._validate_response_size(
            response,
            rows_count=len(differences),
            columns_count=len(request.compare_columns) * 2 + 2  # *2 for sheet1/sheet2, +2 for key and status
        )

        return response

    def get_data_profile(
        self, request: GetDataProfileRequest
    ) -> GetDataProfileResponse:
        """Get data profile for columns.

        Args:
            request: Data profile request

        Returns:
            Data profile response with statistics and top values
        """
        start_time = time.time()

        # Load DataFrame with header detection
        df, _ = self._load_with_header_detection(
            request.file_path, request.sheet_name, request.header_row
        )

        # Determine which columns to profile
        if request.columns:
            # Validate requested columns
            missing_cols = set(request.columns) - set(df.columns)
            if missing_cols:
                available = ", ".join(str(col) for col in df.columns)
                raise ValueError(
                    f"Columns not found: {', '.join(missing_cols)}. Available: {available}"
                )
            columns_to_profile = request.columns
        else:
            # Profile all columns
            columns_to_profile = list(df.columns)

        # Build profiles for each column
        profiles = {}
        for col in columns_to_profile:
            col_str = str(col)
            col_data = df[col]

            # Determine data type
            if pd.api.types.is_datetime64_any_dtype(col_data):
                data_type = "datetime"
            elif pd.api.types.is_integer_dtype(col_data):
                data_type = "integer"
            elif pd.api.types.is_float_dtype(col_data):
                data_type = "float"
            elif pd.api.types.is_bool_dtype(col_data):
                data_type = "boolean"
            else:
                data_type = "string"

            # Basic counts
            total_count = len(col_data)
            null_count = int(col_data.isna().sum())
            null_percentage = round((null_count / total_count * 100) if total_count > 0 else 0.0, 2)
            unique_count = int(col_data.dropna().nunique())

            # Statistical summary (for numeric columns only)
            stats = None
            if data_type in ["integer", "float"]:
                non_null_data = col_data.dropna()
                if len(non_null_data) > 0:
                    stats = ColumnStats(
                        count=int(len(non_null_data)),
                        mean=float(non_null_data.mean()),
                        median=float(non_null_data.median()),
                        std=float(non_null_data.std()) if len(non_null_data) > 1 else 0.0,
                        min=self._format_value(non_null_data.min()),
                        max=self._format_value(non_null_data.max()),
                        q25=float(non_null_data.quantile(0.25)),
                        q75=float(non_null_data.quantile(0.75)),
                        null_count=null_count,
                    )

            # Top N most frequent values
            value_counts = col_data.value_counts().head(request.top_n)
            top_values = [
                {
                    "value": self._format_value(val),
                    "count": int(count),
                    "percentage": round((count / total_count * 100) if total_count > 0 else 0.0, 2)
                }
                for val, count in value_counts.items()
            ]

            # Create column profile
            profiles[col_str] = ColumnProfile(
                column_name=col_str,
                data_type=data_type,
                total_count=total_count,
                null_count=null_count,
                null_percentage=null_percentage,
                unique_count=unique_count,
                stats=stats,
                top_values=top_values,
            )

        # Generate TSV output
        tsv_rows = []
        tsv_rows.append(["Column", "Type", "Total", "Nulls", "Null%", "Unique", "Top Value", "Top Count"])
        
        for col_name, profile in profiles.items():
            top_val = profile.top_values[0]["value"] if profile.top_values else "N/A"
            top_count = profile.top_values[0]["count"] if profile.top_values else 0
            
            tsv_rows.append([
                col_name,
                profile.data_type,
                profile.total_count,
                profile.null_count,
                f"{profile.null_percentage}%",
                profile.unique_count,
                str(top_val),
                top_count,
            ])

        from ..excel.tsv_formatter import TSVFormatter
        tsv_formatter = TSVFormatter()
        tsv = tsv_formatter.format_table(tsv_rows[0], tsv_rows[1:])

        excel_output = ExcelOutput(tsv=tsv, formula=None, references=None)

        metadata = self._get_file_metadata(request.file_path, request.sheet_name)
        metadata.rows_total = len(df)
        metadata.columns_total = len(df.columns)

        performance = self._get_performance_metrics(start_time, len(df), True)

        return GetDataProfileResponse(
            profiles=profiles,
            columns_profiled=len(profiles),
            excel_output=excel_output,
            metadata=metadata,
            performance=performance,
        )
