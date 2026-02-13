"""Inspection operations for Excel files."""

import time
from typing import Any

import pandas as pd
import psutil

from ..core.file_loader import FileLoader
from ..core.header_detector import HeaderDetector
from ..models.requests import (
    FindColumnRequest,
    GetColumnNamesRequest,
    GetSheetInfoRequest,
    InspectFileRequest,
)
from ..models.responses import (
    FileMetadata,
    FindColumnResponse,
    GetColumnNamesResponse,
    GetSheetInfoResponse,
    HeaderDetectionInfo,
    InspectFileResponse,
    PerformanceMetrics,
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

        warning = None
        recommendation = None
        if file_info["format"] == "xls":
            warning = "File is in legacy format. Write operations are not available."
            recommendation = "Consider converting to .xlsx for extended functionality."

        return FileMetadata(
            file_format=file_info["format"],
            sheet_name=sheet_name,
            rows_total=None,
            columns_total=None,
            warning=warning,
            recommendation=recommendation,
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
            row_dict = {str(k): (None if pd.isna(v) else v) for k, v in row_dict.items()}
            sample_rows.append(row_dict)

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
                df = self._loader.load(request.file_path, sheet_name, use_cache=True)
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
