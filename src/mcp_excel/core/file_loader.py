"""File loader with automatic format detection and caching."""

import os
from pathlib import Path
from typing import Optional

import pandas as pd

from .cache import FileCache


class FileLoader:
    """Loads Excel files with automatic format detection and caching."""

    def __init__(self, cache: Optional[FileCache] = None) -> None:
        """Initialize file loader.

        Args:
            cache: Optional FileCache instance. If None, creates default cache.
        """
        self._cache = cache or FileCache()

    def _detect_format(self, file_path: Path) -> str:
        """Detect Excel file format from extension.

        Args:
            file_path: Path to the file

        Returns:
            File format: 'xls' or 'xlsx'

        Raises:
            ValueError: If file format is not supported
        """
        suffix = file_path.suffix.lower()
        if suffix == ".xls":
            return "xls"
        elif suffix == ".xlsx":
            return "xlsx"
        else:
            raise ValueError(f"Unsupported file format: {suffix}. Only .xls and .xlsx are supported.")

    def _get_engine(self, file_format: str) -> str:
        """Get appropriate pandas engine for file format.

        Args:
            file_format: File format ('xls' or 'xlsx')

        Returns:
            Engine name for pandas
        """
        if file_format == "xls":
            return "xlrd"
        else:
            return "openpyxl"

    def load(
        self,
        file_path: str | Path,
        sheet_name: Optional[str | int] = 0,
        header_row: Optional[int] = None,
        use_cache: bool = True,
    ) -> pd.DataFrame:
        """Load Excel file into DataFrame.

        Args:
            file_path: Absolute path to the Excel file
            sheet_name: Sheet name or index (default: 0 = first sheet)
            header_row: Row index to use as header (None = auto-detect)
            use_cache: Whether to use cache (default: True)

        Returns:
            Loaded DataFrame

        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If file format is unsupported
            Exception: If file cannot be read
        """
        path = Path(file_path)

        if not path.exists():
            raise FileNotFoundError(
                f"File not found: {file_path}\n"
                f"Please use an absolute path to the file."
            )

        # Try cache first
        if use_cache:
            sheet_key = str(sheet_name) if sheet_name is not None else "0"
            cached_df = self._cache.get(path, sheet_key)
            if cached_df is not None:
                return cached_df

        # Detect format and engine
        file_format = self._detect_format(path)
        engine = self._get_engine(file_format)

        # Load file
        try:
            df = pd.read_excel(
                path,
                sheet_name=sheet_name,
                engine=engine,
                header=header_row,
            )

            # Cache the result
            if use_cache:
                sheet_key = str(sheet_name) if sheet_name is not None else "0"
                self._cache.put(path, df, sheet_key)

            return df

        except Exception as e:
            raise Exception(f"Failed to load file {file_path}: {str(e)}") from e

    def get_sheet_names(self, file_path: str | Path) -> list[str]:
        """Get list of sheet names in Excel file.

        Args:
            file_path: Absolute path to the Excel file

        Returns:
            List of sheet names

        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If file format is unsupported
        """
        path = Path(file_path)

        if not path.exists():
            raise FileNotFoundError(
                f"File not found: {file_path}\n"
                f"Please use an absolute path to the file."
            )

        file_format = self._detect_format(path)
        engine = self._get_engine(file_format)

        try:
            excel_file = pd.ExcelFile(path, engine=engine)
            return excel_file.sheet_names
        except Exception as e:
            raise Exception(f"Failed to read sheet names from {file_path}: {str(e)}") from e

    def get_file_info(self, file_path: str | Path) -> dict[str, any]:
        """Get basic information about Excel file.

        Args:
            file_path: Absolute path to the Excel file

        Returns:
            Dictionary with file information

        Raises:
            FileNotFoundError: If file doesn't exist
        """
        path = Path(file_path)

        if not path.exists():
            raise FileNotFoundError(
                f"File not found: {file_path}\n"
                f"Please use an absolute path to the file."
            )

        file_format = self._detect_format(path)
        file_size = os.path.getsize(path)
        sheet_names = self.get_sheet_names(path)

        return {
            "format": file_format,
            "size_bytes": file_size,
            "size_mb": round(file_size / 1024 / 1024, 2),
            "sheet_count": len(sheet_names),
            "sheet_names": sheet_names,
        }

    def invalidate_cache(self, file_path: str | Path) -> None:
        """Invalidate cache for specific file.

        Args:
            file_path: Path to the file
        """
        self._cache.invalidate(Path(file_path))

    def clear_cache(self) -> None:
        """Clear entire cache."""
        self._cache.clear()

    def get_cache_stats(self) -> dict[str, any]:
        """Get cache statistics.

        Returns:
            Dictionary with cache stats
        """
        return self._cache.get_stats()
