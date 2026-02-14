# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""DateTime detection for Excel columns."""

from dataclasses import dataclass
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd


@dataclass
class DateTimeInfo:
    """Information about detected datetime column."""
    
    source: str  # "cell_format", "pandas_dtype", or "heuristic"
    confidence: float  # 0.0 to 1.0
    format_string: Optional[str] = None  # Excel format string if available


class DateTimeDetector:
    """Detects datetime columns using metadata and heuristics."""
    
    # Date format indicators in Excel format strings
    DATE_INDICATORS = [
        "dd", "mm", "yyyy", "yy", "hh", "ss",
        "d/m", "m/d", "y-m-d", "h:m:s",
        "mmmm", "mmm",  # Month names
    ]
    
    def detect_datetime_columns(
        self,
        df: pd.DataFrame,
        cell_formats: Optional[Dict[str, List[str]]] = None
    ) -> Dict[str, DateTimeInfo]:
        """Detect which columns contain datetime values.
        
        Args:
            df: DataFrame with data
            cell_formats: Optional dict of {column_name: [format_strings]}
                         from Excel cell formats
        
        Returns:
            Dictionary mapping column names to DateTimeInfo
        """
        datetime_columns = {}
        
        for col in df.columns:
            col_str = str(col)
            
            # Method 1: Check cell formats (if available)
            if cell_formats and col_str in cell_formats:
                if self._is_date_format(cell_formats[col_str]):
                    datetime_columns[col_str] = DateTimeInfo(
                        source="cell_format",
                        confidence=1.0,
                        format_string=cell_formats[col_str][0] if cell_formats[col_str] else None
                    )
                    continue
            
            # Method 2: Check Pandas dtype
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                datetime_columns[col_str] = DateTimeInfo(
                    source="pandas_dtype",
                    confidence=1.0
                )
                continue
            
            # Method 3: Check for datetime.datetime objects in object columns
            # This happens when Pandas auto-parses dates but column has mixed types
            if df[col].dtype == 'object':
                if self._contains_datetime_objects(df[col]):
                    datetime_columns[col_str] = DateTimeInfo(
                        source="pandas_dtype",
                        confidence=1.0
                    )
                    continue
            
            # Method 4: Heuristic for float64 (Excel date numbers)
            if pd.api.types.is_float_dtype(df[col]):
                if self._looks_like_excel_date(df[col]):
                    datetime_columns[col_str] = DateTimeInfo(
                        source="heuristic",
                        confidence=0.9
                    )
        
        return datetime_columns
    
    def _is_date_format(self, format_strings: List[str]) -> bool:
        """Check if any format string indicates a date.
        
        Args:
            format_strings: List of Excel format strings
        
        Returns:
            True if any format string contains date indicators
        """
        for format_str in format_strings:
            if not format_str:
                continue
            
            format_lower = format_str.lower()
            
            # Check for date indicators
            if any(indicator in format_lower for indicator in self.DATE_INDICATORS):
                return True
        
        return False
    
    def _looks_like_excel_date(self, series: pd.Series) -> bool:
        """Heuristic: does this series look like Excel dates?
        
        Args:
            series: Pandas Series to check
        
        Returns:
            True if series likely contains Excel date numbers
        """
        # Remove NaN values
        non_null = series.dropna()
        if len(non_null) == 0:
            return False
        
        # Check 1: Value range
        # Excel dates: 1 (1900-01-01) to ~60000 (2164 year)
        min_val = non_null.min()
        max_val = non_null.max()
        
        if not (1 <= min_val <= 60000 and 1 <= max_val <= 60000):
            return False
        
        # Check 2: Most values are whole numbers or have small fractional parts
        # Dates without time: whole numbers
        # Dates with time: fractional part < 1
        fractional_parts = non_null % 1
        mostly_dates = (fractional_parts < 0.0001).sum() / len(non_null) > 0.3
        
        if not mostly_dates:
            return False
        
        # Check 3: Values are not too dense (not sequential IDs)
        # Dates usually have gaps
        unique_ratio = len(non_null.unique()) / len(non_null)
        not_sequential_ids = unique_ratio > 0.1
        
        if not not_sequential_ids:
            return False
        
        # Check 4: Standard deviation check
        # Date ranges typically have reasonable std dev
        # IDs or counters have very different patterns
        std_dev = non_null.std()
        mean_val = non_null.mean()
        
        # Coefficient of variation should be reasonable for dates
        # Too low = sequential IDs, too high = random numbers
        if mean_val > 0:
            cv = std_dev / mean_val
            if not (0.001 < cv < 2.0):
                return False
        
        return True
    
    def _contains_datetime_objects(self, series: pd.Series) -> bool:
        """Check if series contains datetime.datetime objects.
        
        This happens when Pandas auto-parses dates during read_excel,
        but the column has mixed types (e.g., strings and datetimes).
        
        Args:
            series: Pandas Series to check
        
        Returns:
            True if series contains datetime objects
        """
        # Remove NaN values
        non_null = series.dropna()
        if len(non_null) == 0:
            return False
        
        # Check if majority of values are datetime objects
        datetime_count = 0
        for value in non_null.head(20):  # Sample first 20 values
            if isinstance(value, (datetime, pd.Timestamp)):
                datetime_count += 1
        
        # If >70% are datetime objects, it's a datetime column
        sample_size = min(20, len(non_null))
        return (datetime_count / sample_size) > 0.7
