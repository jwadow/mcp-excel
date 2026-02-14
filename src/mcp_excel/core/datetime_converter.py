# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""DateTime conversion for Excel date numbers."""

import pandas as pd


class DateTimeConverter:
    """Converts Excel date numbers to datetime objects."""
    
    # Excel epoch for Windows (default)
    EXCEL_EPOCH_WINDOWS = pd.Timestamp("1899-12-30")
    
    # Excel epoch for Mac
    EXCEL_EPOCH_MAC = pd.Timestamp("1904-01-01")
    
    def convert_excel_number_to_datetime(
        self,
        value: float,
        epoch: str = "windows"
    ) -> pd.Timestamp:
        """Convert single Excel number to datetime.
        
        Args:
            value: Excel date number (e.g., 46060.7625)
            epoch: "windows" or "mac"
        
        Returns:
            pd.Timestamp object
        """
        if pd.isna(value):
            return pd.NaT
        
        epoch_date = (
            self.EXCEL_EPOCH_WINDOWS if epoch == "windows"
            else self.EXCEL_EPOCH_MAC
        )
        
        # Convert: epoch + number of days
        return epoch_date + pd.Timedelta(days=value)
    
    def convert_column(
        self,
        series: pd.Series,
        epoch: str = "windows"
    ) -> pd.Series:
        """Convert entire column of Excel numbers to datetime.
        
        Args:
            series: Pandas Series with Excel date numbers
            epoch: "windows" or "mac"
        
        Returns:
            Pandas Series with datetime64 dtype
        """
        # Vectorized conversion for performance
        epoch_date = (
            self.EXCEL_EPOCH_WINDOWS if epoch == "windows"
            else self.EXCEL_EPOCH_MAC
        )
        
        # Convert all values at once
        return pd.to_datetime(epoch_date) + pd.to_timedelta(series, unit='D')
    
    def detect_epoch(self, series: pd.Series) -> str:
        """Detect whether to use Windows or Mac epoch.
        
        Args:
            series: Series with Excel date numbers
        
        Returns:
            "windows" or "mac"
        """
        non_null = series.dropna()
        if len(non_null) == 0:
            return "windows"  # Default
        
        min_val = non_null.min()
        
        # If minimum value is less than 1462 (4 years from 1900),
        # it's likely Mac epoch (1904-based)
        # This is because dates before 1904 would be negative in Mac epoch
        if min_val < 1462:
            return "mac"
        
        return "windows"
