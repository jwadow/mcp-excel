# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Intelligent header detection for Excel files with messy structure."""

from typing import Optional

import pandas as pd


class HeaderDetectionResult:
    """Result of header detection analysis."""

    def __init__(
        self,
        header_row: int,
        confidence: float,
        candidates: list[dict[str, any]],
    ) -> None:
        """Initialize detection result.

        Args:
            header_row: Detected header row index
            confidence: Confidence score (0.0 to 1.0)
            candidates: List of candidate rows with their scores
        """
        self.header_row = header_row
        self.confidence = confidence
        self.candidates = candidates


class HeaderDetector:
    """Detects header row in Excel files with non-standard structure."""

    def __init__(
        self,
        scan_rows: int = 20,
        min_confidence: float = 0.8,
    ) -> None:
        """Initialize header detector.

        Args:
            scan_rows: Number of rows to scan for headers
            min_confidence: Minimum confidence threshold for automatic detection
        """
        self._scan_rows = scan_rows
        self._min_confidence = min_confidence

    def _calculate_fill_rate(self, row: pd.Series) -> float:
        """Calculate fill rate for a row.

        Args:
            row: Pandas Series representing a row

        Returns:
            Fill rate (0.0 to 1.0)
        """
        non_empty = row.notna().sum()
        total = len(row)
        return non_empty / total if total > 0 else 0.0

    def _check_uniqueness(self, row: pd.Series) -> bool:
        """Check if row values are unique (no duplicates).

        Args:
            row: Pandas Series representing a row

        Returns:
            True if all non-null values are unique
        """
        non_null_values = row.dropna()
        return len(non_null_values) == len(non_null_values.unique())

    def _is_all_strings(self, row: pd.Series) -> float:
        """Check if row contains mostly strings.

        Args:
            row: Pandas Series representing a row

        Returns:
            Ratio of string values (0.0 to 1.0)
        """
        non_null = row.dropna()
        if len(non_null) == 0:
            return 0.0
        
        string_count = sum(1 for v in non_null if isinstance(v, str))
        return string_count / len(non_null)

    def _has_numeric_only_values(self, row: pd.Series) -> bool:
        """Check if row has values that are purely numeric strings.

        Args:
            row: Pandas Series representing a row

        Returns:
            True if row contains numeric-only string values (like '50765981')
        """
        non_null = row.dropna()
        for val in non_null:
            if isinstance(val, str) and val.strip().isdigit() and len(val) > 4:
                return True
        return False

    def _average_value_length(self, row: pd.Series) -> float:
        """Calculate average length of string values in row.

        Args:
            row: Pandas Series representing a row

        Returns:
            Average length of string values
        """
        non_null = row.dropna()
        if len(non_null) == 0:
            return 0.0
        
        lengths = [len(str(v)) for v in non_null]
        return sum(lengths) / len(lengths)

    def _check_previous_empty_rows(self, df: pd.DataFrame, row_idx: int) -> int:
        """Count empty rows before this row.

        Args:
            df: DataFrame to analyze
            row_idx: Index of candidate row

        Returns:
            Number of empty rows before this row
        """
        if row_idx == 0:
            return 0
        
        empty_count = 0
        for i in range(row_idx - 1, -1, -1):
            row = df.iloc[i]
            if row.notna().sum() == 0:
                empty_count += 1
            else:
                break
        
        return empty_count

    def _analyze_following_rows_consistency(
        self, df: pd.DataFrame, row_idx: int
    ) -> float:
        """Analyze if following rows have consistent structure.

        Args:
            df: DataFrame to analyze
            row_idx: Index of candidate header row

        Returns:
            Consistency score (0.0 to 1.0)
        """
        if row_idx >= len(df) - 3:
            return 0.0

        # Check next 5 rows for consistency
        next_rows = df.iloc[row_idx + 1 : min(row_idx + 6, len(df))]
        
        if len(next_rows) < 2:
            return 0.0

        # Calculate fill rates for following rows
        fill_rates = [self._calculate_fill_rate(next_rows.iloc[i]) for i in range(len(next_rows))]
        
        # Good data rows should have similar fill rates
        if len(fill_rates) < 2:
            return 0.0
        
        mean_fill = sum(fill_rates) / len(fill_rates)
        variance = sum((f - mean_fill) ** 2 for f in fill_rates) / len(fill_rates)
        
        # Low variance = high consistency
        consistency = 1.0 - min(variance * 10, 1.0)
        
        return consistency

    def _score_candidate(
        self, df: pd.DataFrame, row_idx: int
    ) -> float:
        """Score a candidate header row with improved algorithm.

        Args:
            df: DataFrame to analyze
            row_idx: Index of candidate row

        Returns:
            Score (0.0 to 1.0)
        """
        row = df.iloc[row_idx]
        score = 0.0

        # 1. Fill rate should be high (20%)
        fill_rate = self._calculate_fill_rate(row)
        score += 0.20 * fill_rate

        # 2. All values should be strings (25%)
        string_ratio = self._is_all_strings(row)
        score += 0.25 * string_ratio

        # 3. Should NOT have long numeric-only values (20%)
        has_numeric = self._has_numeric_only_values(row)
        if not has_numeric:
            score += 0.20

        # 4. Uniqueness is critical (15%)
        is_unique = self._check_uniqueness(row)
        if is_unique:
            score += 0.15

        # 5. Headers are usually shorter than data (10%)
        avg_length = self._average_value_length(row)
        if 5 <= avg_length <= 30:  # Reasonable header length
            score += 0.10

        # 6. Position bonus - earlier rows more likely (5%)
        position_bonus = max(0, 1.0 - (row_idx / self._scan_rows)) * 0.05
        score += position_bonus

        # 7. Empty rows before increase likelihood (5%)
        empty_before = self._check_previous_empty_rows(df, row_idx)
        if empty_before > 0:
            score += min(empty_before * 0.02, 0.05)

        return score

    def detect(self, df: pd.DataFrame) -> HeaderDetectionResult:
        """Detect header row in DataFrame.

        Args:
            df: DataFrame loaded without header specification

        Returns:
            HeaderDetectionResult with detected row and confidence

        Raises:
            ValueError: If DataFrame is empty
        """
        if df.empty:
            raise ValueError("Cannot detect header in empty DataFrame")

        scan_limit = min(self._scan_rows, len(df))

        # Score all candidate rows
        candidates = []
        for row_idx in range(scan_limit):
            row = df.iloc[row_idx]
            score = self._score_candidate(df, row_idx)

            candidates.append({
                "row": row_idx,
                "score": score,
                "preview": row.head(5).tolist(),
            })

        # Sort by score
        candidates.sort(key=lambda x: x["score"], reverse=True)

        # Best candidate
        best = candidates[0]
        header_row = best["row"]
        confidence = best["score"]

        return HeaderDetectionResult(
            header_row=header_row,
            confidence=confidence,
            candidates=candidates[:3],  # Return top 3 candidates
        )

    def detect_or_ask(
        self, df: pd.DataFrame
    ) -> tuple[Optional[int], Optional[list[dict[str, any]]]]:
        """Detect header or return candidates if confidence is low.

        Args:
            df: DataFrame loaded without header specification

        Returns:
            Tuple of (header_row, candidates).
            If confidence is high: (row_index, None)
            If confidence is low: (None, candidates_list)
        """
        result = self.detect(df)

        if result.confidence >= self._min_confidence:
            return (result.header_row, None)
        else:
            return (None, result.candidates)
