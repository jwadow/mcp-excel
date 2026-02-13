"""Excel formula generator for dynamic calculations."""

from typing import Any, Optional

from ..models.requests import FilterCondition


class FormulaGenerator:
    """Generates Excel formulas from operations and filters."""

    def __init__(self, sheet_name: str) -> None:
        """Initialize formula generator.

        Args:
            sheet_name: Name of the sheet for cell references
        """
        self._sheet_name = self._escape_sheet_name(sheet_name)

    def _escape_sheet_name(self, sheet_name: str) -> str:
        """Escape sheet name for Excel formula.

        Args:
            sheet_name: Raw sheet name

        Returns:
            Escaped sheet name with quotes if needed
        """
        # If sheet name contains spaces or special chars, wrap in quotes
        if " " in sheet_name or any(c in sheet_name for c in ["!", "'", "-"]):
            return f"'{sheet_name}'"
        return sheet_name

    def _escape_value(self, value: Any) -> str:
        """Escape value for use in Excel formula.

        Args:
            value: Value to escape

        Returns:
            Escaped value as string
        """
        if isinstance(value, str):
            # Escape quotes and protect from formula injection
            escaped = value.replace('"', '""')
            if escaped.startswith(("=", "+", "-", "@")):
                escaped = "'" + escaped
            return f'"{escaped}"'
        elif value is None:
            return '""'
        else:
            return str(value)

    def _column_letter(self, col_index: int) -> str:
        """Convert column index to Excel letter.

        Args:
            col_index: Zero-based column index

        Returns:
            Excel column letter (A, B, ..., Z, AA, AB, ...)
        """
        result = ""
        col_index += 1  # Excel is 1-based
        while col_index > 0:
            col_index -= 1
            result = chr(65 + (col_index % 26)) + result
            col_index //= 26
        return result

    def _get_column_range(self, column_name: str, column_index: int) -> str:
        """Get Excel range for a column.

        Args:
            column_name: Column name
            column_index: Zero-based column index

        Returns:
            Excel range (e.g., "Sheet1!$A:$A")
        """
        col_letter = self._column_letter(column_index)
        return f"{self._sheet_name}!${col_letter}:${col_letter}"

    def generate_countif(
        self, column_range: str, value: Any
    ) -> str:
        """Generate COUNTIF formula.

        Args:
            column_range: Excel column range
            value: Value to count

        Returns:
            COUNTIF formula
        """
        escaped_value = self._escape_value(value)
        return f"=COUNTIF({column_range},{escaped_value})"

    def generate_sumif(
        self, criteria_range: str, criteria: Any, sum_range: str
    ) -> str:
        """Generate SUMIF formula.

        Args:
            criteria_range: Range to check criteria
            criteria: Criteria value
            sum_range: Range to sum

        Returns:
            SUMIF formula
        """
        escaped_criteria = self._escape_value(criteria)
        return f"=SUMIF({criteria_range},{escaped_criteria},{sum_range})"

    def generate_sumifs(
        self, sum_range: str, criteria_ranges: list[str], criteria_values: list[Any]
    ) -> str:
        """Generate SUMIFS formula for multiple criteria.

        Args:
            sum_range: Range to sum
            criteria_ranges: List of ranges to check
            criteria_values: List of criteria values

        Returns:
            SUMIFS formula
        """
        criteria_parts = []
        for range_ref, value in zip(criteria_ranges, criteria_values):
            escaped_value = self._escape_value(value)
            criteria_parts.extend([range_ref, escaped_value])

        criteria_str = ",".join(criteria_parts)
        return f"=SUMIFS({sum_range},{criteria_str})"

    def generate_averageif(
        self, criteria_range: str, criteria: Any, average_range: str
    ) -> str:
        """Generate AVERAGEIF formula.

        Args:
            criteria_range: Range to check criteria
            criteria: Criteria value
            average_range: Range to average

        Returns:
            AVERAGEIF formula
        """
        escaped_criteria = self._escape_value(criteria)
        return f"=AVERAGEIF({criteria_range},{escaped_criteria},{average_range})"

    def generate_countifs(
        self, criteria_ranges: list[str], criteria_values: list[Any]
    ) -> str:
        """Generate COUNTIFS formula for multiple criteria.

        Args:
            criteria_ranges: List of ranges to check
            criteria_values: List of criteria values

        Returns:
            COUNTIFS formula
        """
        criteria_parts = []
        for range_ref, value in zip(criteria_ranges, criteria_values):
            escaped_value = self._escape_value(value)
            criteria_parts.extend([range_ref, escaped_value])

        criteria_str = ",".join(criteria_parts)
        return f"=COUNTIFS({criteria_str})"

    def generate_from_filter(
        self,
        operation: str,
        filters: list[FilterCondition],
        column_ranges: dict[str, str],
        target_range: Optional[str] = None,
    ) -> str:
        """Generate formula from filter conditions.

        Args:
            operation: Operation type (count, sum, mean, etc.)
            filters: List of filter conditions
            column_ranges: Mapping of column names to Excel ranges
            target_range: Target range for aggregation (required for sum/mean)

        Returns:
            Excel formula

        Raises:
            ValueError: If operation is not supported or parameters are invalid
        """
        if not filters:
            # No filters - simple aggregation
            if operation == "count" and target_range:
                return f"=COUNTA({target_range})"
            elif operation == "sum" and target_range:
                return f"=SUM({target_range})"
            elif operation == "mean" and target_range:
                return f"=AVERAGE({target_range})"
            else:
                raise ValueError(f"Operation {operation} requires filters or target range")

        # Single filter
        if len(filters) == 1:
            filter_cond = filters[0]
            criteria_range = column_ranges.get(filter_cond.column)
            if not criteria_range:
                raise ValueError(f"Column {filter_cond.column} not found in ranges")

            if filter_cond.operator == "==":
                if operation == "count":
                    return self.generate_countif(criteria_range, filter_cond.value)
                elif operation == "sum" and target_range:
                    return self.generate_sumif(criteria_range, filter_cond.value, target_range)
                elif operation == "mean" and target_range:
                    return self.generate_averageif(criteria_range, filter_cond.value, target_range)

        # Multiple filters - use SUMIFS/COUNTIFS
        criteria_ranges = []
        criteria_values = []

        for filter_cond in filters:
            if filter_cond.operator != "==":
                # Complex operators not supported in simple formulas
                return f"=NA()  // Complex filter not supported in formula"

            criteria_range = column_ranges.get(filter_cond.column)
            if not criteria_range:
                continue

            criteria_ranges.append(criteria_range)
            criteria_values.append(filter_cond.value)

        if operation == "count":
            return self.generate_countifs(criteria_ranges, criteria_values)
        elif operation == "sum" and target_range:
            return self.generate_sumifs(target_range, criteria_ranges, criteria_values)
        else:
            return f"=NA()  // Operation {operation} with multiple filters not supported"

    def get_references(
        self, column_names: list[str], column_indices: dict[str, int]
    ) -> dict[str, Any]:
        """Get cell references for columns.

        Args:
            column_names: List of column names
            column_indices: Mapping of column names to indices

        Returns:
            Dictionary with reference information
        """
        references = {}
        for col_name in column_names:
            col_idx = column_indices.get(col_name)
            if col_idx is not None:
                col_letter = self._column_letter(col_idx)
                references[col_name] = {
                    "column": col_letter,
                    "range": f"${col_letter}:${col_letter}",
                    "full_range": f"{self._sheet_name}!${col_letter}:${col_letter}",
                }

        return {
            "sheet": self._sheet_name,
            "columns": references,
        }
