# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Excel formula generator for dynamic calculations."""

from typing import Any, Optional

import pandas as pd

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
        if isinstance(value, pd.Timestamp):
            # Format datetime as DATE() function
            return self._format_date_for_excel(value)
        elif isinstance(value, str):
            # Escape quotes and protect from formula injection
            escaped = value.replace('"', '""')
            if escaped.startswith(("=", "+", "-", "@")):
                escaped = "'" + escaped
            return f'"{escaped}"'
        elif value is None:
            return '""'
        else:
            return str(value)
    
    def _format_date_for_excel(self, dt: pd.Timestamp) -> str:
        """Format datetime as Excel DATE() function.
        
        Args:
            dt: Pandas Timestamp
        
        Returns:
            Excel DATE() function string
        """
        if pd.isna(dt):
            return '""'
        
        # Excel DATE(year, month, day) function
        return f"DATE({dt.year},{dt.month},{dt.day})"
    
    def _convert_datetime_filters(
        self,
        filters: list[FilterCondition],
        column_types: dict[str, str]
    ) -> list[FilterCondition]:
        """Convert string datetime values in filters to pd.Timestamp.
        
        This ensures that datetime filter values are properly formatted as
        DATE() functions in Excel formulas.
        
        Args:
            filters: List of filter conditions
            column_types: Mapping of column names to types
        
        Returns:
            List of filters with datetime values converted to pd.Timestamp
        """
        converted_filters = []
        
        for filter_cond in filters:
            # Check if this column is a datetime type
            col_type = column_types.get(filter_cond.column)
            
            if col_type == "datetime":
                # Convert string datetime values to pd.Timestamp
                if isinstance(filter_cond.value, str):
                    try:
                        # Parse ISO 8601 string to Timestamp
                        converted_value = pd.to_datetime(filter_cond.value)
                        # Create new filter with converted value
                        filter_cond = FilterCondition(
                            column=filter_cond.column,
                            operator=filter_cond.operator,
                            value=converted_value,
                            values=filter_cond.values,
                            negate=filter_cond.negate
                        )
                    except Exception:
                        # If conversion fails, keep original value
                        pass
                
                # Also convert values list if present (for 'in' operator)
                if filter_cond.values:
                    converted_values = []
                    for val in filter_cond.values:
                        if isinstance(val, str):
                            try:
                                converted_values.append(pd.to_datetime(val))
                            except Exception:
                                converted_values.append(val)
                        else:
                            converted_values.append(val)
                    
                    filter_cond = FilterCondition(
                        column=filter_cond.column,
                        operator=filter_cond.operator,
                        value=filter_cond.value,
                        values=converted_values,
                        negate=filter_cond.negate
                    )
            
            converted_filters.append(filter_cond)
        
        return converted_filters

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
        column_types: Optional[dict[str, str]] = None,
    ) -> Optional[str]:
        """Generate formula from filter conditions.

        Args:
            operation: Operation type (count, sum, mean, etc.)
            filters: List of filter conditions
            column_ranges: Mapping of column names to Excel ranges
            target_range: Target range for aggregation (required for sum/mean)
            column_types: Optional mapping of column names to types (for datetime handling)

        Returns:
            Excel formula string, or None if filters use operators not supported in Excel

        Raises:
            ValueError: If operation is not supported or parameters are invalid
        """
        # Convert datetime filter values if column types are provided
        if column_types:
            filters = self._convert_datetime_filters(filters, column_types)
        
        if not filters:
            # No filters - simple aggregation
            if operation == "count" and target_range:
                return f"=COUNTA({target_range})"
            elif operation == "sum" and target_range:
                return f"=SUM({target_range})"
            elif operation == "mean" and target_range:
                return f"=AVERAGE({target_range})"
            elif operation == "median" and target_range:
                return f"=MEDIAN({target_range})"
            elif operation == "min" and target_range:
                return f"=MIN({target_range})"
            elif operation == "max" and target_range:
                return f"=MAX({target_range})"
            elif operation == "std" and target_range:
                return f"=STDEV({target_range})"
            elif operation == "var" and target_range:
                return f"=VAR({target_range})"
            else:
                raise ValueError(f"Operation {operation} requires filters or target range")

        # Single filter
        if len(filters) == 1:
            filter_cond = filters[0]
            criteria_range = column_ranges.get(filter_cond.column)
            if not criteria_range:
                raise ValueError(f"Column {filter_cond.column} not found in ranges")

            return self._generate_single_filter_formula(
                operation, filter_cond, criteria_range, target_range
            )

        # Multiple filters - use SUMIFS/COUNTIFS
        return self._generate_multiple_filters_formula(
            operation, filters, column_ranges, target_range
        )
    
    def _generate_single_filter_formula(
        self,
        operation: str,
        filter_cond: FilterCondition,
        criteria_range: str,
        target_range: Optional[str],
    ) -> Optional[str]:
        """Generate formula for single filter condition.
        
        Returns:
            Excel formula string, or None if operator is not supported in Excel
        """
        # Negation is not supported in Excel formulas (too complex)
        # Python filtering works correctly, but Excel formula cannot be generated
        if filter_cond.negate:
            return None
        
        operator = filter_cond.operator
        
        # Comparison operators: ==, !=, >, <, >=, <=
        if operator in ["==", "!=", ">", "<", ">=", "<="]:
            criteria = self._format_criteria(operator, filter_cond.value)
            
            if operation == "count":
                return f"=COUNTIF({criteria_range},{criteria})"
            elif operation == "sum" and target_range:
                return f"=SUMIF({criteria_range},{criteria},{target_range})"
            elif operation == "mean" and target_range:
                return f"=AVERAGEIF({criteria_range},{criteria},{target_range})"
        
        # Set operators: in, not_in
        elif operator == "in":
            if not filter_cond.values:
                return "=NA()  // 'in' operator requires values"
            
            # Use SUMPRODUCT for multiple values
            conditions = "+".join(
                f"({criteria_range}={self._escape_value(v)})"
                for v in filter_cond.values
            )
            
            if operation == "count":
                return f"=SUMPRODUCT({conditions})"
            elif operation == "sum" and target_range:
                return f"=SUMPRODUCT(({conditions})*{target_range})"
            elif operation == "mean" and target_range:
                # Average = Sum / Count
                sum_formula = f"SUMPRODUCT(({conditions})*{target_range})"
                count_formula = f"SUMPRODUCT({conditions})"
                return f"={sum_formula}/{count_formula}"
        
        elif operator == "not_in":
            if not filter_cond.values:
                return "=NA()  // 'not_in' operator requires values"
            
            # Total count minus matching count
            conditions = "+".join(
                f"({criteria_range}={self._escape_value(v)})"
                for v in filter_cond.values
            )
            
            if operation == "count":
                return f"=COUNTA({criteria_range})-SUMPRODUCT({conditions})"
            else:
                return "=NA()  // 'not_in' with sum/mean not supported in Excel formulas"
        
        # Text operators: contains, startswith, endswith
        elif operator == "contains":
            # Use wildcards: *text*
            criteria = f'"*{filter_cond.value}*"'
            
            if operation == "count":
                return f"=COUNTIF({criteria_range},{criteria})"
            elif operation == "sum" and target_range:
                return f"=SUMIF({criteria_range},{criteria},{target_range})"
            elif operation == "mean" and target_range:
                return f"=AVERAGEIF({criteria_range},{criteria},{target_range})"
        
        elif operator == "startswith":
            # Use wildcards: text*
            criteria = f'"{filter_cond.value}*"'
            
            if operation == "count":
                return f"=COUNTIF({criteria_range},{criteria})"
            elif operation == "sum" and target_range:
                return f"=SUMIF({criteria_range},{criteria},{target_range})"
            elif operation == "mean" and target_range:
                return f"=AVERAGEIF({criteria_range},{criteria},{target_range})"
        
        elif operator == "endswith":
            # Use wildcards: *text
            criteria = f'"*{filter_cond.value}"'
            
            if operation == "count":
                return f"=COUNTIF({criteria_range},{criteria})"
            elif operation == "sum" and target_range:
                return f"=SUMIF({criteria_range},{criteria},{target_range})"
            elif operation == "mean" and target_range:
                return f"=AVERAGEIF({criteria_range},{criteria},{target_range})"
        
        # Null operators: is_null, is_not_null
        elif operator == "is_null":
            if operation == "count":
                # BUG: Excel formula generation not supported for is_null operator
                #
                # PROBLEM:
                # - COUNTBLANK($A:$A) counts ALL empty cells in entire column (65536 rows in .xls, 1M+ in .xlsx)
                # - COUNTIF($A:$A,"") also counts ALL empty cells in entire column
                # - Python counts only empty cells in loaded data (e.g., 120 rows)
                # - Result: Excel returns 65416, Python returns 0 - massive discrepancy
                #
                # ROOT CAUSE:
                # - Full column references ($A:$A) include all rows in Excel sheet
                # - Excel cannot distinguish "empty cells in data" from "empty cells beyond data"
                # - Pandas loads only actual data rows, ignoring empty rows beyond data
                #
                # CORRECT SOLUTION:
                # - Use specific range instead of full column: $A$4:$A$123 (based on actual data boundaries)
                # - Requires DataRange with start_row/end_row calculated from len(df) + header_row
                # - Would need to pass DataRange to FormulaGenerator and use it ONLY for is_null
                #
                # CURRENT DECISION:
                # - Do not generate formula (return None)
                # - Python result is correct and will be used by agent
                # - Pragmatic approach: avoid complex fix for single edge case operator
                # - If formulas become critical, implement DataRange solution
                return None
            else:
                return "=NA()  // 'is_null' with sum/mean not supported"
        
        elif operator == "is_not_null":
            if operation == "count":
                return f"=COUNTA({criteria_range})"
            elif operation == "sum" and target_range:
                return f"=SUM({target_range})"
            elif operation == "mean" and target_range:
                return f"=AVERAGE({target_range})"
        
        # Regex - not supported in Excel formulas, but works in Python filtering
        elif operator == "regex":
            return None
        
        else:
            return f"=NA()  // Operator '{operator}' not supported"
    
    def _generate_multiple_filters_formula(
        self,
        operation: str,
        filters: list[FilterCondition],
        column_ranges: dict[str, str],
        target_range: Optional[str],
    ) -> Optional[str]:
        """Generate formula for multiple filter conditions.
        
        Returns:
            Excel formula string, or None if filters contain operators not supported in Excel
        """
        # Check if any filter uses negation (NOT operator)
        has_negation = any(f.negate for f in filters)
        
        if has_negation:
            # Negation is not supported in Excel formulas (too complex)
            # Python filtering works correctly, but Excel formula cannot be generated
            return None
        
        # Operators supported in Excel formulas (COUNTIFS/SUMIFS with wildcards)
        excel_supported = ["==", "!=", ">", "<", ">=", "<=", "contains", "startswith", "endswith"]
        
        # Operators that work in Python filtering but not in Excel formulas
        python_only = ["regex", "in", "not_in", "is_null", "is_not_null"]
        
        # Check if any filter uses python_only operators
        has_python_only = any(f.operator in python_only for f in filters)
        
        if has_python_only:
            # Python filtering works correctly, but Excel formula cannot be generated
            return None
        
        # Check for truly unsupported operators
        for filter_cond in filters:
            if filter_cond.operator not in (excel_supported + python_only):
                # Unknown operator - this is an error
                return f"=NA()  // Unsupported operator: '{filter_cond.operator}'"
        
        # All operators are Excel-supported - generate formula
        criteria_ranges = []
        criteria_values = []
        
        for filter_cond in filters:
            criteria_range = column_ranges.get(filter_cond.column)
            if not criteria_range:
                continue
            
            criteria_ranges.append(criteria_range)
            criteria_values.append(
                self._format_criteria(filter_cond.operator, filter_cond.value)
            )
        
        if operation == "count":
            return self.generate_countifs(criteria_ranges, criteria_values)
        elif operation == "sum" and target_range:
            return self.generate_sumifs(target_range, criteria_ranges, criteria_values)
        else:
            return f"=NA()  // Operation '{operation}' with multiple filters not fully supported"
    
    def _format_criteria(self, operator: str, value: Any) -> str:
        """Format criteria for COUNTIF/SUMIF functions.
        
        Args:
            operator: Comparison operator
            value: Value to compare
            
        Returns:
            Formatted criteria string for Excel
            
        Examples:
            == with "text" → "text"
            != with "text" → "<>text"
            > with 10 → ">10"
            >= with 10 → ">=10"
            >= with date → ">="&DATE(2026,1,1)
            contains with "ABC" → "*ABC*"
            startswith with "ABC" → "ABC*"
            endswith with "ABC" → "*ABC"
        """
        # Handle datetime values specially
        if isinstance(value, pd.Timestamp):
            date_func = self._format_date_for_excel(value)
            if operator == "==":
                return date_func
            elif operator in [">", "<", ">=", "<=", "!="]:
                # For comparison operators with dates, use: ">="&DATE(...)
                excel_op = "<>" if operator == "!=" else operator
                return f'"{excel_op}"&{date_func}'
            else:
                return date_func
        
        if operator == "==":
            # Simple equality
            return self._escape_value(value)
        
        elif operator == "!=":
            # Not equal: "<>value"
            if isinstance(value, str):
                # For strings: "<>text"
                escaped = value.replace('"', '""')
                return f'"<>{escaped}"'
            else:
                # For numbers: "<>10"
                return f'"<>{value}"'
        
        elif operator in [">", "<", ">=", "<="]:
            # Comparison operators: ">10", ">=10", etc.
            if isinstance(value, str):
                # For strings: ">text"
                escaped = value.replace('"', '""')
                return f'"{operator}{escaped}"'
            else:
                # For numbers: ">10"
                return f'"{operator}{value}"'
        
        elif operator == "contains":
            # Wildcard pattern: *value*
            if not isinstance(value, str):
                value = str(value)
            escaped = value.replace('"', '""')
            return f'"*{escaped}*"'
        
        elif operator == "startswith":
            # Wildcard pattern: value*
            if not isinstance(value, str):
                value = str(value)
            escaped = value.replace('"', '""')
            return f'"{escaped}*"'
        
        elif operator == "endswith":
            # Wildcard pattern: *value
            if not isinstance(value, str):
                value = str(value)
            escaped = value.replace('"', '""')
            return f'"*{escaped}"'
        
        else:
            # Fallback
            return self._escape_value(value)

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
