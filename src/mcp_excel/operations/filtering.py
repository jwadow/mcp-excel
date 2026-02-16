# Excel MCP Server
# Copyright (C) 2026 Jwadow
# Licensed under AGPL-3.0
# https://github.com/jwadow/mcp-excel

"""Filtering system for DataFrame operations."""

import re
import unicodedata
from difflib import get_close_matches
from typing import Any

import pandas as pd

from ..models.requests import FilterCondition, FilterGroup
from ..core.datetime_converter import DateTimeConverter


class FilterEngine:
    """Engine for applying filters to DataFrames."""
    
    def __init__(self) -> None:
        """Initialize filter engine."""
        self._datetime_converter = DateTimeConverter()
    
    def _normalize_column_name(self, name: str) -> str:
        """Normalize column name for robust matching.
        
        Handles:
        - Non-string column names (pandas can use integers: 0, 1, 2...)
        - Unicode normalization (NFC - composed form)
        - Non-breaking spaces (U+00A0) → regular spaces (U+0020)
        - Leading/trailing whitespace
        - Multiple consecutive spaces → single space
        
        Args:
            name: Column name to normalize (accepts str or int)
        
        Returns:
            Normalized column name (always str)
        
        Examples:
            >>> _normalize_column_name("café")  # NFC or NFD
            "café"  # Always NFC
            >>> _normalize_column_name("Нетто, кг")  # Non-breaking space
            "Нетто, кг"  # Regular space
            >>> _normalize_column_name("  Name  ")
            "Name"
            >>> _normalize_column_name(0)  # Integer column name
            "0"
        """
        # 0. Handle non-string column names (defensive programming)
        # Pandas can use integers (0, 1, 2...) when headers are not detected
        # This makes FilterEngine robust for direct usage (e.g., in tests)
        if not isinstance(name, str):
            name = str(name)
        
        # 1. Unicode normalization (NFC - composed form)
        # This ensures "café" (NFC) and "café" (NFD) are treated as identical
        name = unicodedata.normalize('NFC', name)
        
        # 2. Replace non-breaking spaces (U+00A0) with regular spaces (U+0020)
        # Excel often uses non-breaking spaces, which look identical but compare differently
        name = name.replace('\u00A0', ' ')
        
        # 3. Strip leading/trailing whitespace
        name = name.strip()
        
        # 4. Collapse multiple consecutive spaces into one
        # "Name  Value" → "Name Value"
        name = ' '.join(name.split())
        
        return name

    def apply_filters(
        self,
        df: pd.DataFrame,
        filters: list[FilterCondition | FilterGroup],
        logic: str = "AND",
    ) -> pd.DataFrame:
        """Apply filters to DataFrame with support for nested groups.

        Args:
            df: DataFrame to filter
            filters: List of filter conditions or nested groups
            logic: Logic operator ("AND" or "OR")

        Returns:
            NEW DataFrame (copy) containing only filtered rows.
            Caller owns the returned DataFrame and can modify it freely.

        Raises:
            ValueError: If filter is invalid
        """
        if not filters:
            return df

        # Build mask for each filter (may be condition or group)
        masks = []
        for filter_item in filters:
            if isinstance(filter_item, FilterCondition):
                # Atomic condition
                mask = self._build_filter_mask(df, filter_item)
            elif isinstance(filter_item, FilterGroup):
                # Nested group - RECURSION
                mask = self._build_group_mask(df, filter_item)
            else:
                raise ValueError(f"Invalid filter type: {type(filter_item)}")
            
            masks.append(mask)

        # Combine masks with logic operator
        if logic == "AND":
            combined_mask = masks[0]
            for mask in masks[1:]:
                combined_mask = combined_mask & mask
        elif logic == "OR":
            combined_mask = masks[0]
            for mask in masks[1:]:
                combined_mask = combined_mask | mask
        else:
            raise ValueError(f"Invalid logic operator: {logic}. Must be 'AND' or 'OR'")

        # Return explicit copy for clear ownership (architectural principle)
        return df[combined_mask].copy()

    def count_filtered(
        self,
        df: pd.DataFrame,
        filters: list[FilterCondition | FilterGroup],
        logic: str = "AND",
    ) -> int:
        """Count rows matching filters without materializing DataFrame.

        Args:
            df: DataFrame to filter
            filters: List of filter conditions or nested groups
            logic: Logic operator ("AND" or "OR")

        Returns:
            Count of matching rows
        """
        if not filters:
            return len(df)

        masks = []
        for filter_item in filters:
            if isinstance(filter_item, FilterCondition):
                # Atomic condition
                mask = self._build_filter_mask(df, filter_item)
            elif isinstance(filter_item, FilterGroup):
                # Nested group - RECURSION
                mask = self._build_group_mask(df, filter_item)
            else:
                raise ValueError(f"Invalid filter type: {type(filter_item)}")
            
            masks.append(mask)

        if logic == "AND":
            combined_mask = masks[0]
            for mask in masks[1:]:
                combined_mask = combined_mask & mask
        else:
            combined_mask = masks[0]
            for mask in masks[1:]:
                combined_mask = combined_mask | mask

        return int(combined_mask.sum())

    def _build_group_mask(
        self,
        df: pd.DataFrame,
        group: FilterGroup
    ) -> pd.Series:
        """Build boolean mask for a filter group (recursive).
        
        Supports nested groups for complex logical expressions like:
        - (A AND B) OR C
        - A AND (B OR C)
        - ((A OR B) AND C) OR D
        
        Args:
            df: DataFrame to filter
            group: Filter group with nested conditions
        
        Returns:
            Boolean Series mask
        
        Raises:
            ValueError: If filter is invalid
        """
        if not group.filters:
            # Empty group - return all True mask
            return pd.Series([True] * len(df), index=df.index)
        
        # Recursively build masks for each filter in the group
        masks = []
        for filter_item in group.filters:
            if isinstance(filter_item, FilterCondition):
                # Atomic condition
                mask = self._build_filter_mask(df, filter_item)
            elif isinstance(filter_item, FilterGroup):
                # Nested group - RECURSION
                mask = self._build_group_mask(df, filter_item)
            else:
                raise ValueError(f"Invalid filter type: {type(filter_item)}")
            
            masks.append(mask)
        
        # Combine masks with group's logic operator
        if group.logic == "AND":
            combined_mask = masks[0]
            for mask in masks[1:]:
                combined_mask = combined_mask & mask
        elif group.logic == "OR":
            combined_mask = masks[0]
            for mask in masks[1:]:
                combined_mask = combined_mask | mask
        else:
            raise ValueError(f"Invalid logic operator: {group.logic}")
        
        # Apply negation to entire group if requested (NOT operator)
        if group.negate:
            combined_mask = ~combined_mask
        
        return combined_mask

    def _build_filter_mask(
        self, df: pd.DataFrame, filter_cond: FilterCondition
    ) -> pd.Series:
        """Build boolean mask for a single filter condition.

        Args:
            df: DataFrame to filter
            filter_cond: Filter condition

        Returns:
            Boolean Series mask

        Raises:
            ValueError: If column doesn't exist or filter is invalid
        """
        # Normalize filter column name
        filter_col_normalized = self._normalize_column_name(filter_cond.column)
        
        # Build mapping: normalized name → original DataFrame column name
        normalized_to_original = {
            self._normalize_column_name(col): col
            for col in df.columns
        }
        
        # Find column using normalized comparison
        if filter_col_normalized not in normalized_to_original:
            # Column not found - provide helpful error with fuzzy matching
            suggestions = get_close_matches(
                filter_col_normalized,
                normalized_to_original.keys(),
                n=3,
                cutoff=0.6
            )
            
            available = ", ".join(str(col) for col in df.columns)
            suggestion_text = ""
            if suggestions:
                # Map back to original names for suggestions
                original_suggestions = [
                    normalized_to_original[s] for s in suggestions
                ]
                suggestion_text = f" Did you mean: {', '.join(repr(s) for s in original_suggestions)}?"
            
            raise ValueError(
                f"Column '{filter_cond.column}' not found.{suggestion_text} "
                f"Available columns: {available}"
            )
        
        # Use original column name from DataFrame
        actual_column = normalized_to_original[filter_col_normalized]
        col_data = df[actual_column]
        operator = filter_cond.operator
        
        # Parse filter value for datetime columns
        filter_value = filter_cond.value
        if pd.api.types.is_datetime64_any_dtype(col_data) and filter_value is not None:
            filter_value = self._parse_datetime_value(filter_value)

        # Comparison operators
        if operator == "==":
            mask = col_data == filter_value
        elif operator == "!=":
            mask = col_data != filter_value
        elif operator == ">":
            mask = col_data > filter_value
        elif operator == "<":
            mask = col_data < filter_value
        elif operator == ">=":
            mask = col_data >= filter_value
        elif operator == "<=":
            mask = col_data <= filter_value

        # Set operators
        elif operator == "in":
            if not filter_cond.values:
                raise ValueError("'in' operator requires 'values' parameter")
            # Parse datetime values if column is datetime
            values = filter_cond.values
            if pd.api.types.is_datetime64_any_dtype(col_data):
                values = [self._parse_datetime_value(v) for v in values]
            mask = col_data.isin(values)
        elif operator == "not_in":
            if not filter_cond.values:
                raise ValueError("'not_in' operator requires 'values' parameter")
            # Parse datetime values if column is datetime
            values = filter_cond.values
            if pd.api.types.is_datetime64_any_dtype(col_data):
                values = [self._parse_datetime_value(v) for v in values]
            mask = ~col_data.isin(values)

        # String operators
        elif operator == "contains":
            if not isinstance(filter_cond.value, str):
                raise ValueError("'contains' operator requires string value")
            mask = col_data.astype(str).str.contains(filter_cond.value, na=False, regex=False)
        elif operator == "startswith":
            if not isinstance(filter_cond.value, str):
                raise ValueError("'startswith' operator requires string value")
            mask = col_data.astype(str).str.startswith(filter_cond.value, na=False)
        elif operator == "endswith":
            if not isinstance(filter_cond.value, str):
                raise ValueError("'endswith' operator requires string value")
            mask = col_data.astype(str).str.endswith(filter_cond.value, na=False)
        elif operator == "regex":
            if not isinstance(filter_cond.value, str):
                raise ValueError("'regex' operator requires string value")
            try:
                mask = col_data.astype(str).str.contains(filter_cond.value, na=False, regex=True)
            except re.error as e:
                raise ValueError(f"Invalid regex pattern: {e}")

        # Null operators
        elif operator == "is_null":
            mask = col_data.isna()
        elif operator == "is_not_null":
            mask = col_data.notna()

        else:
            raise ValueError(f"Unsupported operator: {operator}")
        
        # Apply negation if requested (NOT operator)
        if filter_cond.negate:
            mask = ~mask
        
        return mask

    def validate_filters(
        self, df: pd.DataFrame, filters: list[FilterCondition | FilterGroup]
    ) -> tuple[bool, str | None]:
        """Validate filters against DataFrame (recursive).

        Args:
            df: DataFrame to validate against
            filters: List of filter conditions or nested groups

        Returns:
            Tuple of (is_valid, error_message)
        """
        for filter_item in filters:
            if isinstance(filter_item, FilterCondition):
                # Validate atomic condition
                result = self._validate_condition(df, filter_item)
                if not result[0]:
                    return result
            elif isinstance(filter_item, FilterGroup):
                # Validate nested group - RECURSION
                result = self.validate_filters(df, filter_item.filters)
                if not result[0]:
                    return result
            else:
                return (False, f"Invalid filter type: {type(filter_item)}")
        
        return (True, None)
    
    def _validate_condition(
        self,
        df: pd.DataFrame,
        filter_cond: FilterCondition
    ) -> tuple[bool, str | None]:
        """Validate single filter condition.
        
        Args:
            df: DataFrame to validate against
            filter_cond: Filter condition to validate
        
        Returns:
            Tuple of (is_valid, error_message)
        """
        # Build normalized column mapping once
        normalized_to_original = {
            self._normalize_column_name(col): col
            for col in df.columns
        }
        
        # Check column exists using normalized comparison
        filter_col_normalized = self._normalize_column_name(filter_cond.column)
        
        if filter_col_normalized not in normalized_to_original:
            # Column not found - provide helpful error with fuzzy matching
            suggestions = get_close_matches(
                filter_col_normalized,
                normalized_to_original.keys(),
                n=3,
                cutoff=0.6
            )
            
            available = ", ".join(str(col) for col in df.columns)
            suggestion_text = ""
            if suggestions:
                # Map back to original names for suggestions
                original_suggestions = [
                    normalized_to_original[s] for s in suggestions
                ]
                suggestion_text = f" Did you mean: {', '.join(repr(s) for s in original_suggestions)}?"
            
            return (
                False,
                f"Column '{filter_cond.column}' not found.{suggestion_text} Available: {available}",
            )

        # Check operator-specific requirements
        if filter_cond.operator in ["in", "not_in"]:
            if not filter_cond.values:
                return (False, f"Operator '{filter_cond.operator}' requires 'values' parameter")

        elif filter_cond.operator in ["contains", "startswith", "endswith", "regex"]:
            if not isinstance(filter_cond.value, str):
                return (False, f"Operator '{filter_cond.operator}' requires string value")

        elif filter_cond.operator in ["is_null", "is_not_null"]:
            # These don't need value
            pass

        else:
            # Comparison operators need value
            if filter_cond.value is None:
                return (False, f"Operator '{filter_cond.operator}' requires 'value' parameter")

        return (True, None)

    def get_filter_summary(self, filters: list[FilterCondition | FilterGroup], logic: str) -> str:
        """Get human-readable summary of filters (recursive).

        Args:
            filters: List of filter conditions or nested groups
            logic: Logic operator

        Returns:
            Human-readable filter description
        """
        if not filters:
            return "No filters applied"

        parts = []
        for filter_item in filters:
            if isinstance(filter_item, FilterCondition):
                # Format atomic condition
                if filter_item.operator in ["in", "not_in"]:
                    value_str = ", ".join(str(v) for v in (filter_item.values or []))
                    condition = f"{filter_item.column} {filter_item.operator} [{value_str}]"
                elif filter_item.operator in ["is_null", "is_not_null"]:
                    condition = f"{filter_item.column} {filter_item.operator}"
                else:
                    condition = f"{filter_item.column} {filter_item.operator} {filter_item.value}"
                
                # Apply negation if requested
                if filter_item.negate:
                    condition = f"NOT ({condition})"
                
                parts.append(condition)
            
            elif isinstance(filter_item, FilterGroup):
                # Format nested group - RECURSION
                group_summary = self.get_filter_summary(filter_item.filters, filter_item.logic)
                group_str = f"({group_summary})"
                
                # Apply negation to entire group if requested
                if filter_item.negate:
                    group_str = f"NOT {group_str}"
                
                parts.append(group_str)

        return f" {logic} ".join(parts)
    
    def _parse_datetime_value(self, value: Any) -> pd.Timestamp:
        """Parse datetime value from filter.
        
        Args:
            value: Value to parse (string, int, float, or datetime)
        
        Returns:
            pd.Timestamp object
        """
        if isinstance(value, pd.Timestamp):
            return value
        elif isinstance(value, str):
            # Parse ISO 8601 string
            return pd.to_datetime(value)
        elif isinstance(value, (int, float)):
            # Parse Excel date number
            return self._datetime_converter.convert_excel_number_to_datetime(value)
        else:
            # Try to convert whatever it is
            return pd.to_datetime(value)
