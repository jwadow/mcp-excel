"""Filtering system for DataFrame operations."""

import re
from typing import Any

import pandas as pd

from ..models.requests import FilterCondition


class FilterEngine:
    """Engine for applying filters to DataFrames."""

    def apply_filters(
        self,
        df: pd.DataFrame,
        filters: list[FilterCondition],
        logic: str = "AND",
    ) -> pd.DataFrame:
        """Apply filters to DataFrame.

        Args:
            df: DataFrame to filter
            filters: List of filter conditions
            logic: Logic operator ("AND" or "OR")

        Returns:
            Filtered DataFrame

        Raises:
            ValueError: If filter is invalid
        """
        if not filters:
            return df

        # Build mask for each filter
        masks = []
        for filter_cond in filters:
            mask = self._build_filter_mask(df, filter_cond)
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

        return df[combined_mask]

    def count_filtered(
        self,
        df: pd.DataFrame,
        filters: list[FilterCondition],
        logic: str = "AND",
    ) -> int:
        """Count rows matching filters without materializing DataFrame.

        Args:
            df: DataFrame to filter
            filters: List of filter conditions
            logic: Logic operator ("AND" or "OR")

        Returns:
            Count of matching rows
        """
        if not filters:
            return len(df)

        masks = []
        for filter_cond in filters:
            mask = self._build_filter_mask(df, filter_cond)
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
        column = filter_cond.column

        # Check if column exists
        if column not in df.columns:
            available = ", ".join(df.columns.tolist())
            raise ValueError(
                f"Column '{column}' not found. Available columns: {available}"
            )

        col_data = df[column]
        operator = filter_cond.operator

        # Comparison operators
        if operator == "==":
            return col_data == filter_cond.value
        elif operator == "!=":
            return col_data != filter_cond.value
        elif operator == ">":
            return col_data > filter_cond.value
        elif operator == "<":
            return col_data < filter_cond.value
        elif operator == ">=":
            return col_data >= filter_cond.value
        elif operator == "<=":
            return col_data <= filter_cond.value

        # Set operators
        elif operator == "in":
            if not filter_cond.values:
                raise ValueError("'in' operator requires 'values' parameter")
            return col_data.isin(filter_cond.values)
        elif operator == "not_in":
            if not filter_cond.values:
                raise ValueError("'not_in' operator requires 'values' parameter")
            return ~col_data.isin(filter_cond.values)

        # String operators
        elif operator == "contains":
            if not isinstance(filter_cond.value, str):
                raise ValueError("'contains' operator requires string value")
            return col_data.astype(str).str.contains(filter_cond.value, na=False, regex=False)
        elif operator == "startswith":
            if not isinstance(filter_cond.value, str):
                raise ValueError("'startswith' operator requires string value")
            return col_data.astype(str).str.startswith(filter_cond.value, na=False)
        elif operator == "endswith":
            if not isinstance(filter_cond.value, str):
                raise ValueError("'endswith' operator requires string value")
            return col_data.astype(str).str.endswith(filter_cond.value, na=False)
        elif operator == "regex":
            if not isinstance(filter_cond.value, str):
                raise ValueError("'regex' operator requires string value")
            try:
                return col_data.astype(str).str.contains(filter_cond.value, na=False, regex=True)
            except re.error as e:
                raise ValueError(f"Invalid regex pattern: {e}")

        # Null operators
        elif operator == "is_null":
            return col_data.isna()
        elif operator == "is_not_null":
            return col_data.notna()

        else:
            raise ValueError(f"Unsupported operator: {operator}")

    def validate_filters(
        self, df: pd.DataFrame, filters: list[FilterCondition]
    ) -> tuple[bool, str | None]:
        """Validate filters against DataFrame.

        Args:
            df: DataFrame to validate against
            filters: List of filter conditions

        Returns:
            Tuple of (is_valid, error_message)
        """
        for filter_cond in filters:
            # Check column exists
            if filter_cond.column not in df.columns:
                available = ", ".join(df.columns.tolist())
                return (
                    False,
                    f"Column '{filter_cond.column}' not found. Available: {available}",
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

    def get_filter_summary(self, filters: list[FilterCondition], logic: str) -> str:
        """Get human-readable summary of filters.

        Args:
            filters: List of filter conditions
            logic: Logic operator

        Returns:
            Human-readable filter description
        """
        if not filters:
            return "No filters applied"

        parts = []
        for f in filters:
            if f.operator in ["in", "not_in"]:
                value_str = f", ".join(str(v) for v in (f.values or []))
                parts.append(f"{f.column} {f.operator} [{value_str}]")
            elif f.operator in ["is_null", "is_not_null"]:
                parts.append(f"{f.column} {f.operator}")
            else:
                parts.append(f"{f.column} {f.operator} {f.value}")

        return f" {logic} ".join(parts)
