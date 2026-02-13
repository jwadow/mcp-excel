"""TSV formatter for Excel copy-paste functionality."""

from typing import Any


class TSVFormatter:
    """Formats data as TSV for Excel copy-paste."""

    def format_table(
        self, headers: list[str], rows: list[list[Any]]
    ) -> str:
        """Format data as TSV table.

        Args:
            headers: Column headers
            rows: Data rows

        Returns:
            TSV-formatted string
        """
        lines = []

        # Add headers
        lines.append("\t".join(str(h) for h in headers))

        # Add rows
        for row in rows:
            lines.append("\t".join(self._format_cell(cell) for cell in row))

        return "\n".join(lines)

    def format_single_value(
        self, label: str, value: Any, formula: str | None = None
    ) -> str:
        """Format single value with optional formula.

        Args:
            label: Label for the value
            value: The value
            formula: Optional Excel formula

        Returns:
            TSV-formatted string
        """
        if formula:
            return f"{label}\t{formula}"
        else:
            return f"{label}\t{self._format_cell(value)}"

    def format_key_value_pairs(
        self, pairs: dict[str, Any]
    ) -> str:
        """Format dictionary as key-value TSV.

        Args:
            pairs: Dictionary of key-value pairs

        Returns:
            TSV-formatted string
        """
        lines = []
        for key, value in pairs.items():
            lines.append(f"{key}\t{self._format_cell(value)}")
        return "\n".join(lines)

    def format_matrix(
        self, row_labels: list[str], col_labels: list[str], data: list[list[Any]]
    ) -> str:
        """Format matrix with row and column labels.

        Args:
            row_labels: Labels for rows
            col_labels: Labels for columns
            data: Matrix data

        Returns:
            TSV-formatted string
        """
        lines = []

        # Header row with column labels
        header = [""] + col_labels
        lines.append("\t".join(str(h) for h in header))

        # Data rows with row labels
        for row_label, row_data in zip(row_labels, data):
            row = [row_label] + [self._format_cell(cell) for cell in row_data]
            lines.append("\t".join(row))

        return "\n".join(lines)

    def _format_cell(self, value: Any) -> str:
        """Format single cell value.

        Args:
            value: Cell value

        Returns:
            Formatted string
        """
        if value is None:
            return ""
        elif isinstance(value, bool):
            return "TRUE" if value else "FALSE"
        elif isinstance(value, (int, float)):
            return str(value)
        elif isinstance(value, str):
            # Escape tabs and newlines
            escaped = value.replace("\t", " ").replace("\n", " ").replace("\r", "")
            return escaped
        else:
            return str(value)

    def _escape_formula(self, formula: str) -> str:
        """Ensure formula starts with = and is properly formatted.

        Args:
            formula: Excel formula

        Returns:
            Properly formatted formula
        """
        formula = formula.strip()
        if not formula.startswith("="):
            formula = "=" + formula
        return formula
