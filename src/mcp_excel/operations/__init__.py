"""Operations: inspection, filtering, aggregation, statistics, multi-sheet."""

from .data_operations import DataOperations
from .filtering import FilterEngine
from .inspection import InspectionOperations

__all__ = [
    "DataOperations",
    "FilterEngine",
    "InspectionOperations",
]
