from enum import Enum


class CitationType(Enum):
    """Types of citations found in academic documents."""

    AUTHOR_YEAR = "author_year"  # e.g., (Smith, 2020)
    NUMERIC = "numeric"  # e.g., [1], [2]
    FOOTNOTE = "footnote"  # e.g., superscript numbers
    UNKNOWN = "unknown"
