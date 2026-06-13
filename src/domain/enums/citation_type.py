from enum import Enum


class CitationType(Enum):
    """Types of citations found in academic documents."""

    AUTHOR_YEAR = "author_year"
    NUMERIC = "numeric"
    FOOTNOTE = "footnote"
    UNKNOWN = "unknown"
