from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.citation_type import CitationType


@dataclass(frozen=True)
class CitationDTO(BaseDTO):
    """Represents a citation found in the document."""

    text: str
    citation_type: CitationType
    location: int  # Paragraph index where found
    author: str | None = None
    year: str | None = None

    def __str__(self) -> str:
        """Return a short string representation truncated at 50 characters."""
        return f"CitationDTO({self.text[:50]}...)"
