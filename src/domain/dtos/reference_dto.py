from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class Reference(BaseDTO):
    """Represents a reference in the bibliography."""

    text: str
    authors: str | None = None
    year: str | None = None
    title: str | None = None
    source: str | None = None

    def __str__(self) -> str:
        """Return a formatted string with authors and year."""
        return f"Reference({self.authors}, {self.year})"
