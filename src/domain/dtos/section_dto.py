from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.section_type import SectionType


@dataclass(frozen=True)
class Section(BaseDTO):
    """Represents a section in an academic document."""

    title: str
    content: str
    section_type: SectionType | None = None
    start_position: int = 0
    end_position: int = 0
    level: int = 1  # Heading level (1, 2, 3, etc.)

    def __post_init__(self) -> None:
        """Validate that section title is not empty."""
        if not self.title:
            raise ValueError("Section title cannot be empty")
