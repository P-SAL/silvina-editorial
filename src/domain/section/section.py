from dataclasses import dataclass

from src.domain.entities.base_entity import BaseEntity
from src.domain.enums.section_type import SectionType


@dataclass
class Section(BaseEntity):
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

    def get_word_count(self) -> int:
        """Return the word count of the section content."""
        return len(self.content.split())

    def is_empty(self) -> bool:
        """Return True if the section content has no non-whitespace characters."""
        return len(self.content.strip()) == 0
