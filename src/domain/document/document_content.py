from dataclasses import dataclass, field

from src.domain.entities.base_entity import BaseEntity
from src.domain.reference.reference import Reference


@dataclass
class DocumentContent(BaseEntity):
    """Represents the extracted content of a document."""

    word_count: int
    char_count: int
    paragraph_count: int = 0
    title: str | None = None
    authors: str | None = None
    abstract: str | None = None
    keywords: list[str] = field(default_factory=list)
    references: list[Reference] = field(default_factory=list)
    paragraphs: list[str] = field(default_factory=list)
    sections: dict[str, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        """Compute word count from paragraphs when word_count is zero."""
        if self.word_count == 0 and self.paragraphs:
            self.word_count = sum(len(paragraph.split()) for paragraph in self.paragraphs)
