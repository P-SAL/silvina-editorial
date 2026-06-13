from dataclasses import dataclass, field

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.reference_dto import Reference


@dataclass(frozen=True)
class DocumentContent(BaseDTO):
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
