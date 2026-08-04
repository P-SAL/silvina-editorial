from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.reference_dto import ReferenceDTO


@dataclass(frozen=True)
class CitationExtractionResultDTO(BaseDTO):
    """Holds the full result of citation and reference extraction from a document."""

    citations: list[CitationDTO]
    references: list[ReferenceDTO]
    section_type: str
