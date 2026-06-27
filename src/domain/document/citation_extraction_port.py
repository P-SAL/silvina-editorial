from abc import ABC, abstractmethod

from src.domain.dtos.citation_dto import CitationDTO


class CitationExtractionPort(ABC):
    """Port for extracting citations from a document."""

    @abstractmethod
    def extract_citations(self, docx_path: str) -> list[CitationDTO]:
        """Extract all citations found in the document."""
