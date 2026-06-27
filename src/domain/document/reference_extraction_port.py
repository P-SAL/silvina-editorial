from abc import ABC, abstractmethod

from src.domain.dtos.reference_dto import ReferenceDTO


class ReferenceExtractionPort(ABC):
    """Port for extracting references from a document."""

    @abstractmethod
    def extract_references(self, docx_path: str) -> tuple[list[ReferenceDTO], str]:
        """Extract all references and the section title from the document."""
