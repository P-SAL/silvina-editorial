from abc import ABC, abstractmethod

from src.domain.dtos.eumic_violation_dto import EumicViolationDTO


class DocumentFormatInspectionPort(ABC):
    """Port for inspecting a document's format compliance against EUMIC standards."""

    @abstractmethod
    def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        """Return EUMIC format violations found in the document."""
