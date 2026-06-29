from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler


class VerifyEumicUseCase:
    """Orchestrates EUMIC format compliance verification for a document."""

    def __init__(self, format_inspection_port: DocumentFormatInspectionPort) -> None:
        self._format_inspection_port = format_inspection_port

    @generic_error_handler
    def execute(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        """Return EUMIC format violations found in the document at docx_path."""
        return self._format_inspection_port.inspect(docx_path=docx_path, word_count=word_count)
