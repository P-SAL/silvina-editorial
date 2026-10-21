from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO


class DocumentFormatInspector:
    """Domain service that wraps document format inspection."""

    def __init__(self, document_format_inspection_port: DocumentFormatInspectionPort) -> None:
        self._document_format_inspection_port = document_format_inspection_port

    def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        """Inspect the document for EUMIC editorial standard violations."""
        return self._document_format_inspection_port.inspect(
            docx_path=docx_path, word_count=word_count
        )
