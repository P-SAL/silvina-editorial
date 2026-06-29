from src.application.verify_eumic_use_case import VerifyEumicUseCase
from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.infrastructure.adapters.document.docx_eumic_adapter import DocxEumicAdapter


class VerifyEumicUseCaseWiring:
    """Factory for building a ready-to-use VerifyEumicUseCase."""

    def create_use_case(self) -> VerifyEumicUseCase:
        """Return a fully assembled VerifyEumicUseCase."""
        return VerifyEumicUseCase(
            format_inspection_port=self._get_document_format_inspection_port()
        )

    def _get_document_format_inspection_port(self) -> DocumentFormatInspectionPort:
        """Return the DocxEumicAdapter as the document format inspection port."""
        return DocxEumicAdapter()
