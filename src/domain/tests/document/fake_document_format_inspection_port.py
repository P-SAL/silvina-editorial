from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO


class FakeDocumentFormatInspectionPort(DocumentFormatInspectionPort):
    """Configurable fake for DocumentFormatInspectionPort used in application-layer tests."""

    def __init__(
        self,
        violations: list[EumicViolationDTO] | None = None,
        error: Exception | None = None,
    ) -> None:
        self._violations = violations if violations is not None else []
        self._error = error

    def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        """Return configured violations or raise configured error."""
        if self._error is not None:
            raise self._error
        return self._violations
