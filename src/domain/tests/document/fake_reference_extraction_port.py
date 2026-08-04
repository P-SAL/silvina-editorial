from src.domain.document.reference_extraction_port import ReferenceExtractionPort
from src.domain.dtos.reference_dto import ReferenceDTO


class FakeReferenceExtractionPort(ReferenceExtractionPort):
    """Test double for ReferenceExtractionPort with configurable return or exception."""

    def __init__(
        self,
        result: tuple[list[ReferenceDTO], str] | None = None,
        error: Exception | None = None,
    ) -> None:
        self._result = result if result is not None else ([], "Referencias")
        self._error = error

    def extract_references(self, docx_path: str) -> tuple[list[ReferenceDTO], str]:
        """Return the configured result or raise the configured exception."""
        if self._error is not None:
            raise self._error
        return self._result
