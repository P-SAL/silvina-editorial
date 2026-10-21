from src.domain.document.citation_extraction_port import CitationExtractionPort
from src.domain.dtos.citation_dto import CitationDTO


class FakeCitationExtractionPort(CitationExtractionPort):
    """Test double for CitationExtractionPort with configurable return or exception."""

    def __init__(
        self,
        citations: list[CitationDTO] | None = None,
        error: Exception | None = None,
    ) -> None:
        self._citations = citations if citations is not None else []
        self._error = error

    def extract_citations(self, docx_path: str) -> list[CitationDTO]:
        """Return the configured citations or raise the configured exception."""
        if self._error is not None:
            raise self._error
        return self._citations
