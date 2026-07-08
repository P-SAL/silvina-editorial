from src.domain.document.citation_extraction_port import CitationExtractionPort
from src.domain.document.reference_extraction_port import ReferenceExtractionPort
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.reference_dto import ReferenceDTO


class CitationExtractor:
    """Domain service that extracts citations and references from a document."""

    def __init__(
        self,
        citation_extraction_port: CitationExtractionPort,
        reference_extraction_port: ReferenceExtractionPort,
    ) -> None:
        self._citation_extraction_port = citation_extraction_port
        self._reference_extraction_port = reference_extraction_port

    def extract_citations_and_references(
        self, docx_path: str
    ) -> tuple[list[CitationDTO], list[ReferenceDTO], str]:
        """Extract citations and references and return them alongside the section type."""
        citations = self._citation_extraction_port.extract_citations(docx_path=docx_path)
        references, section_type = self._reference_extraction_port.extract_references(
            docx_path=docx_path
        )
        return citations, references, section_type
