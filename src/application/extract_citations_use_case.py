from src.domain.document.citation_extraction_port import CitationExtractionPort
from src.domain.document.reference_extraction_port import ReferenceExtractionPort
from src.domain.dtos.citation_extraction_result_dto import CitationExtractionResultDTO
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler


class ExtractCitationsUseCase:
    def __init__(
        self,
        citation_port: CitationExtractionPort,
        reference_port: ReferenceExtractionPort,
    ) -> None:
        self._citation_port = citation_port
        self._reference_port = reference_port

    @generic_error_handler
    def execute(self, docx_path: str) -> CitationExtractionResultDTO:
        citations = self._citation_port.extract_citations(docx_path=docx_path)
        references, section_type = self._reference_port.extract_references(docx_path=docx_path)
        return CitationExtractionResultDTO(
            citations=citations,
            references=references,
            section_type=section_type,
        )
