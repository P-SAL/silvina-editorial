from src.application.extract_citations_use_case import ExtractCitationsUseCase
from src.domain.document.citation_extraction_port import CitationExtractionPort
from src.domain.document.document_text_port import DocumentTextPort
from src.domain.document.reference_extraction_port import ReferenceExtractionPort
from src.infrastructure.adapters.document.docx_citation_adapter import DocxCitationAdapter
from src.infrastructure.adapters.document.docx_reference_adapter import DocxReferenceAdapter
from src.infrastructure.adapters.document.docx_text_adapter import DocxTextAdapter


class ExtractCitationsUseCaseWiring:
    def create_use_case(self) -> ExtractCitationsUseCase:
        return ExtractCitationsUseCase(
            citation_port=self._get_citation_port(),
            reference_port=self._get_reference_port(),
        )

    def _get_citation_port(self) -> CitationExtractionPort:
        return DocxCitationAdapter(document_text_port=self._get_document_text_port())

    def _get_reference_port(self) -> ReferenceExtractionPort:
        return DocxReferenceAdapter(document_text_port=self._get_document_text_port())

    def _get_document_text_port(self) -> DocumentTextPort:
        return DocxTextAdapter()
