from unittest import TestCase

from src.application.extract_citations_use_case import ExtractCitationsUseCase
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.citation_extraction_result_dto import CitationExtractionResultDTO
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.enums.citation_type import CitationType
from src.domain.tests.document.fake_citation_extraction_port import FakeCitationExtractionPort
from src.domain.tests.document.fake_reference_extraction_port import FakeReferenceExtractionPort

_CITATION = CitationDTO(text="(Smith, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0)
_REFERENCE = ReferenceDTO(text="Smith, J. (2020). Title. Journal.")


class TestExtractCitationsUseCase(TestCase):
    def test_execute_returns_citation_extraction_result_dto(self):
        use_case = ExtractCitationsUseCase(
            citation_port=FakeCitationExtractionPort(),
            reference_port=FakeReferenceExtractionPort(),
        )

        result = use_case.execute(docx_path="doc.docx")

        self.assertIsInstance(result, CitationExtractionResultDTO)

    def test_execute_maps_citations_references_and_section_type(self):
        use_case = ExtractCitationsUseCase(
            citation_port=FakeCitationExtractionPort(citations=[_CITATION]),
            reference_port=FakeReferenceExtractionPort(result=([_REFERENCE], "Bibliografía")),
        )

        result = use_case.execute(docx_path="doc.docx")

        self.assertEqual(result.citations, [_CITATION])
        self.assertEqual(result.references, [_REFERENCE])
        self.assertEqual(result.section_type, "Bibliografía")
