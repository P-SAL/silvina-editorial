from unittest import TestCase

from src.domain.citation.citation_extractor import CitationExtractor
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.enums.citation_type import CitationType
from src.domain.tests.document.fake_citation_extraction_port import FakeCitationExtractionPort
from src.domain.tests.document.fake_reference_extraction_port import FakeReferenceExtractionPort


class TestCitationExtractor(TestCase):
    def test_extract_citations_and_references_returns_expected_tuple(self):
        citation = CitationDTO(
            text="(Smith, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0
        )
        reference = ReferenceDTO(text="Smith, J. (2020). Title.")
        citation_extraction_port = FakeCitationExtractionPort(citations=[citation])
        reference_extraction_port = FakeReferenceExtractionPort(result=([reference], "Referencias"))

        extractor = CitationExtractor(
            citation_extraction_port=citation_extraction_port,
            reference_extraction_port=reference_extraction_port,
        )
        citations, references, section_type = extractor.extract_citations_and_references(
            docx_path="test.docx"
        )

        self.assertEqual(citations, [citation])
        self.assertEqual(references, [reference])
        self.assertEqual(section_type, "Referencias")

    def test_extract_citations_and_references_returns_empty_when_none_found(self):
        citation_extraction_port = FakeCitationExtractionPort(citations=[])
        reference_extraction_port = FakeReferenceExtractionPort(result=([], "Referencias"))

        extractor = CitationExtractor(
            citation_extraction_port=citation_extraction_port,
            reference_extraction_port=reference_extraction_port,
        )
        citations, references, section_type = extractor.extract_citations_and_references(
            docx_path="test.docx"
        )

        self.assertEqual(citations, [])
        self.assertEqual(references, [])
        self.assertEqual(section_type, "Referencias")
