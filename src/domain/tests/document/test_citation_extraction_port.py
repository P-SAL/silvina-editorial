from inspect import signature
from unittest import TestCase

from src.domain.document.citation_extraction_port import CitationExtractionPort
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.enums.citation_type import CitationType
from src.domain.tests.document.fake_citation_extraction_port import FakeCitationExtractionPort


class TestCitationExtractionPort(TestCase):
    def test_s1a_cannot_instantiate_abstract_class(self):
        with self.assertRaises(TypeError):
            CitationExtractionPort()

    def test_s1b_extract_citations_has_docx_path_str_parameter(self):
        sig = signature(CitationExtractionPort.extract_citations)
        self.assertIn("docx_path", sig.parameters)
        self.assertEqual(sig.parameters["docx_path"].annotation, str)

    def test_s1b_extract_citations_returns_list_of_citation_dto(self):
        sig = signature(CitationExtractionPort.extract_citations)
        self.assertEqual(sig.return_annotation, list[CitationDTO])

    def test_s5a_fake_returns_configured_citations(self):
        citation = CitationDTO(
            text="(Smith, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0
        )
        fake = FakeCitationExtractionPort(citations=[citation])
        result = fake.extract_citations(docx_path="test.docx")
        self.assertEqual(result, [citation])

    def test_s5b_fake_raises_configured_exception(self):
        error = ValueError("extraction failed")
        fake = FakeCitationExtractionPort(error=error)
        with self.assertRaises(ValueError):
            fake.extract_citations(docx_path="test.docx")
