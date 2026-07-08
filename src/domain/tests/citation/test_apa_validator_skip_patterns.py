from unittest import TestCase

from src.domain.citation.apa_validator import ApaValidator
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.enums.apa_error_type import ApaErrorType
from src.domain.enums.citation_type import CitationType


class TestApaValidatorSkipPatterns(TestCase):
    def setUp(self):
        self.validator = ApaValidator()

    def test_s11_acronym_organization_no_capitalization_or_comma_error(self):
        violations = self.validator.validate_citation("(UNESCO 2020)", 0)
        error_types = [v.error_type for v in violations]
        self.assertNotIn(ApaErrorType.CAPITALIZATION_ERROR, error_types)
        self.assertNotIn(ApaErrorType.COMMA_ERROR, error_types)

    def test_s11b_arxiv_no_violations(self):
        violations = self.validator.validate_citation("(arXiv:2404.19573)", 0)
        self.assertEqual(violations, [])

    def test_s11c_doi_no_violations(self):
        violations = self.validator.validate_citation("(doi:10.1234/foo)", 0)
        self.assertEqual(violations, [])

    def test_repositorio_no_violations(self):
        violations = self.validator.validate_citation("(repositorio trazable)", 0)
        self.assertEqual(violations, [])

    def test_no_hay_dataset_no_violations(self):
        violations = self.validator.validate_citation("(no hay dataset)", 0)
        self.assertEqual(violations, [])

    def test_lowercase_start_two_years_no_violations(self):
        violations = self.validator.validate_citation("(años 2024, 2025)", 0)
        self.assertEqual(violations, [])

    def test_multiword_two_years_no_violations(self):
        violations = self.validator.validate_citation("(some word 2020 2021)", 0)
        self.assertEqual(violations, [])

    def test_validate_all_citations_only_processes_author_year_type(self):
        author_year = CitationDTO(
            text="(Smith & Jones, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0
        )
        numeric = CitationDTO(
            text="(Doe & Roe, 1999)", citation_type=CitationType.NUMERIC, location=1
        )
        paragraphs = ["Paragraph 0 text contents", "Paragraph 1 text contents"]

        violations = self.validator.validate_all_citations(
            citations=[author_year, numeric], paragraphs=paragraphs
        )

        self.assertEqual([v.citation_text for v in violations], ["(Smith & Jones, 2020)"])

    def test_validate_all_citations_builds_preview_from_paragraph_at_location(self):
        citation = CitationDTO(
            text="(García & Pérez, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0
        )
        paragraphs = ["Paragraph 0 text contents"]

        violations = self.validator.validate_all_citations(
            citations=[citation], paragraphs=paragraphs
        )

        self.assertTrue(violations)
        self.assertEqual(violations[0].paragraph_preview, "Paragraph 0 text contents")

    def test_validate_all_citations_falls_back_to_empty_string_when_location_out_of_bounds(self):
        citation = CitationDTO(
            text="(García & Pérez, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=5
        )
        paragraphs = ["Only paragraph"]

        violations = self.validator.validate_all_citations(
            citations=[citation], paragraphs=paragraphs
        )

        self.assertTrue(violations)
        self.assertEqual(violations[0].paragraph_preview, "")

    def test_validate_all_citations_returns_empty_list_for_empty_citations(self):
        violations = self.validator.validate_all_citations(citations=[], paragraphs=[])
        self.assertEqual(violations, [])
