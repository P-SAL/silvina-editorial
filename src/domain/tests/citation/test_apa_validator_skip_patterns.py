from unittest import TestCase

from src.domain.citation.apa_validator import ApaValidator
from src.domain.enums.apa_error_type import ApaErrorType


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
