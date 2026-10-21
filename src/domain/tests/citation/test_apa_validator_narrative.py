from unittest import TestCase

from src.domain.citation.apa_validator import ApaValidator
from src.domain.enums.apa_error_type import ApaErrorType


class TestApaValidatorNarrative(TestCase):
    def setUp(self):
        self.validator = ApaValidator()

    def test_s08_conjunction_error_ampersand_narrative(self):
        violations = self.validator.validate_citation("García & Pérez (2020)", 1)
        error_types = [v.error_type for v in violations]
        self.assertIn(ApaErrorType.CONJUNCTION_ERROR, error_types)

    def test_s09_et_al_format_error_narrative(self):
        violations = self.validator.validate_citation("García et. al. (2020)", 2)
        error_types = [v.error_type for v in violations]
        self.assertIn(ApaErrorType.ET_AL_FORMAT_ERROR, error_types)

    def test_s10_spacing_error_missing_space_before_year(self):
        violations = self.validator.validate_citation("García(2020)", 3)
        error_types = [v.error_type for v in violations]
        self.assertIn(ApaErrorType.SPACING_ERROR, error_types)
