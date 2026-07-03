from unittest import TestCase

from src.application.validate_apa_use_case import ValidateApaUseCase
from src.domain.citation.apa_validator import ApaValidator
from src.domain.dtos.apa_validation_result_dto import ApaValidationResultDTO


class TestValidateApaUseCase(TestCase):
    def setUp(self):
        self.use_case = ValidateApaUseCase(validator=ApaValidator())

    def test_s12_empty_list_returns_valid_result(self):
        result = self.use_case.execute(citations=[])
        self.assertIsInstance(result, ApaValidationResultDTO)
        self.assertTrue(result.is_valid)
        self.assertEqual(result.violation_count, 0)
        self.assertEqual(result.violations, [])

    def test_s13_list_with_violation_returns_invalid_result(self):
        citations = [("(García & Pérez, 2020)", 1, "")]
        result = self.use_case.execute(citations=citations)
        self.assertFalse(result.is_valid)
        self.assertEqual(result.violation_count, 1)
        self.assertEqual(len(result.violations), 1)
