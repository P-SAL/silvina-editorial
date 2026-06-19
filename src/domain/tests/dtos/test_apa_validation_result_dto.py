from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.enums.apa_error_type import ApaErrorType
from src.domain.dtos.apa_violation_dto import ApaViolationDTO
from src.domain.dtos.apa_validation_result_dto import ApaValidationResultDTO


class TestApaValidationResultDTO(TestCase):
    def test_valid_result_fields(self):
        result = ApaValidationResultDTO(is_valid=True, violation_count=0, violations=[])
        self.assertTrue(result.is_valid)
        self.assertEqual(result.violation_count, 0)
        self.assertEqual(result.violations, [])

    def test_invalid_result_fields(self):
        v = ApaViolationDTO(
            citation_text="(García & Pérez, 2020)",
            error_type=ApaErrorType.CONJUNCTION_ERROR,
            location=1,
            explanation="APA 7 requiere y",
            correction="(García y Pérez, 2020)",
        )
        result = ApaValidationResultDTO(is_valid=False, violation_count=1, violations=[v])
        self.assertFalse(result.is_valid)
        self.assertEqual(result.violation_count, 1)
        self.assertEqual(len(result.violations), 1)

    def test_frozen_raises_on_mutation(self):
        result = ApaValidationResultDTO(is_valid=True, violation_count=0, violations=[])
        with self.assertRaises(FrozenInstanceError):
            result.is_valid = False  # type: ignore[misc]

    def test_is_valid_invariant_zero_count(self):
        result = ApaValidationResultDTO(is_valid=True, violation_count=0, violations=[])
        self.assertTrue(result.is_valid == (result.violation_count == 0))

    def test_is_valid_invariant_nonzero_count(self):
        v = ApaViolationDTO(
            citation_text="(x, 2020)",
            error_type=ApaErrorType.CAPITALIZATION_ERROR,
            location=0,
            explanation="x",
            correction="(X, 2020)",
        )
        result = ApaValidationResultDTO(is_valid=False, violation_count=1, violations=[v])
        self.assertTrue(result.is_valid == (result.violation_count == 0))
