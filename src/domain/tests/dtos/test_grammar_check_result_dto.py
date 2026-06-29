from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.dtos.grammar_error_dto import GrammarErrorDTO


class TestGrammarCheckResultDTO(TestCase):
    def _make_error(self) -> GrammarErrorDTO:
        return GrammarErrorDTO(
            number=1,
            message="Test error",
            context="context",
            offset=0,
            length=4,
            replacements=[],
        )

    def _make_dto(self) -> GrammarCheckResultDTO:
        return GrammarCheckResultDTO(
            score=8.5,
            feedback="Pocos errores",
            errors=[self._make_error()],
        )

    def test_constructs_with_three_fields(self):
        dto = self._make_dto()
        self.assertEqual(dto.score, 8.5)
        self.assertEqual(dto.feedback, "Pocos errores")
        self.assertEqual(len(dto.errors), 1)

    def test_score_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.score = 10.0

    def test_feedback_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.feedback = "Sin errores"

    def test_errors_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.errors = []
