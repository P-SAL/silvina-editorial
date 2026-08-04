from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.grammar_error_dto import GrammarErrorDTO


class TestGrammarErrorDTO(TestCase):
    def _make_dto(self) -> GrammarErrorDTO:
        return GrammarErrorDTO(
            number=1,
            message="Test error",
            context="some context",
            offset=5,
            length=4,
            replacements=["fix1", "fix2"],
        )

    def test_constructs_with_six_fields(self):
        dto = self._make_dto()
        self.assertEqual(dto.number, 1)
        self.assertEqual(dto.message, "Test error")
        self.assertEqual(dto.context, "some context")
        self.assertEqual(dto.offset, 5)
        self.assertEqual(dto.length, 4)
        self.assertEqual(dto.replacements, ["fix1", "fix2"])

    def test_number_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.number = 2

    def test_message_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.message = "other"

    def test_replacements_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.replacements = []
