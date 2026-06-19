from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.enums.apa_error_type import ApaErrorType
from src.domain.dtos.apa_violation_dto import ApaViolationDTO


class TestApaViolationDTO(TestCase):
    def _make_violation(self, **kwargs):
        defaults = {
            "citation_text": "(García, 2020)",
            "error_type": ApaErrorType.COMMA_ERROR,
            "location": 1,
            "explanation": "test",
            "correction": "(García, 2020)",
        }
        defaults.update(kwargs)
        return ApaViolationDTO(**defaults)

    def test_fields_present(self):
        v = self._make_violation()
        self.assertEqual(v.citation_text, "(García, 2020)")
        self.assertEqual(v.error_type, ApaErrorType.COMMA_ERROR)
        self.assertEqual(v.location, 1)
        self.assertEqual(v.explanation, "test")
        self.assertEqual(v.correction, "(García, 2020)")
        self.assertEqual(v.paragraph_preview, "")

    def test_paragraph_preview_default_is_empty_string(self):
        v = self._make_violation()
        self.assertEqual(v.paragraph_preview, "")

    def test_frozen_raises_on_mutation(self):
        v = self._make_violation()
        with self.assertRaises(FrozenInstanceError):
            v.citation_text = "mutated"  # type: ignore[misc]
