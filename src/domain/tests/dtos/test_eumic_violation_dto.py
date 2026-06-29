from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.eumic_violation_dto import EumicViolationDTO
from src.domain.enums.severity_level import SeverityLevel


class TestEumicViolationDTO(TestCase):
    def _make_dto(self) -> EumicViolationDTO:
        return EumicViolationDTO(
            category="Formato General",
            message="Margen incorrecto",
            severity=SeverityLevel.WARNING,
        )

    def test_constructs_with_required_fields(self):
        dto = self._make_dto()
        self.assertEqual(dto.category, "Formato General")
        self.assertEqual(dto.message, "Margen incorrecto")
        self.assertEqual(dto.severity, SeverityLevel.WARNING)

    def test_details_defaults_to_empty_string(self):
        dto = self._make_dto()
        self.assertEqual(dto.details, "")

    def test_details_accepts_non_empty_value(self):
        dto = EumicViolationDTO(
            category="Figuras",
            message="Sin título",
            severity=SeverityLevel.CRITICAL,
            details="2 imágenes sin caption",
        )
        self.assertEqual(dto.details, "2 imágenes sin caption")

    def test_category_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.category = "other"

    def test_message_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.message = "other"

    def test_severity_field_raises_frozen_instance_error_on_mutation(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.severity = SeverityLevel.INFO

    def test_severity_is_severity_level_enum(self):
        dto = self._make_dto()
        self.assertIsInstance(dto.severity, SeverityLevel)

    def test_two_identical_dtos_are_equal(self):
        dto_a = self._make_dto()
        dto_b = self._make_dto()
        self.assertEqual(dto_a, dto_b)
