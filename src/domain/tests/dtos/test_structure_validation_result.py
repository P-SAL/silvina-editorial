from dataclasses import FrozenInstanceError
from typing import Any, get_type_hints
from unittest import TestCase

from src.domain.dtos.structure_validation_result_dto import StructureValidationResult


class TestStructureValidationResult(TestCase):
    def test_structure_validation_result_is_subclass_of_base_dto(self):
        from src.domain.dtos.base_dto import BaseDTO

        self.assertTrue(issubclass(StructureValidationResult, BaseDTO))

    def test_structure_validation_result_is_immutable(self):
        result = StructureValidationResult(is_valid=True)
        with self.assertRaises(FrozenInstanceError):
            result.is_valid = False

    def test_str_for_valid_structure(self):
        result = StructureValidationResult(is_valid=True, missing_sections=[])
        self.assertEqual(str(result), "Structure: Valid")

    def test_str_for_invalid_structure_with_two_missing(self):
        result = StructureValidationResult(
            is_valid=False,
            missing_sections=["abstract", "conclusion"],
        )
        self.assertEqual(str(result), "Structure: Invalid (2 missing)")

    def test_section_details_annotation_uses_typing_any(self):
        hints = get_type_hints(StructureValidationResult)
        expected = dict[str, dict[str, Any]]
        self.assertEqual(hints["section_details"], expected)
