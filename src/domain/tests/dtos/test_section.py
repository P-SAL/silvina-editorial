from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.section_dto import SectionDTO
from src.domain.enums.section_type import SectionType


class TestSectionDTO(TestCase):
    def test_section_is_subclass_of_base_dto(self):
        self.assertTrue(issubclass(SectionDTO, BaseDTO))

    def test_section_with_empty_title_raises_value_error(self):
        with self.assertRaises(ValueError):
            SectionDTO(title="", content="Some content")

    def test_section_without_section_type_has_section_type_none(self):
        section = SectionDTO(title="Introduction", content="Some content")
        self.assertIsNone(section.section_type)

    def test_section_with_explicit_section_type_preserves_it(self):
        section = SectionDTO(
            title="Introduction",
            content="Some content",
            section_type=SectionType.INTRODUCTION,
        )
        self.assertEqual(section.section_type, SectionType.INTRODUCTION)

    def test_section_is_immutable(self):
        section = SectionDTO(title="Introduction", content="Some content")
        with self.assertRaises(FrozenInstanceError):
            section.title = "Modified"

    def test_section_field_values_match_constructor_arguments(self):
        section = SectionDTO(
            title="Methods",
            content="We used X",
            section_type=SectionType.INTRODUCTION,
            start_position=10,
            end_position=20,
            level=2,
        )
        self.assertEqual(section.title, "Methods")
        self.assertEqual(section.content, "We used X")
        self.assertEqual(section.start_position, 10)
        self.assertEqual(section.end_position, 20)
        self.assertEqual(section.level, 2)
