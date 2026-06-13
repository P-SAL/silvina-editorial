from unittest import TestCase

from src.domain.entities.base_entity import BaseEntity
from src.domain.enums.section_type import SectionType


class TestSection(TestCase):
    def _import_section(self):
        from src.domain.section.section import Section

        return Section

    def test_section_is_subclass_of_base_entity(self):
        Section = self._import_section()
        self.assertTrue(issubclass(Section, BaseEntity))

    def test_section_with_empty_title_raises_value_error(self):
        Section = self._import_section()
        with self.assertRaises(ValueError):
            Section(title="", content="Some content")

    def test_section_without_section_type_has_section_type_none(self):
        Section = self._import_section()
        section = Section(title="Introduction", content="Some content")
        self.assertIsNone(section.section_type)

    def test_section_with_explicit_section_type_preserves_it(self):
        Section = self._import_section()
        section = Section(
            title="Introduction",
            content="Some content",
            section_type=SectionType.INTRODUCTION,
        )
        self.assertEqual(section.section_type, SectionType.INTRODUCTION)

    def test_get_word_count_returns_word_count_of_content(self):
        Section = self._import_section()
        section = Section(title="Intro", content="one two three four")
        self.assertEqual(section.get_word_count(), 4)

    def test_is_empty_returns_true_for_blank_content(self):
        Section = self._import_section()
        section = Section(title="Intro", content="   ")
        self.assertTrue(section.is_empty())

    def test_is_empty_returns_false_for_non_blank_content(self):
        Section = self._import_section()
        section = Section(title="Intro", content="Hello world")
        self.assertFalse(section.is_empty())
