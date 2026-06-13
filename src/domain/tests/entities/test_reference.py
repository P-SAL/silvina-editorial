import inspect
from unittest import TestCase

from src.domain.entities.base_entity import BaseEntity


class TestReference(TestCase):
    def _import_reference(self):
        from src.domain.reference.reference import Reference
        return Reference

    def test_reference_is_subclass_of_base_entity(self):
        Reference = self._import_reference()
        self.assertTrue(issubclass(Reference, BaseEntity))

    def test_reference_instantiation_with_required_field_only(self):
        Reference = self._import_reference()
        reference = Reference(text="Some reference text")
        self.assertIsNone(reference.authors)
        self.assertIsNone(reference.year)
        self.assertIsNone(reference.title)
        self.assertIsNone(reference.source)

    def test_reference_str_returns_formatted_string(self):
        Reference = self._import_reference()
        reference = Reference(text="Some text", authors="Smith", year="2020")
        self.assertEqual(str(reference), "Reference(Smith, 2020)")

    def test_reference_str_when_authors_and_year_are_none(self):
        Reference = self._import_reference()
        reference = Reference(text="Some text")
        result = str(reference)
        self.assertIsInstance(result, str)
