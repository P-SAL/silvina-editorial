from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.reference_dto import ReferenceDTO


class TestReferenceDTO(TestCase):
    def test_reference_is_subclass_of_base_dto(self):
        self.assertTrue(issubclass(ReferenceDTO, BaseDTO))

    def test_reference_instantiation_with_required_field_only(self):
        reference = ReferenceDTO(text="Some reference text")
        self.assertIsNone(reference.authors)
        self.assertIsNone(reference.year)
        self.assertIsNone(reference.title)
        self.assertIsNone(reference.source)

    def test_reference_field_values_match_constructor_arguments(self):
        reference = ReferenceDTO(
            text="Some text",
            authors="Smith",
            year="2020",
            title="A Study",
            source="Journal of X",
        )
        self.assertEqual(reference.text, "Some text")
        self.assertEqual(reference.authors, "Smith")
        self.assertEqual(reference.year, "2020")
        self.assertEqual(reference.title, "A Study")
        self.assertEqual(reference.source, "Journal of X")

    def test_reference_is_immutable(self):
        reference = ReferenceDTO(text="Some text")
        with self.assertRaises(FrozenInstanceError):
            reference.text = "Modified"

    def test_reference_str_returns_formatted_string(self):
        reference = ReferenceDTO(text="Some text", authors="Smith", year="2020")
        self.assertEqual(str(reference), "ReferenceDTO(Smith, 2020)")

    def test_reference_str_when_authors_and_year_are_none(self):
        reference = ReferenceDTO(text="Some text")
        result = str(reference)
        self.assertIsInstance(result, str)
