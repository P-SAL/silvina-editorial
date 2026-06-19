from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.enums.citation_type import CitationType


class TestCitationDTO(TestCase):
    def test_citation_is_subclass_of_base_dto(self):
        self.assertTrue(issubclass(CitationDTO, BaseDTO))

    def test_citation_instantiation_with_required_fields_only(self):
        citation = CitationDTO(text="Some text", citation_type=CitationType.AUTHOR_YEAR, location=0)
        self.assertIsNone(citation.author)
        self.assertIsNone(citation.year)

    def test_citation_field_values_match_constructor_arguments(self):
        citation = CitationDTO(
            text="Some text",
            citation_type=CitationType.NUMERIC,
            location=1,
            author="Smith",
            year="2020",
        )
        self.assertEqual(citation.text, "Some text")
        self.assertEqual(citation.citation_type, CitationType.NUMERIC)
        self.assertEqual(citation.location, 1)
        self.assertEqual(citation.author, "Smith")
        self.assertEqual(citation.year, "2020")

    def test_citation_is_immutable(self):
        citation = CitationDTO(text="Some text", citation_type=CitationType.FOOTNOTE, location=0)
        with self.assertRaises(FrozenInstanceError):
            citation.text = "Modified"

    def test_citation_as_dict_contains_expected_keys(self):
        citation = CitationDTO(text="Some text", citation_type=CitationType.NUMERIC, location=1)
        result = citation.as_dict()
        self.assertIn("text", result)
        self.assertIn("citation_type", result)
        self.assertIn("location", result)
        self.assertIn("author", result)
        self.assertIn("year", result)

    def test_citation_str_truncates_at_50_chars(self):
        long_text = "A" * 60
        citation = CitationDTO(text=long_text, citation_type=CitationType.FOOTNOTE, location=2)
        result = str(citation)
        self.assertTrue(result.startswith("CitationDTO("))
        self.assertIn("...", result)
