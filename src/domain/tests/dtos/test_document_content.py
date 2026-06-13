from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.document_content_dto import DocumentContent
from src.domain.dtos.reference_dto import Reference


class TestDocumentContent(TestCase):
    def test_document_content_is_subclass_of_base_dto(self):
        self.assertTrue(issubclass(DocumentContent, BaseDTO))

    def test_document_content_is_immutable(self):
        document = DocumentContent(word_count=100, char_count=500)
        with self.assertRaises(FrozenInstanceError):
            document.word_count = 0

    def test_document_content_field_values_match_constructor_arguments(self):
        document = DocumentContent(word_count=42, char_count=200, paragraph_count=3)
        self.assertEqual(document.word_count, 42)
        self.assertEqual(document.char_count, 200)
        self.assertEqual(document.paragraph_count, 3)

    def test_document_content_optional_fields_default_to_none(self):
        document = DocumentContent(word_count=10, char_count=50)
        self.assertIsNone(document.title)
        self.assertIsNone(document.authors)
        self.assertIsNone(document.abstract)

    def test_document_content_list_fields_default_to_empty(self):
        document = DocumentContent(word_count=10, char_count=50)
        self.assertEqual(document.keywords, [])
        self.assertEqual(document.references, [])
        self.assertEqual(document.paragraphs, [])

    def test_document_content_sections_defaults_to_empty_dict(self):
        document = DocumentContent(word_count=10, char_count=50)
        self.assertEqual(document.sections, {})

    def test_document_content_references_is_list_of_reference_instances(self):
        reference = Reference(text="Some reference")
        document = DocumentContent(word_count=10, char_count=50, references=[reference])
        self.assertEqual(len(document.references), 1)
        self.assertIsInstance(document.references[0], Reference)
