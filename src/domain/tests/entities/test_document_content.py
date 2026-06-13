import inspect
from unittest import TestCase

from src.domain.entities.base_entity import BaseEntity


class TestDocumentContent(TestCase):
    def _import_document_content(self):
        from src.domain.document.document_content import DocumentContent
        return DocumentContent

    def _import_reference(self):
        from src.domain.reference.reference import Reference
        return Reference

    def test_document_content_is_subclass_of_base_entity(self):
        DocumentContent = self._import_document_content()
        self.assertTrue(issubclass(DocumentContent, BaseEntity))

    def test_document_content_computes_word_count_from_paragraphs_when_zero(self):
        DocumentContent = self._import_document_content()
        document = DocumentContent(
            word_count=0,
            char_count=100,
            paragraphs=["hello world", "foo"],
        )
        self.assertEqual(document.word_count, 3)

    def test_document_content_preserves_explicit_word_count(self):
        DocumentContent = self._import_document_content()
        document = DocumentContent(
            word_count=42,
            char_count=200,
            paragraphs=["hello world"],
        )
        self.assertEqual(document.word_count, 42)

    def test_document_content_field_types_use_modern_syntax(self):
        DocumentContent = self._import_document_content()
        source = inspect.getsource(DocumentContent)
        self.assertNotIn("List[", source)
        self.assertNotIn("Dict[", source)
        self.assertNotIn("Optional[", source)

    def test_document_content_references_is_list_of_reference_instances(self):
        DocumentContent = self._import_document_content()
        Reference = self._import_reference()
        ref = Reference(text="Some reference")
        document = DocumentContent(
            word_count=10,
            char_count=50,
            references=[ref],
        )
        self.assertEqual(len(document.references), 1)
        self.assertIsInstance(document.references[0], Reference)
