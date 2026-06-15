from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.document_errors import DocumentError, DocumentUnreadable


class TestDocumentUnreadable(TestCase):
    def test_is_subclass_of_document_error(self):
        self.assertTrue(issubclass(DocumentUnreadable, DocumentError))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise DocumentUnreadable()

    def test_message_is_not_empty(self):
        self.assertGreater(len(DocumentUnreadable.MESSAGE), 0)
