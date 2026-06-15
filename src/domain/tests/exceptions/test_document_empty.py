from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.document_errors import DocumentEmpty, DocumentError


class TestDocumentEmpty(TestCase):
    def test_is_subclass_of_document_error(self):
        self.assertTrue(issubclass(DocumentEmpty, DocumentError))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise DocumentEmpty()

    def test_message_is_not_empty(self):
        self.assertGreater(len(DocumentEmpty.MESSAGE), 0)
