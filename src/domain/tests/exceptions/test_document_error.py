from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.document_errors import DocumentError


class TestDocumentError(TestCase):
    def test_is_subclass_of_base_src_error(self):
        self.assertTrue(issubclass(DocumentError, BaseSrcError))
