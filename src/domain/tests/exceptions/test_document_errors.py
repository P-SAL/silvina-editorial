from unittest import TestCase

from src.domain.exceptions.base_src_error import (
    BaseSrcError,
    SrcBaseNotFound,
    SrcBaseWarning,
)
from src.domain.exceptions.document_errors import (
    DocumentEmpty,
    DocumentNotFound,
    DocumentUnreadable,
)


class TestDocumentNotFound(TestCase):
    def test_is_subclass_of_src_base_not_found(self):
        self.assertTrue(issubclass(DocumentNotFound, SrcBaseNotFound))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise DocumentNotFound()

    def test_message_is_not_empty(self):
        self.assertGreater(len(DocumentNotFound.MESSAGE), 0)


class TestDocumentEmpty(TestCase):
    def test_is_subclass_of_src_base_warning(self):
        self.assertTrue(issubclass(DocumentEmpty, SrcBaseWarning))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise DocumentEmpty()

    def test_message_is_not_empty(self):
        self.assertGreater(len(DocumentEmpty.MESSAGE), 0)


class TestDocumentUnreadable(TestCase):
    def test_is_subclass_of_src_base_warning(self):
        self.assertTrue(issubclass(DocumentUnreadable, SrcBaseWarning))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise DocumentUnreadable()

    def test_message_is_not_empty(self):
        self.assertGreater(len(DocumentUnreadable.MESSAGE), 0)
