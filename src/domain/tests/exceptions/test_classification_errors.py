from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError, SrcBaseWarning
from src.domain.exceptions.classification_errors import ClassificationFailed


class TestClassificationFailed(TestCase):
    def test_is_subclass_of_src_base_warning(self):
        self.assertTrue(issubclass(ClassificationFailed, SrcBaseWarning))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise ClassificationFailed()

    def test_message_is_not_empty(self):
        self.assertGreater(len(ClassificationFailed.MESSAGE), 0)
