from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.language_model_errors import LanguageModelError, LanguageModelUnavailable


class TestLanguageModelUnavailable(TestCase):
    def test_is_subclass_of_language_model_error(self):
        self.assertTrue(issubclass(LanguageModelUnavailable, LanguageModelError))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise LanguageModelUnavailable()

    def test_message_is_not_empty(self):
        self.assertGreater(len(LanguageModelUnavailable.MESSAGE), 0)
