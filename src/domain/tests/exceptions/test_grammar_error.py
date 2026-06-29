from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.grammar_errors import GrammarError


class TestGrammarError(TestCase):
    def test_is_subclass_of_base_src_error(self):
        self.assertTrue(issubclass(GrammarError, BaseSrcError))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise GrammarError()
