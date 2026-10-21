from unittest import TestCase

from src.domain.exceptions.base_src_error import SrcBaseWarning
from src.domain.exceptions.grammar_errors import GrammarCheckUnavailable


class TestGrammarCheckUnavailable(TestCase):
    def test_is_subclass_of_src_base_warning(self):
        self.assertTrue(issubclass(GrammarCheckUnavailable, SrcBaseWarning))

    def test_message_equals_expected_string(self):
        self.assertEqual(
            GrammarCheckUnavailable.MESSAGE,
            "The grammar check service is unavailable.",
        )

    def test_is_catchable_as_src_base_warning(self):
        with self.assertRaises(SrcBaseWarning):
            raise GrammarCheckUnavailable()
