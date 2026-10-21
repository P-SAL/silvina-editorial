from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.reference_errors import ReferenceExtractionError, ReferenceParsingFailed


class TestReferenceError(TestCase):
    def test_s3a_reference_extraction_error_mro_includes_base_src_error(self):
        self.assertTrue(issubclass(ReferenceExtractionError, BaseSrcError))

    def test_s3a_reference_parsing_failed_mro(self):
        mro = ReferenceParsingFailed.__mro__
        names = [cls.__name__ for cls in mro]
        self.assertIn("ReferenceParsingFailed", names)
        self.assertIn("ReferenceExtractionError", names)
        self.assertIn("BaseSrcError", names)
        self.assertIn("Exception", names)
        rp_index = names.index("ReferenceParsingFailed")
        re_index = names.index("ReferenceExtractionError")
        bs_index = names.index("BaseSrcError")
        ex_index = names.index("Exception")
        self.assertLess(rp_index, re_index)
        self.assertLess(re_index, bs_index)
        self.assertLess(bs_index, ex_index)

    def test_s3b_reference_parsing_failed_message_is_non_empty_string(self):
        self.assertIsInstance(ReferenceParsingFailed.MESSAGE, str)
        self.assertGreater(len(ReferenceParsingFailed.MESSAGE), 0)
