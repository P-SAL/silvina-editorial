from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.citation_errors import CitationError, CitationParsingFailed


class TestCitationParsingFailed(TestCase):
    def test_is_subclass_of_citation_error(self):
        self.assertTrue(issubclass(CitationParsingFailed, CitationError))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise CitationParsingFailed()

    def test_message_is_not_empty(self):
        self.assertGreater(len(CitationParsingFailed.MESSAGE), 0)
