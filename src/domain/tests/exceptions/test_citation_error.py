from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.citation_errors import CitationError


class TestCitationError(TestCase):
    def test_is_subclass_of_base_src_error(self):
        self.assertTrue(issubclass(CitationError, BaseSrcError))
