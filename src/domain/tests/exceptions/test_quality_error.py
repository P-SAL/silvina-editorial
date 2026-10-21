from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.quality_errors import QualityError


class TestQualityError(TestCase):
    def test_is_subclass_of_base_src_error(self):
        self.assertTrue(issubclass(QualityError, BaseSrcError))
