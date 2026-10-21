from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError
from src.domain.exceptions.quality_errors import QualityAnalysisFailed, QualityError


class TestQualityAnalysisFailed(TestCase):
    def test_is_subclass_of_quality_error(self):
        self.assertTrue(issubclass(QualityAnalysisFailed, QualityError))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise QualityAnalysisFailed()

    def test_message_is_not_empty(self):
        self.assertGreater(len(QualityAnalysisFailed.MESSAGE), 0)
