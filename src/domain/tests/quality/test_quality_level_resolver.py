from unittest import TestCase

from src.domain.enums.quality_level import QualityLevel
from src.domain.quality.quality_level_resolver import QualityLevelResolver


class TestQualityLevelResolver(TestCase):
    def setUp(self):
        self.resolver = QualityLevelResolver()

    def test_score_of_nine_point_zero_returns_excellent(self):
        self.assertEqual(self.resolver.resolve(9.0), QualityLevel.EXCELLENT)

    def test_score_just_below_nine_returns_good(self):
        self.assertEqual(self.resolver.resolve(8.9), QualityLevel.GOOD)

    def test_score_of_seven_point_zero_returns_good(self):
        self.assertEqual(self.resolver.resolve(7.0), QualityLevel.GOOD)

    def test_score_just_below_seven_returns_acceptable(self):
        self.assertEqual(self.resolver.resolve(6.9), QualityLevel.ACCEPTABLE)

    def test_score_of_five_point_zero_returns_acceptable(self):
        self.assertEqual(self.resolver.resolve(5.0), QualityLevel.ACCEPTABLE)

    def test_score_just_below_five_returns_needs_improvement(self):
        self.assertEqual(self.resolver.resolve(4.9), QualityLevel.NEEDS_IMPROVEMENT)

    def test_score_of_three_point_zero_returns_needs_improvement(self):
        self.assertEqual(self.resolver.resolve(3.0), QualityLevel.NEEDS_IMPROVEMENT)

    def test_score_just_below_three_returns_poor(self):
        self.assertEqual(self.resolver.resolve(2.9), QualityLevel.POOR)
