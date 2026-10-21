from unittest import TestCase

from src.domain.enums.recommendation_priority import RecommendationPriority


class TestRecommendationPriority(TestCase):
    def test_members_and_values(self):
        self.assertEqual(RecommendationPriority.HIGH.value, "alta")
        self.assertEqual(RecommendationPriority.MEDIUM.value, "media")
        self.assertEqual(RecommendationPriority.LOW.value, "baja")

    def test_member_count(self):
        self.assertEqual(len(RecommendationPriority), 3)
