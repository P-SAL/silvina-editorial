from unittest import TestCase

from src.domain.enums.classification_category import ClassificationCategory


class TestClassificationCategory(TestCase):
    def test_members_and_values(self):
        self.assertEqual(ClassificationCategory.RESEARCH_ARTICLE.value, "research_article")
        self.assertEqual(ClassificationCategory.REVIEW_ARTICLE.value, "review_article")
        self.assertEqual(ClassificationCategory.REFLECTION_ARTICLE.value, "reflection_article")
        self.assertEqual(ClassificationCategory.SHORT_ARTICLE.value, "short_article")
        self.assertEqual(ClassificationCategory.CASE_REPORT.value, "case_report")
        self.assertEqual(ClassificationCategory.UNKNOWN.value, "unknown")

    def test_member_count(self):
        self.assertEqual(len(ClassificationCategory), 6)
