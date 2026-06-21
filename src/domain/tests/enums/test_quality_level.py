from unittest import TestCase

from src.domain.enums.quality_level import QualityLevel


class TestQualityLevel(TestCase):
    def test_members_and_values(self):
        self.assertEqual(QualityLevel.EXCELLENT.value, "Excelente")
        self.assertEqual(QualityLevel.GOOD.value, "Bueno")
        self.assertEqual(QualityLevel.ACCEPTABLE.value, "Aceptable")
        self.assertEqual(QualityLevel.NEEDS_IMPROVEMENT.value, "Requiere mejoras")
        self.assertEqual(QualityLevel.POOR.value, "Deficiente")

    def test_member_count(self):
        self.assertEqual(len(QualityLevel), 5)
