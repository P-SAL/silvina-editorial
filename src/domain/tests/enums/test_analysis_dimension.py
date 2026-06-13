from unittest import TestCase

from src.domain.enums.analysis_dimension import AnalysisDimension


class TestAnalysisDimension(TestCase):
    def test_members_and_values(self):
        self.assertEqual(AnalysisDimension.ACADEMIC_RIGOR.value, "academic_rigor")
        self.assertEqual(AnalysisDimension.METHODOLOGICAL_CLARITY.value, "methodological_clarity")
        self.assertEqual(AnalysisDimension.ARGUMENTATION.value, "argumentation")
        self.assertEqual(AnalysisDimension.LITERATURE_REVIEW.value, "literature_review")
        self.assertEqual(AnalysisDimension.ORIGINALITY.value, "originality")
        self.assertEqual(AnalysisDimension.WRITING_QUALITY.value, "writing_quality")
        self.assertEqual(AnalysisDimension.STRUCTURE.value, "structure")
        self.assertEqual(AnalysisDimension.CITATION_QUALITY.value, "citation_quality")

    def test_member_count(self):
        self.assertEqual(len(AnalysisDimension), 8)
