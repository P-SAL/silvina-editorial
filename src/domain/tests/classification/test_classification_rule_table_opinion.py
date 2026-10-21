from unittest import TestCase

from src.domain.classification.classification_rule_table import ClassificationRuleTable
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO
from src.domain.enums.article_type import ArticleType


class TestClassificationRuleTableOpinion(TestCase):
    def setUp(self) -> None:
        self._rule_table = ClassificationRuleTable()

    def test_case_19_no_signals_detected_yields_opinion(self) -> None:
        signals = ClassificationSignalsDTO(
            has_sufficient_reference_count=False,
            has_recent_references=False,
            has_methodological_vocabulary=False,
            has_research_intent=False,
            has_evidence_based_contribution=False,
            has_theoretical_justification=False,
        )

        matched_rule = self._rule_table.evaluate(signals)

        self.assertEqual(matched_rule.article_type, ArticleType.OPINION)
        self.assertIsNone(matched_rule.confidence)
        self.assertIn(
            "No se detectaron señales de investigación científica ni de divulgación",
            matched_rule.reasoning_template,
        )

    def test_rule_table_has_eighteen_rows_including_opinion_fallback(self) -> None:
        self.assertEqual(len(ClassificationRuleTable._ROWS), 18)
