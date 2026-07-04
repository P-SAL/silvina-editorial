from unittest import TestCase
from unittest.mock import MagicMock

from src.domain.dtos.publication_verdict_dto import PublicationVerdictDTO
from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.dtos.recommendation_settings_dto import RecommendationSettingsDTO
from src.domain.enums.publication_verdict import PublicationVerdict
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.recommendation.recommendation_builder import RecommendationBuilder


def _default_settings(**overrides) -> RecommendationSettingsDTO:
    defaults = {
        "publish_threshold": 7.0,
        "quality_threshold": 7.0,
        "grammar_threshold": 7.0,
        "dimension_threshold": 6.0,
        "citation_match_threshold": 90.0,
        "critical_citation_match_threshold": 50.0,
        "citation_count_threshold": 10,
        "classification_confidence_threshold": 0.7,
        "critical_quality_threshold": 5.0,
        "critical_grammar_threshold": 5.0,
    }
    defaults.update(overrides)
    return RecommendationSettingsDTO(**defaults)


def _make_classification(confidence=0.9):
    m = MagicMock()
    m.confidence = confidence
    return m


def _make_quality(overall_score=8.0, dimension_scores=None):
    m = MagicMock()
    m.overall_score = overall_score
    m.dimension_scores = dimension_scores if dimension_scores is not None else {}
    return m


def _make_grammar(score=8.0):
    m = MagicMock()
    m.score = score
    return m


def _make_structure(is_valid=True, missing_sections=None):
    m = MagicMock()
    m.is_valid = is_valid
    m.missing_sections = missing_sections if missing_sections is not None else []
    return m


def _make_citations(total=12, matched=12, unmatched=0, unmatched_citations=None):
    m = MagicMock()
    m.total_citations = total
    m.matched_count = matched
    m.unmatched_count = unmatched
    m.unmatched_citations = unmatched_citations if unmatched_citations is not None else []
    return m


def _make_apa(violations=None):
    m = MagicMock()
    m.violations = violations if violations is not None else []
    return m


class TestRecommendationBuilder(TestCase):
    def setUp(self):
        self.settings = _default_settings()
        self.builder = RecommendationBuilder(self.settings)

    def _build(self, **overrides):
        defaults = {
            "classification": _make_classification(),
            "quality": _make_quality(),
            "structure": _make_structure(),
            "citations": _make_citations(),
            "apa_validation": _make_apa(),
            "grammar": _make_grammar(),
        }
        defaults.update(overrides)
        return self.builder.build(**defaults)

    # --- Return types ---

    def test_build_returns_tuple_of_recommendations_and_verdict(self):
        result = self._build()
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 2)

    def test_recommendations_are_list_of_recommendation_dto(self):
        recs, _ = self._build()
        self.assertIsInstance(recs, list)
        for rec in recs:
            self.assertIsInstance(rec, RecommendationDTO)

    def test_verdict_is_publication_verdict_dto(self):
        _, verdict = self._build()
        self.assertIsInstance(verdict, PublicationVerdictDTO)

    # --- APPROVED verdict ---

    def test_approved_verdict_when_all_thresholds_satisfied(self):
        _, verdict = self._build()
        self.assertEqual(verdict.verdict, PublicationVerdict.APPROVED)
        self.assertIn("APTO", verdict.message)

    def test_no_specific_recommendations_when_all_thresholds_satisfied(self):
        recs, _ = self._build()
        self.assertEqual(recs, [])

    # --- Quality ---

    def test_high_priority_when_quality_below_threshold(self):
        recs, _ = self._build(quality=_make_quality(overall_score=6.5))
        self.assertTrue(
            any(
                r.priority == RecommendationPriority.HIGH and "calidad" in r.message.lower()
                for r in recs
            )
        )

    def test_warning_verdict_when_quality_below_publish_threshold(self):
        _, verdict = self._build(quality=_make_quality(overall_score=6.5))
        self.assertEqual(verdict.verdict, PublicationVerdict.WARNING)

    def test_critical_verdict_when_quality_below_critical_threshold(self):
        _, verdict = self._build(quality=_make_quality(overall_score=4.5))
        self.assertEqual(verdict.verdict, PublicationVerdict.CRITICAL)

    def test_critical_verdict_uses_custom_critical_quality_threshold(self):
        self.settings = _default_settings(critical_quality_threshold=8.0)
        self.builder = RecommendationBuilder(self.settings)
        _, verdict = self._build(quality=_make_quality(overall_score=7.5))
        self.assertEqual(verdict.verdict, PublicationVerdict.CRITICAL)

    # --- Grammar ---

    def test_high_priority_when_grammar_below_threshold(self):
        recs, _ = self._build(grammar=_make_grammar(score=6.0))
        self.assertTrue(any("ramática" in r.message for r in recs))

    def test_critical_verdict_when_grammar_below_critical_threshold(self):
        _, verdict = self._build(grammar=_make_grammar(score=4.5))
        self.assertEqual(verdict.verdict, PublicationVerdict.CRITICAL)

    def test_critical_verdict_uses_custom_critical_grammar_threshold(self):
        self.settings = _default_settings(critical_grammar_threshold=8.0)
        self.builder = RecommendationBuilder(self.settings)
        _, verdict = self._build(grammar=_make_grammar(score=7.5))
        self.assertEqual(verdict.verdict, PublicationVerdict.CRITICAL)

    # --- Dimension scores ---

    def test_medium_priority_for_low_dimension_score(self):
        dim_scores = {"coherencia": {"score": 5.0, "feedback": "Mejorar coherencia"}}
        recs, _ = self._build(quality=_make_quality(dimension_scores=dim_scores))
        self.assertTrue(
            any(
                r.priority == RecommendationPriority.MEDIUM and "coherencia" in r.message
                for r in recs
            )
        )

    def test_no_dimension_recommendation_when_all_scores_above_threshold(self):
        dim_scores = {"coherencia": {"score": 7.0, "feedback": "Bien"}}
        recs, _ = self._build(quality=_make_quality(dimension_scores=dim_scores))
        self.assertFalse(any("coherencia" in r.message for r in recs))

    # --- Structure ---

    def test_high_priority_for_each_missing_section(self):
        recs, _ = self._build(
            structure=_make_structure(is_valid=False, missing_sections=["Resumen", "Conclusiones"])
        )
        high_missing = [
            r for r in recs if r.priority == RecommendationPriority.HIGH and "Falta" in r.message
        ]
        self.assertEqual(len(high_missing), 2)

    def test_critical_verdict_when_structure_invalid(self):
        _, verdict = self._build(
            structure=_make_structure(is_valid=False, missing_sections=["Resumen"])
        )
        self.assertEqual(verdict.verdict, PublicationVerdict.CRITICAL)

    # --- Citations match rate ---

    def test_high_priority_when_match_rate_below_threshold(self):
        recs, _ = self._build(
            citations=_make_citations(
                total=10, matched=5, unmatched=5, unmatched_citations=["Ref1"]
            )
        )
        self.assertTrue(
            any(
                r.priority == RecommendationPriority.HIGH and "coincidencia" in r.message.lower()
                for r in recs
            )
        )

    def test_medium_priority_when_unmatched_but_above_threshold(self):
        recs, _ = self._build(
            citations=_make_citations(
                total=11, matched=10, unmatched=1, unmatched_citations=["Ref1"]
            )
        )
        self.assertTrue(
            any(
                r.priority == RecommendationPriority.MEDIUM and "no tienen referencia" in r.message
                for r in recs
            )
        )

    def test_medium_priority_when_total_citations_below_threshold(self):
        recs, _ = self._build(citations=_make_citations(total=5, matched=5, unmatched=0))
        self.assertTrue(
            any(
                r.priority == RecommendationPriority.MEDIUM and "bajo de citas" in r.message.lower()
                for r in recs
            )
        )

    def test_critical_verdict_when_match_rate_below_critical_threshold(self):
        _, verdict = self._build(
            citations=_make_citations(total=10, matched=4, unmatched=6, unmatched_citations=["R1"])
        )
        self.assertEqual(verdict.verdict, PublicationVerdict.CRITICAL)

    # --- Zero citations ---

    def test_critical_verdict_when_zero_citations(self):
        _, verdict = self._build(citations=_make_citations(total=0, matched=0, unmatched=0))
        self.assertEqual(verdict.verdict, PublicationVerdict.CRITICAL)
        self.assertIn("No se detectaron citas", verdict.message)

    # --- Classification confidence ---

    def test_low_priority_when_confidence_below_threshold(self):
        recs, _ = self._build(classification=_make_classification(confidence=0.5))
        self.assertTrue(
            any(
                r.priority == RecommendationPriority.LOW and "confianza baja" in r.message
                for r in recs
            )
        )

    def test_no_low_recommendation_when_confidence_is_none(self):
        recs, _ = self._build(classification=_make_classification(confidence=None))
        self.assertFalse(any(r.priority == RecommendationPriority.LOW for r in recs))

    # --- APA violations ---

    def test_warning_verdict_when_apa_violations_present(self):
        _, verdict = self._build(apa_validation=_make_apa(violations=[MagicMock()]))
        self.assertEqual(verdict.verdict, PublicationVerdict.WARNING)
