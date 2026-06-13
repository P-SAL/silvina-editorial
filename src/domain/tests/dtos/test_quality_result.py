from dataclasses import FrozenInstanceError
from typing import Any, get_type_hints
from unittest import TestCase

from src.domain.dtos.quality_result_dto import QualityResult
from src.domain.enums.quality_level import QualityLevel


class TestQualityResult(TestCase):
    def test_quality_result_is_subclass_of_base_dto(self):
        from src.domain.dtos.base_dto import BaseDTO

        self.assertTrue(issubclass(QualityResult, BaseDTO))

    def test_quality_result_is_immutable(self):
        result = QualityResult(overall_score=8.5, quality_level=QualityLevel.GOOD)
        with self.assertRaises(FrozenInstanceError):
            result.overall_score = 5.0

    def test_quality_result_str_returns_score_and_level(self):
        result = QualityResult(overall_score=8.5, quality_level=QualityLevel.GOOD)
        self.assertEqual(str(result), "Quality: 8.5/10 (Bueno)")

    def test_quality_analysis_result_does_not_exist_in_src(self):
        with self.assertRaises(ImportError):
            from src.domain.dtos.quality_analysis_result_dto import QualityAnalysisResult  # noqa: F401

    def test_dimension_scores_annotation_uses_typing_any(self):
        hints = get_type_hints(QualityResult)
        expected = dict[str, dict[str, Any]]
        self.assertEqual(hints["dimension_scores"], expected)
