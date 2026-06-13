from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.analysis_result_dto import AnalysisResult
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResult
from src.domain.dtos.classification_result_dto import ClassificationResult
from src.domain.dtos.quality_result_dto import QualityResult
from src.domain.dtos.structure_validation_result_dto import StructureValidationResult
from src.domain.document.document_content import DocumentContent
from src.domain.enums.article_size import ArticleSize
from src.domain.enums.article_type import ArticleType
from src.domain.enums.quality_level import QualityLevel


class TestAnalysisResult(TestCase):
    def _make_analysis_result(self) -> AnalysisResult:
        document_content = DocumentContent(word_count=500, char_count=3000)
        classification = ClassificationResult(
            article_type=ArticleType.CIENTIFICO,
            article_size=ArticleSize.LARGO,
            confidence=0.9,
            reasoning="Scientific article",
        )
        quality = QualityResult(
            overall_score=8.0,
            quality_level=QualityLevel.GOOD,
        )
        structure = StructureValidationResult(is_valid=True)
        citations = CitationAnalysisResult(
            total_citations=10,
            total_references=8,
            matched_count=8,
            unmatched_count=2,
        )
        return AnalysisResult(
            filename="test_article.docx",
            document_content=document_content,
            classification=classification,
            quality=quality,
            structure=structure,
            citations=citations,
        )

    def test_analysis_result_is_subclass_of_base_dto(self):
        from src.domain.dtos.base_dto import BaseDTO

        self.assertTrue(issubclass(AnalysisResult, BaseDTO))

    def test_analysis_result_is_immutable(self):
        result = self._make_analysis_result()
        with self.assertRaises(FrozenInstanceError):
            result.filename = "other.docx"

    def test_to_dict_returns_all_required_top_level_keys(self):
        result = self._make_analysis_result()
        output = result.to_dict()
        for key in ("filename", "timestamp", "classification", "quality", "structure", "citations"):
            self.assertIn(key, output)

    def test_to_dict_classification_matches_legacy_shape(self):
        result = self._make_analysis_result()
        classification_dict = result.to_dict()["classification"]
        self.assertEqual(
            set(classification_dict.keys()),
            {"category", "confidence", "reasoning"},
        )
        self.assertEqual(classification_dict["category"], ArticleType.CIENTIFICO.value)

    def test_to_dict_timestamp_is_iso8601_string(self):
        result = self._make_analysis_result()
        timestamp_value = result.to_dict()["timestamp"]
        self.assertIsInstance(timestamp_value, str)
        self.assertEqual(timestamp_value, result.timestamp.isoformat())
