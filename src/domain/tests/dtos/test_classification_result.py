from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.classification_result_dto import ClassificationResult
from src.domain.enums.article_size import ArticleSize
from src.domain.enums.article_type import ArticleType


class TestClassificationResult(TestCase):
    def test_classification_result_is_subclass_of_base_dto(self):
        from src.domain.dtos.base_dto import BaseDTO

        self.assertTrue(issubclass(ClassificationResult, BaseDTO))

    def test_classification_result_instantiation_with_correct_fields(self):
        result = ClassificationResult(
            article_type=ArticleType.CIENTIFICO,
            article_size=ArticleSize.LARGO,
            confidence=0.9,
            reasoning="Well structured article",
        )
        self.assertEqual(result.article_type, ArticleType.CIENTIFICO)
        self.assertEqual(result.article_size, ArticleSize.LARGO)
        self.assertAlmostEqual(result.confidence, 0.9)
        self.assertEqual(result.reasoning, "Well structured article")

    def test_classification_result_is_immutable(self):
        result = ClassificationResult(
            article_type=ArticleType.CIENTIFICO,
            article_size=ArticleSize.LARGO,
            confidence=0.8,
            reasoning="Test",
        )
        with self.assertRaises(FrozenInstanceError):
            result.confidence = 0.5

    def test_create_factory_builds_valid_instance(self):
        result = ClassificationResult.create(
            article_type=ArticleType.OPINION,
            article_size=ArticleSize.CORTO,
            confidence=0.75,
            reasoning="Opinion piece",
        )
        self.assertEqual(result.article_type, ArticleType.OPINION)
        self.assertEqual(result.article_size, ArticleSize.CORTO)
        self.assertIsNotNone(result.timestamp)

    def test_create_factory_with_none_confidence(self):
        result = ClassificationResult.create(
            article_type=ArticleType.UNKNOWN,
            article_size=ArticleSize.FUERA_RANGO,
            confidence=None,
            reasoning="Unknown",
        )
        self.assertIsNone(result.confidence)

    def test_create_factory_result_is_frozen(self):
        result = ClassificationResult.create(
            article_type=ArticleType.CIENTIFICO,
            article_size=ArticleSize.LARGO,
            confidence=0.9,
            reasoning="Test",
        )
        with self.assertRaises(FrozenInstanceError):
            result.reasoning = "Modified"

    def test_str_contains_enum_values_and_confidence_percentage(self):
        result = ClassificationResult(
            article_type=ArticleType.CIENTIFICO,
            article_size=ArticleSize.LARGO,
            confidence=0.9,
            reasoning="Test",
        )
        string_repr = str(result)
        self.assertIn("científico", string_repr)
        self.assertIn("largo", string_repr)
        self.assertIn("%", string_repr)
