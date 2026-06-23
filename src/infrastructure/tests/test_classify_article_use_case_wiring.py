from inspect import signature
from unittest import TestCase

from src.application.classify_article_use_case import ClassifyArticleUseCase
from src.domain.classification.article_classifier import ArticleClassifier
from src.infrastructure.wirings.classify_article_use_case_wiring import (
    ClassifyArticleUseCaseWiring,
)


class TestClassifyArticleUseCaseWiring(TestCase):
    def setUp(self):
        self.wiring = ClassifyArticleUseCaseWiring()

    def test_create_use_case_returns_correct_type(self):
        use_case = self.wiring.create_use_case()
        self.assertIsInstance(use_case, ClassifyArticleUseCase)

    def test_domain_service_constructor_has_no_temperature_or_num_predict_defaults(self):
        parameters = signature(ArticleClassifier.__init__).parameters
        self.assertEqual(parameters["temperature"].default, parameters["temperature"].empty)
        self.assertEqual(parameters["num_predict"].default, parameters["num_predict"].empty)
