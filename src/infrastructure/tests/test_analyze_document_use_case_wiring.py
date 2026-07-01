import os
from unittest import TestCase
from unittest.mock import patch

from src.application.analyze_document_use_case import AnalyzeDocumentUseCase
from src.domain.dtos.recommendation_settings_dto import RecommendationSettingsDTO
from src.infrastructure.wirings.analyze_document_use_case_wiring import AnalyzeDocumentUseCaseWiring


class TestAnalyzeDocumentUseCaseWiring(TestCase):
    def test_create_use_case_returns_analyze_document_use_case_instance(self):
        result = AnalyzeDocumentUseCaseWiring().create_use_case()
        self.assertIsInstance(result, AnalyzeDocumentUseCase)

    def test_create_use_case_wires_all_sub_use_cases(self):
        result = AnalyzeDocumentUseCaseWiring().create_use_case()
        self.assertIsNotNone(result._read_document_use_case)
        self.assertIsNotNone(result._extract_content_use_case)
        self.assertIsNotNone(result._extract_citations_use_case)
        self.assertIsNotNone(result._validate_apa_use_case)
        self.assertIsNotNone(result._check_grammar_use_case)
        self.assertIsNotNone(result._classify_article_use_case)
        self.assertIsNotNone(result._analyze_quality_use_case)
        self.assertIsNotNone(result._validate_structure_use_case)
        self.assertIsNotNone(result._match_citations_use_case)
        self.assertIsNotNone(result._verify_eumic_use_case)
        self.assertIsNotNone(result._recommendation_builder)

    def test_env_var_overrides_quality_threshold(self):
        with patch.dict(os.environ, {"RECOMMENDATION_QUALITY_THRESHOLD": "6.5"}):
            result = AnalyzeDocumentUseCaseWiring().create_use_case()
        settings: RecommendationSettingsDTO = result._recommendation_builder._settings
        self.assertAlmostEqual(settings.quality_threshold, 6.5)

    def test_env_var_overrides_grammar_threshold(self):
        with patch.dict(os.environ, {"RECOMMENDATION_GRAMMAR_THRESHOLD": "5.0"}):
            result = AnalyzeDocumentUseCaseWiring().create_use_case()
        settings: RecommendationSettingsDTO = result._recommendation_builder._settings
        self.assertAlmostEqual(settings.grammar_threshold, 5.0)

    def test_default_thresholds_when_env_vars_absent(self):
        env_without = {k: v for k, v in os.environ.items() if not k.startswith("RECOMMENDATION_")}
        with patch.dict(os.environ, env_without, clear=True):
            result = AnalyzeDocumentUseCaseWiring().create_use_case()
        settings: RecommendationSettingsDTO = result._recommendation_builder._settings
        self.assertAlmostEqual(settings.publish_threshold, 7.0)
        self.assertAlmostEqual(settings.quality_threshold, 7.0)
        self.assertAlmostEqual(settings.grammar_threshold, 7.0)
        self.assertAlmostEqual(settings.dimension_threshold, 6.0)
        self.assertAlmostEqual(settings.citation_match_threshold, 90.0)
        self.assertAlmostEqual(settings.critical_citation_match_threshold, 50.0)
        self.assertEqual(settings.citation_count_threshold, 10)
        self.assertAlmostEqual(settings.classification_confidence_threshold, 0.7)
