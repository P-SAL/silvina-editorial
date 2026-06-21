from unittest import TestCase

from src.application.analyze_quality_use_case import AnalyzeQualityUseCase
from src.application.tests.fake_llm_generator_adapter import FakeLlmGeneratorAdapterForTest
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.quality.quality_analyzer import QualityAnalyzer
from src.domain.quality.quality_response_parser import QualityResponseParser
from src.domain.quality.quality_text_sampler import QualityTextSampler


class TestAnalyzeQualityUseCase(TestCase):
    def setUp(self):
        self.analyzer = QualityAnalyzer(
            llm_generator=FakeLlmGeneratorAdapterForTest(),
            text_sampler=QualityTextSampler(),
            response_parser=QualityResponseParser(),
            clarity_coherence_prompt_template="{text_sample}",
            argumentation_conclusions_prompt_template="{text_sample}",
        )
        self.use_case = AnalyzeQualityUseCase(analyzer=self.analyzer)
        self.document_content = DocumentContentDTO(
            word_count=500,
            char_count=3000,
            paragraphs=["Intro paragraph with enough words to pass the sample threshold."] * 10,
        )

    def test_execute_matches_domain_service_result(self):
        expected = self.analyzer.analyze(self.document_content, article_type=None)
        result = self.use_case.execute(self.document_content, article_type=None)

        self.assertEqual(result.overall_score, expected.overall_score)
        self.assertEqual(result.quality_level, expected.quality_level)
        self.assertEqual(result.dimension_scores, expected.dimension_scores)
