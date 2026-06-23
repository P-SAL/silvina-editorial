from unittest import TestCase

from src.application.classify_article_use_case import ClassifyArticleUseCase
from src.domain.classification.article_classification_response_parser import (
    ArticleClassificationResponseParser,
)
from src.domain.classification.article_classification_text_sampler import (
    ArticleClassificationTextSampler,
)
from src.domain.classification.article_classifier import ArticleClassifier
from src.domain.classification.article_size_classifier import ArticleSizeClassifier
from src.domain.classification.classification_rule_table import ClassificationRuleTable
from src.domain.classification.imryd_signal_detector import ImrydSignalDetector
from src.domain.classification.methodological_vocabulary_detector import (
    MethodologicalVocabularyDetector,
)
from src.domain.classification.reference_signal_detector import ReferenceSignalDetector
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.tests.classification.fake_llm_generator_adapter import FakeLlmGeneratorAdapter


class TestClassifyArticleUseCase(TestCase):
    def setUp(self):
        self.classifier = ArticleClassifier(
            llm_generator=FakeLlmGeneratorAdapter(
                responses=["S4: NO\nS5: NO\nS6: NO", "S4: NO\nS5: NO\nS6: NO"]
            ),
            signal_detector=ImrydSignalDetector(),
            article_size_classifier=ArticleSizeClassifier(),
            text_sampler=ArticleClassificationTextSampler(),
            response_parser=ArticleClassificationResponseParser(),
            signal_prompt_template="{title} {text_sample}",
            temperature=0.1,
            num_predict=300,
            methodological_vocabulary_detector=MethodologicalVocabularyDetector(),
            reference_signal_detector=ReferenceSignalDetector(),
            rule_table=ClassificationRuleTable(),
        )
        self.use_case = ClassifyArticleUseCase(classifier=self.classifier)
        self.document_content = DocumentContentDTO(
            word_count=10,
            char_count=50,
            paragraphs=["A short opinion paragraph with no structural signals at all."],
        )

    def test_execute_returns_domain_service_result_unchanged(self):
        expected = self.classifier.classify(self.document_content)
        result = self.use_case.execute(self.document_content)

        self.assertEqual(result.article_type, expected.article_type)
        self.assertEqual(result.article_size, expected.article_size)
        self.assertEqual(result.confidence, expected.confidence)
        self.assertEqual(result.reasoning, expected.reasoning)
