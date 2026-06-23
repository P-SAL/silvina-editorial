# ruff: noqa: E402
"""
Smoke tests: parity between legacy ArticleClassifier and new ClassifyArticleUseCase.

Legacy path:
    WordReader -> paragraphs -> LegacyClassifier.classify_article() -> ClassificationResult

New path:
    WordReader -> paragraphs -> DocumentContentDTO -> ClassifyArticleUseCase.execute()
    -> ClassificationResultDTO

Both sides' LLM calls are mocked with a canned response so this test never touches a
live Ollama instance — only the S4/S5/S6 signal-extraction network call is faked; real
.docx parsing and the deterministic signals (IMRyD override, S2a/S2b/S3) still run.

Run with: python -m pytest tests/smoke/ -v
"""

import sys
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch

ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(ROOT))

from data_access.word_reader import WordReader
from business_logic.article_classifier import ArticleClassifier as LegacyClassifier
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.infrastructure.wirings.classify_article_use_case_wiring import (
    ClassifyArticleUseCaseWiring,
)

DOCS = ROOT / "docs" / "sample-documents"
_CANNED_RESPONSE = {"response": "S4: SI\nS5: SI\nS6: SI"}
_DOCUMENTS = ["1. test_Científico.docx", "2. test_divulgacion_v2.docx", "3. test_opinion_v2.docx"]


class TestClassifyArticleParity(TestCase):
    @classmethod
    def setUpClass(cls):
        cls.reader = WordReader()
        cls.legacy = LegacyClassifier()
        cls.use_case = ClassifyArticleUseCaseWiring().create_use_case()

    def _run(self, filename: str):
        paragraphs = self.reader.read_word_document(str(DOCS / filename))
        document_content = DocumentContentDTO(
            word_count=sum(len(paragraph.split()) for paragraph in paragraphs),
            char_count=sum(len(paragraph) for paragraph in paragraphs),
            paragraphs=paragraphs,
        )
        with patch(
            "business_logic.article_classifier.ollama.Client.generate",
            return_value=_CANNED_RESPONSE,
        ) as legacy_generate:
            legacy_result = self.legacy.classify_article(document_content)
        with patch(
            "src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.generate",
            return_value=_CANNED_RESPONSE,
        ) as new_generate:
            new_result = self.use_case.execute(document_content)
        return legacy_result, new_result, legacy_generate, new_generate

    def test_cientifico_parity(self):
        legacy, new, legacy_generate, new_generate = self._run(_DOCUMENTS[0])
        self._assert_mocks_intercepted_the_same_number_of_calls(legacy_generate, new_generate)
        self.assertEqual(new.article_type.value, legacy.article_type.value)
        self.assertEqual(new.confidence, legacy.confidence)

    def test_divulgacion_parity(self):
        legacy, new, legacy_generate, new_generate = self._run(_DOCUMENTS[1])
        self._assert_mocks_intercepted_the_same_number_of_calls(legacy_generate, new_generate)
        self.assertEqual(new.article_type.value, legacy.article_type.value)
        self.assertEqual(new.confidence, legacy.confidence)

    def test_opinion_parity(self):
        legacy, new, legacy_generate, new_generate = self._run(_DOCUMENTS[2])
        self._assert_mocks_intercepted_the_same_number_of_calls(legacy_generate, new_generate)
        self.assertEqual(new.article_type.value, legacy.article_type.value)
        self.assertEqual(new.confidence, legacy.confidence)

    def _assert_mocks_intercepted_the_same_number_of_calls(self, legacy_generate, new_generate):
        self.assertEqual(new_generate.call_count, legacy_generate.call_count)
