from unittest import TestCase

from src.domain.classification.methodological_vocabulary_detector import (
    MethodologicalVocabularyDetector,
)
from src.domain.dtos.document_content_dto import DocumentContentDTO


class TestMethodologicalVocabularyDetector(TestCase):
    def setUp(self) -> None:
        self._detector = MethodologicalVocabularyDetector()

    def test_four_general_terms_with_one_hard_term_satisfies_methodological_vocabulary_signal(
        self,
    ) -> None:
        document_content = self._build_document_content(
            paragraphs=[
                "Se aplico una metodologia cuantitativa con hipotesis claras.",
                "El analisis estadistico confirma la correlacion observada.",
            ]
        )

        self.assertTrue(self._detector.detect(document_content))

    def test_four_general_terms_with_zero_hard_terms_does_not_satisfy_signal(self) -> None:
        document_content = self._build_document_content(
            paragraphs=[
                "Se aplico una metodologia cuantitativa con hipotesis claras y variables.",
            ]
        )

        self.assertFalse(self._detector.detect(document_content))

    def test_accent_insensitive_matching_treats_accented_and_unaccented_terms_identically(
        self,
    ) -> None:
        accented_document = self._build_document_content(
            paragraphs=[
                "La metodología, hipótesis, variables y análisis estadístico fueron claros.",
            ]
        )
        unaccented_document = self._build_document_content(
            paragraphs=[
                "La metodologia, hipotesis, variables y analisis estadistico fueron claros.",
            ]
        )

        self.assertTrue(self._detector.detect(accented_document))
        self.assertTrue(self._detector.detect(unaccented_document))

    def _build_document_content(self, paragraphs: list[str]) -> DocumentContentDTO:
        return DocumentContentDTO(word_count=100, char_count=1000, paragraphs=paragraphs)
