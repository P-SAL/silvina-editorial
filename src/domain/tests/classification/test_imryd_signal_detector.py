from unittest import TestCase

from src.domain.classification.imryd_signal_detector import ImrydSignalDetector
from src.domain.dtos.document_content_dto import DocumentContentDTO


class TestImrydSignalDetector(TestCase):
    def _build_document_content(self, paragraphs: list[str]) -> DocumentContentDTO:
        full_text = " ".join(paragraphs)
        return DocumentContentDTO(
            word_count=len(full_text.split()),
            char_count=len(full_text),
            paragraphs=paragraphs,
        )

    def test_long_body_paragraph_is_never_treated_as_section_header(self):
        paragraph = (
            "Esta introducción extensa describe resultados preliminares "
            "obtenidos durante la fase exploratoria del estudio realizado."
        )
        document_content = self._build_document_content([paragraph])
        detector = ImrydSignalDetector()

        signals = detector.detect(document_content)

        self.assertFalse(signals["has_introduction"])
        self.assertFalse(signals["has_results"])
        self.assertFalse(signals["has_methods"])
        self.assertFalse(signals["has_discussion"])
        self.assertFalse(signals["has_conclusion"])

    def test_all_four_core_sections_present_yields_imryd_complete_true(self):
        paragraphs = ["Introduction", "Methods", "Results", "Discussion"]
        document_content = self._build_document_content(paragraphs)
        detector = ImrydSignalDetector()

        signals = detector.detect(document_content)

        self.assertTrue(signals["has_introduction"])
        self.assertTrue(signals["has_methods"])
        self.assertTrue(signals["has_results"])
        self.assertTrue(signals["has_discussion"])
        self.assertFalse(signals["has_conclusion"])
        self.assertTrue(signals["imryd_complete"])

    def test_conclusion_alone_does_not_satisfy_imryd_complete(self):
        document_content = self._build_document_content(["Conclusion"])
        detector = ImrydSignalDetector()

        signals = detector.detect(document_content)

        self.assertTrue(signals["has_conclusion"])
        self.assertFalse(signals["imryd_complete"])

    def test_bilingual_keyword_matching_covers_spanish_and_english_variants(self):
        spanish_paragraphs = ["Metodología", "Resultados", "Discusión"]
        english_paragraphs = ["Methodology", "Results", "Discussion"]
        detector = ImrydSignalDetector()

        spanish_signals = detector.detect(self._build_document_content(spanish_paragraphs))
        english_signals = detector.detect(self._build_document_content(english_paragraphs))

        self.assertTrue(spanish_signals["has_methods"])
        self.assertTrue(spanish_signals["has_results"])
        self.assertTrue(spanish_signals["has_discussion"])
        self.assertEqual(
            (
                spanish_signals["has_methods"],
                spanish_signals["has_results"],
                spanish_signals["has_discussion"],
            ),
            (
                english_signals["has_methods"],
                english_signals["has_results"],
                english_signals["has_discussion"],
            ),
        )
