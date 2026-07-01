"""
Unit tests for main.py's handling of Ollama/LLM backend failures.
"""

import io
import os
import sys
import unittest
from contextlib import redirect_stdout
from unittest.mock import MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from src.domain.exceptions.language_model_errors import LanguageModelUnavailable


class TestAnalyzeDocumentOllamaFailure(unittest.TestCase):
    def setUp(self):
        from main import SilvinaEditorialAssistant

        self.assistant = SilvinaEditorialAssistant.__new__(SilvinaEditorialAssistant)
        self.assistant._analyze_document_use_case = MagicMock()
        self.assistant._export_report_use_case = MagicMock()
        self.assistant._last_report_input = None

    def test_reraises_language_model_unavailable_without_wrapping(self):
        self.assistant._analyze_document_use_case.execute.side_effect = LanguageModelUnavailable()
        with self.assertRaises(LanguageModelUnavailable):
            self.assistant.analyze_document("some/path.docx")

    def test_does_not_store_report_when_llm_unavailable(self):
        self.assistant._analyze_document_use_case.execute.side_effect = LanguageModelUnavailable()
        with self.assertRaises(LanguageModelUnavailable):
            self.assistant.analyze_document("some/path.docx")
        self.assertIsNone(self.assistant._last_report_input)

    def test_prints_clean_language_model_message_not_generic_error(self):
        self.assistant._analyze_document_use_case.execute.side_effect = LanguageModelUnavailable()
        captured_output = io.StringIO()
        with redirect_stdout(captured_output):
            with self.assertRaises(LanguageModelUnavailable):
                self.assistant.analyze_document("some/path.docx")
        printed_text = captured_output.getvalue()
        self.assertIn(LanguageModelUnavailable.MESSAGE, printed_text)
        self.assertNotIn("Error durante el análisis", printed_text)


if __name__ == "__main__":
    unittest.main()
