"""
Unit tests for the CLI argument parser and main() entry point exit codes in main.py.
"""

import io
import os
import sys
import unittest
from contextlib import redirect_stdout
from unittest.mock import patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from src.domain.enums.article_size import ArticleSize
from src.domain.enums.article_type import ArticleType
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.exceptions.language_model_errors import LanguageModelUnavailable


class TestBuildArgumentParser(unittest.TestCase):
    def setUp(self):
        from main import _build_argument_parser

        self.parser = _build_argument_parser()

    def test_document_path_defaults_to_none(self):
        arguments = self.parser.parse_args([])
        self.assertIsNone(arguments.document_path)

    def test_output_dir_defaults_to_none(self):
        arguments = self.parser.parse_args([])
        self.assertIsNone(arguments.output_dir)

    def test_word_report_path_defaults_to_none(self):
        arguments = self.parser.parse_args([])
        self.assertIsNone(arguments.word_report_path)

    def test_json_report_path_defaults_to_none(self):
        arguments = self.parser.parse_args([])
        self.assertIsNone(arguments.json_report_path)

    def test_parses_all_explicit_arguments(self):
        arguments = self.parser.parse_args(
            [
                "documento.docx",
                "--output-dir",
                "/tmp/salida",
                "--word-report-path",
                "/tmp/salida/reporte.docx",
                "--json-report-path",
                "/tmp/salida/reporte.json",
            ]
        )
        self.assertEqual(arguments.document_path, "documento.docx")
        self.assertEqual(arguments.output_dir, "/tmp/salida")
        self.assertEqual(arguments.word_report_path, "/tmp/salida/reporte.docx")
        self.assertEqual(arguments.json_report_path, "/tmp/salida/reporte.json")


class TestMainExitCodes(unittest.TestCase):
    def test_exits_2_when_file_does_not_exist(self):
        from main import main

        with patch.object(sys, "argv", ["main.py", "nonexistent_document_xyz.docx"]):
            with self.assertRaises(SystemExit) as context:
                main()
        self.assertEqual(context.exception.code, 2)

    def test_exits_2_when_extension_is_not_docx(self):
        from main import main

        with patch.object(sys, "argv", ["main.py", __file__]):
            with self.assertRaises(SystemExit) as context:
                main()
        self.assertEqual(context.exception.code, 2)

    def test_exits_1_when_analyze_document_raises_base_src_error(self):
        from main import main

        fixture_document_path = os.path.join(
            os.path.dirname(__file__),
            "fixtures",
            "capacidades_razonamiento_emergente_LLMs.docx",
        )
        with patch.object(sys, "argv", ["main.py", fixture_document_path]):
            with patch("main.SilvinaEditorialAssistant") as mock_assistant_class:
                mock_assistant_class.return_value.analyze_document.side_effect = (
                    LanguageModelUnavailable()
                )
                with self.assertRaises(SystemExit) as context:
                    main()
        self.assertEqual(context.exception.code, 1)


def _build_legacy_results(total_citations: int) -> dict:
    return {
        "filename": "documento.docx",
        "document_info": {
            "title": "Un Titulo",
            "authors": "Autor Uno",
            "word_count": 1200,
            "char_count": 8000,
            "estimated_pages": 4,
        },
        "classification": {
            "category": ArticleType.CIENTIFICO,
            "article_size": ArticleSize.CORTO,
            "confidence": 0.85,
            "reasoning": "Contiene metodología IMRyD",
        },
        "quality_analysis": {
            "overall_score": 8.2,
            "quality_level": "buena",
            "gramatica": {"score": 9.0, "feedback": "Sin errores relevantes", "errors": []},
            "dimensions": {"claridad": {"score": 8.0, "feedback": "Bien"}},
        },
        "structure_validation": {
            "is_valid": True,
            "missing_sections": [],
            "details": {},
        },
        "citations_analysis": {
            "total_citations": total_citations,
            "total_references": total_citations,
            "matched_count": total_citations,
            "unmatched_count": 0,
            "by_type": {},
            "unmatched_citations": [],
            "apa_violations": 0,
            "apa_compliant": True,
        },
        "apa_validation": {"violations": [], "report": ""},
        "recommendations": [
            {"priority": RecommendationPriority.HIGH.value, "message": "Corregir gramática"},
        ],
    }


class TestMainConsoleSummaryCitationsCount(unittest.TestCase):
    def test_prints_actual_citation_count_from_citations_analysis(self):
        from main import main

        fixture_document_path = os.path.join(
            os.path.dirname(__file__),
            "fixtures",
            "capacidades_razonamiento_emergente_LLMs.docx",
        )
        with patch.object(sys, "argv", ["main.py", fixture_document_path]):
            with patch("main.SilvinaEditorialAssistant") as mock_assistant_class:
                mock_assistant_class.return_value.analyze_document.return_value = (
                    _build_legacy_results(total_citations=7)
                )
                captured_output = io.StringIO()
                with redirect_stdout(captured_output):
                    main()

        self.assertIn("CITAS: 7 detectadas", captured_output.getvalue())
        self.assertNotIn("CITAS: 0 detectadas", captured_output.getvalue())


if __name__ == "__main__":
    unittest.main()
