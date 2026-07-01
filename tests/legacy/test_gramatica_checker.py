"""
Unit tests for business_logic/gramatica_checker.py
Injects language_tool_python mock into sys.modules to avoid Java dependency.
"""

import sys
import os
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Inject language_tool_python mock before any module-level import can fail
if "language_tool_python" not in sys.modules:
    _mock_ltp_mod = MagicMock()
    sys.modules["language_tool_python"] = _mock_ltp_mod


def _make_mock_match(
    message="Error detected",
    context="context text",
    offset=0,
    error_length=5,
    rule_issue_type="grammar",
    replacements=None,
):
    match = MagicMock()
    match.message = message
    match.context = context
    match.offset = offset
    match.error_length = error_length
    match.rule_issue_type = rule_issue_type
    match.replacements = replacements or ["suggestion"]
    return match


class TestGramaticaCheckerNoErrors(unittest.TestCase):
    """Zero errors → score 10.0."""

    def test_no_errors_returns_perfect_score(self):
        mock_tool = MagicMock()
        mock_tool.check.return_value = []
        sys.modules["language_tool_python"].LanguageTool.return_value = mock_tool

        from business_logic.gramatica_checker import check_gramatica

        score, feedback, errors = check_gramatica(["Texto sin errores gramaticales."])

        self.assertEqual(score, 10.0)
        self.assertIn("Sin errores", feedback)
        self.assertEqual(errors, [])


class TestGramaticaCheckerFewErrors(unittest.TestCase):
    """1-5 errors → score 8.5."""

    def test_few_errors_score_8_5(self):
        mock_tool = MagicMock()
        mock_tool.check.return_value = [_make_mock_match(f"Error {i}") for i in range(3)]
        sys.modules["language_tool_python"].LanguageTool.return_value = mock_tool

        from business_logic.gramatica_checker import check_gramatica

        score, feedback, errors = check_gramatica(["Texto con errores."])

        self.assertEqual(score, 8.5)
        self.assertEqual(len(errors), 3)


class TestGramaticaCheckerManyErrors(unittest.TestCase):
    """6-15 errors → score 7.0."""

    def test_medium_errors_score_7(self):
        mock_tool = MagicMock()
        mock_tool.check.return_value = [_make_mock_match(f"Error {i}") for i in range(10)]
        sys.modules["language_tool_python"].LanguageTool.return_value = mock_tool

        from business_logic.gramatica_checker import check_gramatica

        score, feedback, errors = check_gramatica(["Texto."])

        self.assertEqual(score, 7.0)


class TestGramaticaCheckerExcessErrors(unittest.TestCase):
    """More than 15 errors → score 5.0."""

    def test_excess_errors_score_5(self):
        mock_tool = MagicMock()
        mock_tool.check.return_value = [_make_mock_match(f"Error {i}") for i in range(20)]
        sys.modules["language_tool_python"].LanguageTool.return_value = mock_tool

        from business_logic.gramatica_checker import check_gramatica

        score, feedback, errors = check_gramatica(["Texto."])

        self.assertEqual(score, 5.0)


class TestGramaticaCheckerMisspellingFilter(unittest.TestCase):
    """Misspelling-type matches should be filtered out."""

    def test_misspellings_not_counted(self):
        mock_tool = MagicMock()
        mock_tool.check.return_value = [
            _make_mock_match(rule_issue_type="misspelling"),
            _make_mock_match(rule_issue_type="grammar"),
        ]
        sys.modules["language_tool_python"].LanguageTool.return_value = mock_tool

        from business_logic.gramatica_checker import check_gramatica

        score, feedback, errors = check_gramatica(["Texto."])

        # Only 1 grammar error (misspelling filtered)
        self.assertEqual(score, 8.5)
        self.assertEqual(len(errors), 1)


class TestGramaticaCheckerException(unittest.TestCase):
    """If LanguageTool raises, returns fallback tuple."""

    def test_exception_returns_fallback(self):
        sys.modules["language_tool_python"].LanguageTool.side_effect = Exception("Java not found")

        from business_logic.gramatica_checker import check_gramatica

        score, feedback, errors = check_gramatica(["Texto."])

        self.assertEqual(score, 7.0)
        self.assertIn("no disponible", feedback.lower())
        self.assertEqual(errors, [])

    def tearDown(self):
        # Reset side_effect after exception test
        sys.modules["language_tool_python"].LanguageTool.side_effect = None


class TestGramaticaCheckerReturnStructure(unittest.TestCase):
    """Verify error detail structure."""

    def test_error_detail_has_required_keys(self):
        mock_tool = MagicMock()
        mock_match = _make_mock_match(
            message="Test error",
            context="some context",
            offset=5,
            error_length=3,
            replacements=["fix1", "fix2"],
        )
        mock_tool.check.return_value = [mock_match]
        sys.modules["language_tool_python"].LanguageTool.return_value = mock_tool

        from business_logic.gramatica_checker import check_gramatica

        _, _, errors = check_gramatica(["Texto con error."])

        self.assertEqual(len(errors), 1)
        error = errors[0]
        for key in ["number", "message", "context", "offset", "length", "replacements"]:
            self.assertIn(key, error)


if __name__ == "__main__":
    unittest.main()
