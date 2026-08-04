import shutil
from importlib.util import find_spec
from unittest import TestCase, skipIf
from unittest.mock import MagicMock, patch

from src.domain.exceptions.grammar_errors import GrammarCheckUnavailable

_JAVA_AVAILABLE = shutil.which("java") is not None
_LANGUAGE_TOOL_AVAILABLE = find_spec("language_tool_python") is not None

if _LANGUAGE_TOOL_AVAILABLE:
    from src.infrastructure.adapters.grammar.language_tool_adapter import LanguageToolAdapter


@skipIf(
    not _JAVA_AVAILABLE or not _LANGUAGE_TOOL_AVAILABLE,
    "Java or language_tool_python not available",
)
class TestLanguageToolAdapter(TestCase):
    def test_tool_is_none_after_construction_before_any_check_call(self):
        adapter = LanguageToolAdapter(max_replacements=3)
        self.assertIsNone(adapter._tool)

    @patch("src.infrastructure.adapters.grammar.language_tool_adapter.language_tool_python")
    def test_misspelling_matches_are_filtered_from_results(self, mock_language_tool_python):
        mock_tool = MagicMock()
        mock_language_tool_python.LanguageTool.return_value = mock_tool

        grammar_match = MagicMock()
        grammar_match.rule_issue_type = "grammar"
        grammar_match.message = "Grammar error"
        grammar_match.context = "some context"
        grammar_match.offset = 0
        grammar_match.error_length = 4
        grammar_match.replacements = ["fix"]

        spelling_match = MagicMock()
        spelling_match.rule_issue_type = "misspelling"

        mock_tool.check.return_value = [grammar_match, spelling_match, grammar_match]

        adapter = LanguageToolAdapter(max_replacements=3)
        result = adapter.check(paragraphs=["some text"])

        self.assertEqual(len(result), 2)

    @patch("src.infrastructure.adapters.grammar.language_tool_adapter.language_tool_python")
    def test_output_is_capped_at_ten_errors(self, mock_language_tool_python):
        mock_tool = MagicMock()
        mock_language_tool_python.LanguageTool.return_value = mock_tool

        def make_match(index: int) -> MagicMock:
            match = MagicMock()
            match.rule_issue_type = "grammar"
            match.message = f"error {index}"
            match.context = "ctx"
            match.offset = 0
            match.error_length = 3
            match.replacements = []
            return match

        mock_tool.check.return_value = [make_match(index=index) for index in range(12)]

        adapter = LanguageToolAdapter(max_replacements=3)
        result = adapter.check(paragraphs=["text"])

        self.assertEqual(len(result), 10)

    @patch("src.infrastructure.adapters.grammar.language_tool_adapter.language_tool_python")
    def test_raises_grammar_check_unavailable_on_backend_failure(self, mock_language_tool_python):
        mock_language_tool_python.LanguageTool.side_effect = RuntimeError("Java crash")

        adapter = LanguageToolAdapter(max_replacements=3)

        with self.assertRaises(GrammarCheckUnavailable):
            adapter.check(paragraphs=["text"])

    @patch("src.infrastructure.adapters.grammar.language_tool_adapter.language_tool_python")
    def test_custom_max_replacements_limits_suggestions_per_error(self, mock_language_tool_python):
        mock_tool = MagicMock()
        mock_language_tool_python.LanguageTool.return_value = mock_tool

        grammar_match = MagicMock()
        grammar_match.rule_issue_type = "grammar"
        grammar_match.message = "Grammar error"
        grammar_match.context = "some context"
        grammar_match.offset = 0
        grammar_match.error_length = 4
        grammar_match.replacements = ["a", "b", "c", "d", "e"]

        mock_tool.check.return_value = [grammar_match]

        adapter = LanguageToolAdapter(max_replacements=2)
        result = adapter.check(paragraphs=["some text"])

        self.assertEqual(len(result[0].replacements), 2)

    @patch("src.infrastructure.adapters.grammar.language_tool_adapter.language_tool_python")
    def test_default_max_replacements_caps_suggestions_at_three(self, mock_language_tool_python):
        mock_tool = MagicMock()
        mock_language_tool_python.LanguageTool.return_value = mock_tool

        grammar_match = MagicMock()
        grammar_match.rule_issue_type = "grammar"
        grammar_match.message = "Grammar error"
        grammar_match.context = "some context"
        grammar_match.offset = 0
        grammar_match.error_length = 4
        grammar_match.replacements = ["a", "b", "c", "d", "e"]

        mock_tool.check.return_value = [grammar_match]

        adapter = LanguageToolAdapter(max_replacements=3)
        result = adapter.check(paragraphs=["some text"])

        self.assertEqual(len(result[0].replacements), 3)
