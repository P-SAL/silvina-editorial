from importlib.util import find_spec
from unittest import TestCase, skipIf

from src.application.check_grammar_use_case import CheckGrammarUseCase

_LANGUAGE_TOOL_AVAILABLE = find_spec("language_tool_python") is not None

if _LANGUAGE_TOOL_AVAILABLE:
    from src.infrastructure.adapters.grammar.language_tool_adapter import LanguageToolAdapter
    from src.infrastructure.wirings.check_grammar_use_case_wiring import CheckGrammarUseCaseWiring


@skipIf(not _LANGUAGE_TOOL_AVAILABLE, "language_tool_python not available")
class TestCheckGrammarUseCaseWiring(TestCase):
    def test_create_use_case_returns_check_grammar_use_case_instance(self):
        result = CheckGrammarUseCaseWiring().create_use_case()
        self.assertIsInstance(result, CheckGrammarUseCase)

    def test_create_use_case_wires_language_tool_adapter_as_grammar_port(self):
        result = CheckGrammarUseCaseWiring().create_use_case()
        self.assertIsInstance(result._grammar_port, LanguageToolAdapter)
