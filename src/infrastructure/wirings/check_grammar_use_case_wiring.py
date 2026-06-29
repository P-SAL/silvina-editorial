from src.application.check_grammar_use_case import CheckGrammarUseCase
from src.domain.grammar.grammar_check_port import GrammarCheckPort
from src.infrastructure.adapters.grammar.language_tool_adapter import LanguageToolAdapter


class CheckGrammarUseCaseWiring:
    """Factory for building a ready-to-use CheckGrammarUseCase."""

    def create_use_case(self) -> CheckGrammarUseCase:
        """Return a fully assembled CheckGrammarUseCase."""
        return CheckGrammarUseCase(grammar_port=self._get_grammar_check_port())

    def _get_grammar_check_port(self) -> GrammarCheckPort:
        """Return the LanguageToolAdapter as the grammar check port."""
        return LanguageToolAdapter()
