from src.domain.dtos.grammar_error_dto import GrammarErrorDTO
from src.domain.grammar.grammar_check_port import GrammarCheckPort


class FakeGrammarCheckPort(GrammarCheckPort):
    """Test double for GrammarCheckPort with configurable results or exceptions."""

    def __init__(
        self,
        errors: list[GrammarErrorDTO] | None = None,
        error: Exception | None = None,
    ) -> None:
        self._errors = errors
        self._error = error

    def check(self, paragraphs: list[str]) -> list[GrammarErrorDTO]:
        """Return configured error list, raise configured exception, or return empty list."""
        if self._error is not None:
            raise self._error
        return self._errors or []
