import language_tool_python
from typing import Any

from src.domain.dtos.grammar_error_dto import GrammarErrorDTO
from src.domain.exceptions.grammar_errors import GrammarCheckUnavailable
from src.domain.grammar.grammar_check_port import GrammarCheckPort

_MAX_PARAGRAPHS = 20
_MAX_CHARS = 5000
_MAX_ERRORS = 10


class LanguageToolAdapter(GrammarCheckPort):
    """Adapter that uses LanguageTool for grammar checking with lazy Java initialization."""

    def __init__(self, max_replacements: int, language: str = "es") -> None:
        self._language = language
        self._max_replacements = max_replacements
        self._tool = None

    def check(self, paragraphs: list[str]) -> list[GrammarErrorDTO]:
        """Return grammar errors found in the given paragraphs."""
        self._initialize_tool_if_needed()
        text = self._build_sample_text(paragraphs=paragraphs)
        try:
            raw_matches = self._tool.check(text)
        except Exception as exc:
            raise GrammarCheckUnavailable() from exc
        grammar_matches = [match for match in raw_matches if match.rule_issue_type != "misspelling"]
        return [
            self._map_to_dto(number=index + 1, match=match)
            for index, match in enumerate(grammar_matches[:_MAX_ERRORS])
        ]

    def _build_sample_text(self, paragraphs: list[str]) -> str:
        """Return a text sample truncated to the configured limits."""
        return "\n".join(paragraphs[:_MAX_PARAGRAPHS])[:_MAX_CHARS]

    def _initialize_tool_if_needed(self) -> None:
        """Initialize the LanguageTool instance on first call."""
        if self._tool is not None:
            return
        try:
            self._tool = language_tool_python.LanguageTool(self._language)
        except Exception as exc:
            raise GrammarCheckUnavailable() from exc

    def _map_to_dto(self, number: int, match: Any) -> GrammarErrorDTO:
        """Convert a LanguageTool match to a GrammarErrorDTO."""
        return GrammarErrorDTO(
            number=number,
            message=match.message,
            context=match.context,
            offset=match.offset,
            length=match.error_length,
            replacements=match.replacements[: self._max_replacements],
        )
