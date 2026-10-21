from abc import ABC, abstractmethod

from src.domain.dtos.grammar_error_dto import GrammarErrorDTO


class GrammarCheckPort(ABC):
    """Port for grammar checking services."""

    @abstractmethod
    def check(self, paragraphs: list[str]) -> list[GrammarErrorDTO]:
        """Return grammar errors found in the given paragraphs."""
