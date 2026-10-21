from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.grammar_error_dto import GrammarErrorDTO


@dataclass(frozen=True)
class GrammarCheckResultDTO(BaseDTO):
    """The complete result of a grammar check on a document."""

    score: float
    feedback: str
    errors: list[GrammarErrorDTO]
