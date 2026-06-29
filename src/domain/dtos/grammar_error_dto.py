from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class GrammarErrorDTO(BaseDTO):
    """A single grammar error found in the text."""

    number: int
    message: str
    context: str
    offset: int
    length: int
    replacements: list[str]
