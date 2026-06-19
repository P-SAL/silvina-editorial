from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class DimensionScoreDTO(BaseDTO):
    """A single dimension's parsed score and feedback text."""

    score: float
    feedback: str
