from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class QualityLevelThresholdsDTO(BaseDTO):
    """Score tier boundaries used to resolve a QualityLevel."""

    excellent_threshold: float
    good_threshold: float
    acceptable_threshold: float
    needs_improvement_threshold: float
