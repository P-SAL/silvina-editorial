from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.recommendation_priority import RecommendationPriority


@dataclass(frozen=True)
class RecommendationDTO(BaseDTO):
    """Immutable data transfer object representing an editorial recommendation."""

    priority: RecommendationPriority
    message: str
