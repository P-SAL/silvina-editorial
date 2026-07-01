from abc import ABC, abstractmethod

from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.recommendation.analysis_context import AnalysisContext


class RecommendationRule(ABC):
    """Abstract base for a single recommendation evaluation rule."""

    @abstractmethod
    def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]:
        """Return zero or more recommendations based on the analysis context."""
