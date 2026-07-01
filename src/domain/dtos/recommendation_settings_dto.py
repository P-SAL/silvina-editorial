from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class RecommendationSettingsDTO(BaseDTO):
    """Configuration settings for generating quality and formatting recommendations."""

    publish_threshold: float
    quality_threshold: float
    grammar_threshold: float
    dimension_threshold: float
    citation_match_threshold: float
    critical_citation_match_threshold: float
    citation_count_threshold: int
    classification_confidence_threshold: float
