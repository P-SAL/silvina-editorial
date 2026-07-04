from os import getenv

from src.domain.dtos.recommendation_settings_dto import RecommendationSettingsDTO


class RecommendationConfig:
    """Builds RecommendationSettingsDTO from environment variables at call time."""

    @classmethod
    def build_settings(cls) -> RecommendationSettingsDTO:
        return RecommendationSettingsDTO(
            publish_threshold=float(getenv("RECOMMENDATION_PUBLISH_THRESHOLD", "7.0")),
            quality_threshold=float(getenv("RECOMMENDATION_QUALITY_THRESHOLD", "7.0")),
            grammar_threshold=float(getenv("RECOMMENDATION_GRAMMAR_THRESHOLD", "7.0")),
            dimension_threshold=float(getenv("RECOMMENDATION_DIMENSION_THRESHOLD", "6.0")),
            citation_match_threshold=float(
                getenv("RECOMMENDATION_CITATION_MATCH_THRESHOLD", "90.0")
            ),
            critical_citation_match_threshold=float(
                getenv("RECOMMENDATION_CRITICAL_CITATION_MATCH_THRESHOLD", "50.0")
            ),
            citation_count_threshold=int(getenv("RECOMMENDATION_CITATION_COUNT_THRESHOLD", "10")),
            classification_confidence_threshold=float(
                getenv("RECOMMENDATION_CLASSIFICATION_CONFIDENCE_THRESHOLD", "0.7")
            ),
            critical_quality_threshold=float(
                getenv("RECOMMENDATION_CRITICAL_QUALITY_THRESHOLD", "5.0")
            ),
            critical_grammar_threshold=float(
                getenv("RECOMMENDATION_CRITICAL_GRAMMAR_THRESHOLD", "5.0")
            ),
        )
