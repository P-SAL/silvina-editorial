from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.recommendation.analysis_context import AnalysisContext
from src.domain.recommendation.recommendation_rule import RecommendationRule


class QualityRule(RecommendationRule):
    def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]:
        if context.quality.overall_score >= context.settings.quality_threshold:
            return []
        return [
            RecommendationDTO(
                priority=RecommendationPriority.HIGH,
                message=f"La calidad semántica ({context.quality.overall_score:.1f}/10) necesita mejorar. Revise las dimensiones con puntuación baja.",
            )
        ]
