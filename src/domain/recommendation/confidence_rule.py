from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.recommendation.analysis_context import AnalysisContext
from src.domain.recommendation.recommendation_rule import RecommendationRule


class ConfidenceRule(RecommendationRule):
    def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]:
        confidence = context.classification.confidence
        if confidence is None or confidence >= context.settings.classification_confidence_threshold:
            return []
        return [
            RecommendationDTO(
                priority=RecommendationPriority.LOW,
                message=f"La clasificación tiene confianza baja ({confidence:.1%}). Verifique que el documento siga la estructura típica de su categoría.",
            )
        ]
