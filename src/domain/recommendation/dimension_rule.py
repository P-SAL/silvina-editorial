from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.recommendation.analysis_context import AnalysisContext
from src.domain.recommendation.recommendation_rule import RecommendationRule


class DimensionRule(RecommendationRule):
    def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]:
        recommendations = []
        for dimension_name, dim_data in context.quality.dimension_scores.items():
            score = dim_data["score"]
            if score < context.settings.dimension_threshold:
                feedback = dim_data.get("feedback", "")
                recommendations.append(
                    RecommendationDTO(
                        priority=RecommendationPriority.MEDIUM,
                        message=f'Dimensión "{dimension_name}" tiene puntuación baja ({score:.1f}). {feedback}',
                    )
                )
        return recommendations
