from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.enums.recommendation_priority import RecommendationPriority
from src.domain.recommendation.analysis_context import AnalysisContext
from src.domain.recommendation.recommendation_rule import RecommendationRule


class StructureRule(RecommendationRule):
    def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]:
        if context.structure.is_valid:
            return []
        return [
            RecommendationDTO(
                priority=RecommendationPriority.HIGH,
                message=f'Falta la sección requerida: "{section}". Complete esta sección según las normas EUMIC.',
            )
            for section in context.structure.missing_sections
        ]
