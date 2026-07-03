from src.domain.enums.quality_level import QualityLevel


class QualityLevelResolver:
    """Domain service mapping overall score to QualityLevel."""

    def resolve(self, score: float) -> QualityLevel:
        if score >= 9.0:
            return QualityLevel.EXCELLENT
        if score >= 7.0:
            return QualityLevel.GOOD
        if score >= 5.0:
            return QualityLevel.ACCEPTABLE
        if score >= 3.0:
            return QualityLevel.NEEDS_IMPROVEMENT
        return QualityLevel.POOR
