from src.domain.enums.quality_level import QualityLevel


class QualityLevelResolver:
    """Domain service mapping overall score to QualityLevel."""

    def __init__(
        self,
        *,
        excellent_threshold: float = 9.0,
        good_threshold: float = 7.0,
        acceptable_threshold: float = 5.0,
        needs_improvement_threshold: float = 3.0,
    ) -> None:
        self._excellent_threshold = excellent_threshold
        self._good_threshold = good_threshold
        self._acceptable_threshold = acceptable_threshold
        self._needs_improvement_threshold = needs_improvement_threshold

    def resolve(self, score: float) -> QualityLevel:
        if score >= self._excellent_threshold:
            return QualityLevel.EXCELLENT
        if score >= self._good_threshold:
            return QualityLevel.GOOD
        if score >= self._acceptable_threshold:
            return QualityLevel.ACCEPTABLE
        if score >= self._needs_improvement_threshold:
            return QualityLevel.NEEDS_IMPROVEMENT
        return QualityLevel.POOR
