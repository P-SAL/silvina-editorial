from src.domain.dtos.quality_level_thresholds_dto import QualityLevelThresholdsDTO
from src.domain.enums.quality_level import QualityLevel


class QualityLevelResolver:
    """Domain service mapping overall score to QualityLevel."""

    def __init__(self, *, thresholds: QualityLevelThresholdsDTO) -> None:
        self._thresholds = thresholds

    def resolve(self, score: float) -> QualityLevel:
        if score >= self._thresholds.excellent_threshold:
            return QualityLevel.EXCELLENT
        if score >= self._thresholds.good_threshold:
            return QualityLevel.GOOD
        if score >= self._thresholds.acceptable_threshold:
            return QualityLevel.ACCEPTABLE
        if score >= self._thresholds.needs_improvement_threshold:
            return QualityLevel.NEEDS_IMPROVEMENT
        return QualityLevel.POOR
