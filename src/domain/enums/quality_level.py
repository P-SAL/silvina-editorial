from enum import Enum


class QualityLevel(Enum):
    """Quality levels for article assessment."""

    EXCELLENT = "Excelente"  # 9.0 - 10.0
    GOOD = "Bueno"  # 7.0 - 8.9
    ACCEPTABLE = "Aceptable"  # 5.0 - 6.9
    NEEDS_IMPROVEMENT = "Requiere mejoras"  # 3.0 - 4.9
    POOR = "Deficiente"  # 0.0 - 2.9


_QUALITY_LEVEL_SCORE_THRESHOLDS: dict[QualityLevel, float] = {
    QualityLevel.EXCELLENT: 9.0,
    QualityLevel.GOOD: 7.0,
    QualityLevel.ACCEPTABLE: 5.0,
    QualityLevel.NEEDS_IMPROVEMENT: 3.0,
}


def get_quality_level_from_score(score: float) -> QualityLevel:
    """Map an overall quality score to its corresponding QualityLevel."""
    for level, threshold in _QUALITY_LEVEL_SCORE_THRESHOLDS.items():
        if score >= threshold:
            return level
    return QualityLevel.POOR
