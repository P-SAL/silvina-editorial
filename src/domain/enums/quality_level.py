from enum import Enum


class QualityLevel(Enum):
    """Quality levels for article assessment."""

    EXCELLENT = "Excelente"  # 9.0 - 10.0
    GOOD = "Bueno"  # 7.0 - 8.9
    ACCEPTABLE = "Aceptable"  # 5.0 - 6.9
    NEEDS_IMPROVEMENT = "Requiere mejoras"  # 3.0 - 4.9
    POOR = "Deficiente"  # 0.0 - 2.9


def get_quality_level_from_score(score: float) -> QualityLevel:
    """Map a numeric overall score to its corresponding QualityLevel."""
    if score >= 9.0:
        return QualityLevel.EXCELLENT
    if score >= 7.0:
        return QualityLevel.GOOD
    if score >= 5.0:
        return QualityLevel.ACCEPTABLE
    if score >= 3.0:
        return QualityLevel.NEEDS_IMPROVEMENT
    return QualityLevel.POOR
