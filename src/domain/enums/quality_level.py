from enum import Enum


class QualityLevel(Enum):
    """Quality levels for article assessment."""

    EXCELLENT = "Excelente"
    GOOD = "Bueno"
    ACCEPTABLE = "Aceptable"
    NEEDS_IMPROVEMENT = "Requiere mejoras"
    POOR = "Deficiente"
