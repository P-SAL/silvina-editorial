from enum import Enum


class QualityLevel(Enum):
    """Quality levels for article assessment."""

    EXCELLENT = "Excelente"  # 9.0 - 10.0
    GOOD = "Bueno"  # 7.0 - 8.9
    ACCEPTABLE = "Aceptable"  # 5.0 - 6.9
    NEEDS_IMPROVEMENT = "Requiere mejoras"  # 3.0 - 4.9
    POOR = "Deficiente"  # 0.0 - 2.9
