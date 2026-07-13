from enum import Enum


class QualityLevel(Enum):
    """Quality levels for article assessment."""

    EXCELLENT = ("Excelente", 9.0)
    GOOD = ("Bueno", 7.0)
    ACCEPTABLE = ("Aceptable", 5.0)
    NEEDS_IMPROVEMENT = ("Requiere mejoras", 3.0)
    POOR = ("Deficiente", 0.0)

    def __new__(cls, label: str, min_threshold: float):
        obj = object.__new__(cls)
        obj._value_ = label
        obj.min_threshold = min_threshold
        return obj

    @classmethod
    def from_score(cls, score: float) -> "QualityLevel":
        """Map an overall quality score to its corresponding QualityLevel."""
        for level in sorted(cls, key=lambda x: x.min_threshold, reverse=True):
            if score >= level.min_threshold:
                return level
        return cls.POOR
