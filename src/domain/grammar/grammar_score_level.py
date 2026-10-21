from enum import Enum


class GrammarScoreLevel(Enum):
    PERFECT = (0, 10.0, "Sin errores gramaticales")
    MINOR = (5, 8.5, "Pocos errores gramaticales")
    MODERATE = (15, 7.0, "Errores gramaticales moderados")
    SEVERE = (None, 5.0, "Muchos errores gramaticales")

    def __init__(self, max_errors: int | None, score: float, feedback: str) -> None:
        self.max_errors = max_errors
        self.score = score
        self.feedback = feedback

    @classmethod
    def from_error_count(cls, error_count: int) -> "GrammarScoreLevel":
        for level in cls:
            if level.max_errors is not None and error_count <= level.max_errors:
                return level
        return cls.SEVERE
