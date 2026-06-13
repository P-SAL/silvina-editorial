from enum import Enum


class ArticleType(Enum):
    """Article type classification."""

    CIENTIFICO = "científico"
    DIVULGACION = "divulgación"
    OPINION = "opinión"
    UNKNOWN = "unknown"
