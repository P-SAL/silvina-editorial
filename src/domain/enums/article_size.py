from enum import Enum


class ArticleSize(Enum):
    """Article size classification based on character count."""

    LARGO = "largo"  # 36,000 - 40,000 chars
    CORTO = "corto"  # 16,000 - 24,000 chars
    NO_DEFINIDO = "no_definido"  # 24,001 - 35,999 chars
    FUERA_RANGO = "fuera_rango"  # Outside all ranges
