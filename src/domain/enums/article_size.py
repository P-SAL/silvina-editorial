from enum import Enum


class ArticleSize(Enum):
    """Article size classification based on character count."""

    LARGO = "largo"
    CORTO = "corto"
    NO_DEFINIDO = "no_definido"
    FUERA_RANGO = "fuera_rango"
