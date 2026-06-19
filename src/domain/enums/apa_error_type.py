from enum import Enum


class ApaErrorType(str, Enum):
    CONJUNCTION_ERROR = "Conjunción incorrecta"
    COMMA_ERROR = "Puntuación incorrecta"
    CAPITALIZATION_ERROR = "Mayúsculas/minúsculas incorrectas"
    ET_AL_FORMAT_ERROR = "Formato 'et al.' incorrecto"
    PAGE_FORMAT_ERROR = "Formato de página incorrecto"
    SPACING_ERROR = "Espaciado incorrecto"
    YEAR_FORMAT_ERROR = "Formato de año incorrecto"
    PARENTHESES_ERROR = "Paréntesis incorrectos"
