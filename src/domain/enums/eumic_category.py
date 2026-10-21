from enum import Enum


class EumicCategory(str, Enum):
    FORMAT = "Formato General"
    FIGURES = "Figuras"
    TABLES = "Tablas"
    FORMULAS = "Fórmulas"
    ABSTRACT_KEYWORDS = "Resumen y Palabras Clave"
