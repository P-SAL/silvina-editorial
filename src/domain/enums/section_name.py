from enum import Enum


class SectionName(str, Enum):
    """Canonical section names for academic articles."""

    SUMMARY = "Resumen"
    INTRODUCTION = "Introducción"
    METHODOLOGY = "Metodología"
    RESULTS = "Resultados"
    DISCUSSION = "Discusión"
    ARGUMENTATION = "Argumentación"
    DEVELOPMENT = "Desarrollo"
    CONCLUSIONS = "Conclusiones"
    REFERENCES = "Referencias"
