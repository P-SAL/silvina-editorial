from enum import Enum


class SectionType(Enum):
    """Common sections in academic articles."""

    TITLE = "title"
    ABSTRACT = "abstract"
    RESUMEN = "resumen"
    KEYWORDS = "keywords"
    PALABRAS_CLAVE = "palabras_clave"
    INTRODUCTION = "introduction"
    INTRODUCCION = "introduccion"
    METHODOLOGY = "methodology"
    METODOLOGIA = "metodologia"
    RESULTS = "results"
    RESULTADOS = "resultados"
    DISCUSSION = "discussion"
    DISCUSION = "discusion"
    CONCLUSIONS = "conclusions"
    CONCLUSIONES = "conclusiones"
    REFERENCES = "references"
    REFERENCIAS = "referencias"
    BIBLIOGRAPHY = "bibliography"
    BIBLIOGRAFIA = "bibliografia"
    ACKNOWLEDGMENTS = "acknowledgments"
    AGRADECIMIENTOS = "agradecimientos"
    APPENDIX = "appendix"
    ANEXO = "anexo"
