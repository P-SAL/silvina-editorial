from enum import Enum


class QualityDimension(Enum):
    """The 4 semantic dimensions scored during quality analysis."""

    CLARIDAD = "claridad"
    COHERENCIA = "coherencia"
    ARGUMENTACION = "argumentacion"
    CONCLUSIONES = "conclusiones"
