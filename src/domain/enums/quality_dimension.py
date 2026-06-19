from enum import Enum


class QualityDimension(Enum):
    """The 4 semantic dimensions scored during quality analysis."""

    CLARITY = "claridad"
    COHERENCE = "coherencia"
    ARGUMENTATION = "argumentacion"
    CONCLUSIONS = "conclusiones"
