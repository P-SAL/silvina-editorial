from enum import Enum


class AnalysisDimension(Enum):
    """Dimensions evaluated in quality analysis."""

    ACADEMIC_RIGOR = "academic_rigor"
    METHODOLOGICAL_CLARITY = "methodological_clarity"
    ARGUMENTATION = "argumentation"
    LITERATURE_REVIEW = "literature_review"
    ORIGINALITY = "originality"
    WRITING_QUALITY = "writing_quality"
    STRUCTURE = "structure"
    CITATION_QUALITY = "citation_quality"
