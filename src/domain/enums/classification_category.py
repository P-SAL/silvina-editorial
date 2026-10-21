from enum import Enum


class ClassificationCategory(Enum):
    """Categories of academic articles according to EUMIC standards."""

    RESEARCH_ARTICLE = "research_article"
    REVIEW_ARTICLE = "review_article"
    REFLECTION_ARTICLE = "reflection_article"
    SHORT_ARTICLE = "short_article"
    CASE_REPORT = "case_report"
    UNKNOWN = "unknown"
