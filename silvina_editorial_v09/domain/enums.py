"""
enums.py
Enumeration types for Silvina Editorial Assistant v0.7
Defines all enum types used across the application.
"""

from enum import Enum

class ArticleType(Enum):
    """Article type classification."""
    CIENTIFICO = "científico"
    DIVULGACION = "divulgación"
    OPINION = "opinión"
    UNKNOWN = "unknown"

class ArticleSize(Enum):
    """Article size classification based on character count."""
    LARGO = "largo"           # 36,000 - 40,000 chars
    CORTO = "corto"           # 16,000 - 24,000 chars
    NO_DEFINIDO = "no_definido"  # 24,001 - 35,999 chars
    FUERA_RANGO = "fuera_rango"  # Outside all ranges

def classify_article_size(char_count: int) -> ArticleSize:
    """Classify article size based on character count with spaces."""
    if 36000 <= char_count <= 40000:
        return ArticleSize.LARGO
    elif 16000 <= char_count <= 24000:
        return ArticleSize.CORTO
    elif 24001 <= char_count <= 35999:
        return ArticleSize.NO_DEFINIDO
    else:
        return ArticleSize.FUERA_RANGO

class CitationType(Enum):
    """Types of citations found in academic documents."""
    AUTHOR_YEAR = "author_year"  # e.g., (Smith, 2020)
    NUMERIC = "numeric"          # e.g., [1], [2]
    FOOTNOTE = "footnote"        # e.g., superscript numbers
    UNKNOWN = "unknown"


class ClassificationCategory(Enum):
    """Categories of academic articles according to EUMIC standards."""
    RESEARCH_ARTICLE = "research_article"
    REVIEW_ARTICLE = "review_article"
    REFLECTION_ARTICLE = "reflection_article"
    SHORT_ARTICLE = "short_article"
    CASE_REPORT = "case_report"
    UNKNOWN = "unknown"


class QualityLevel(Enum):
    """Quality levels for article assessment."""
    EXCELLENT = "Excelente"              # 9.0 - 10.0
    GOOD = "Bueno"                       # 7.0 - 8.9
    ACCEPTABLE = "Aceptable"             # 5.0 - 6.9
    NEEDS_IMPROVEMENT = "Requiere mejoras"  # 3.0 - 4.9
    POOR = "Deficiente"                  # 0.0 - 2.9


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


class ValidationStatus(Enum):
    """Status of validation checks."""
    PASSED = "passed"
    FAILED = "failed"
    WARNING = "warning"
    NOT_APPLICABLE = "not_applicable"


class RecommendationPriority(Enum):
    """Priority levels for recommendations."""
    HIGH = "alta"
    MEDIUM = "media"
    LOW = "baja"


# Helper functions for enum operations

def get_quality_level_from_score(score: float) -> QualityLevel:
    """
    Convert numeric score to quality level.
    
    Args:
        score: Numeric score (0-10)
        
    Returns:
        Corresponding QualityLevel enum
    """
    if score >= 9.0:
        return QualityLevel.EXCELLENT
    elif score >= 7.0:
        return QualityLevel.GOOD
    elif score >= 5.0:
        return QualityLevel.ACCEPTABLE
    elif score >= 3.0:
        return QualityLevel.NEEDS_IMPROVEMENT
    else:
        return QualityLevel.POOR


def get_citation_type_from_pattern(citation_text: str) -> CitationType:
    """
    Detect citation type from citation text pattern.
    
    Args:
        citation_text: Text of the citation
        
    Returns:
        Detected CitationType enum
    """
    import re
    
    # Author-year pattern: (Author, Year) or (Author et al., Year)
    if re.search(r'\([A-Z][a-z]+.*?\d{4}\)', citation_text):
        return CitationType.AUTHOR_YEAR
    
    # Numeric pattern: [1], [2], etc.
    if re.search(r'\[\d+\]', citation_text):
        return CitationType.NUMERIC
    
    # Footnote pattern: superscript numbers
    if re.search(r'\d+', citation_text) and len(citation_text) < 5:
        return CitationType.FOOTNOTE
    
    return CitationType.UNKNOWN


def classify_section_by_name(section_name: str) -> SectionType:
    """
    Classify section by its name/header.
    
    Args:
        section_name: Name or header of the section
        
    Returns:
        Corresponding SectionType enum or None
    """
    section_lower = section_name.lower().strip()
    
    # Map common variations to section types
    section_mapping = {
        'título': SectionType.TITLE,
        'title': SectionType.TITLE,
        'resumen': SectionType.RESUMEN,
        'abstract': SectionType.ABSTRACT,
        'palabras clave': SectionType.PALABRAS_CLAVE,
        'keywords': SectionType.KEYWORDS,
        'introducción': SectionType.INTRODUCCION,
        'introduction': SectionType.INTRODUCTION,
        'metodología': SectionType.METODOLOGIA,
        'methodology': SectionType.METHODOLOGY,
        'métodos': SectionType.METODOLOGIA,
        'methods': SectionType.METHODOLOGY,
        'resultados': SectionType.RESULTADOS,
        'results': SectionType.RESULTS,
        'discusión': SectionType.DISCUSION,
        'discussion': SectionType.DISCUSSION,
        'conclusiones': SectionType.CONCLUSIONES,
        'conclusions': SectionType.CONCLUSIONS,
        'referencias': SectionType.REFERENCIAS,
        'references': SectionType.REFERENCES,
        'bibliografía': SectionType.BIBLIOGRAFIA,
        'bibliography': SectionType.BIBLIOGRAPHY,
        'agradecimientos': SectionType.AGRADECIMIENTOS,
        'acknowledgments': SectionType.ACKNOWLEDGMENTS,
        'anexo': SectionType.ANEXO,
        'appendix': SectionType.APPENDIX,
    }
    
    # Check for exact matches
    for key, section_type in section_mapping.items():
        if key in section_lower:
            return section_type
    
    return None


def get_required_sections_for_category(category: ClassificationCategory) -> list:
    """
    Get list of required sections for an article category.
    
    Args:
        category: Article classification category
        
    Returns:
        List of required SectionType enums
    """
    required_sections = {
        ClassificationCategory.RESEARCH_ARTICLE: [
            SectionType.RESUMEN,
            SectionType.ABSTRACT,
            SectionType.INTRODUCCION,
            SectionType.METODOLOGIA,
            SectionType.RESULTADOS,
            SectionType.DISCUSION,
            SectionType.CONCLUSIONES,
            SectionType.REFERENCIAS
        ],
        ClassificationCategory.REVIEW_ARTICLE: [
            SectionType.RESUMEN,
            SectionType.ABSTRACT,
            SectionType.INTRODUCCION,
            SectionType.CONCLUSIONES,
            SectionType.REFERENCIAS
        ],
        ClassificationCategory.REFLECTION_ARTICLE: [
            SectionType.RESUMEN,
            SectionType.ABSTRACT,
            SectionType.INTRODUCCION,
            SectionType.CONCLUSIONES,
            SectionType.REFERENCIAS
        ],
        ClassificationCategory.SHORT_ARTICLE: [
            SectionType.RESUMEN,
            SectionType.INTRODUCCION,
            SectionType.CONCLUSIONES,
            SectionType.REFERENCIAS
        ],
        ClassificationCategory.CASE_REPORT: [
            SectionType.RESUMEN,
            SectionType.INTRODUCCION,
            SectionType.CONCLUSIONES,
            SectionType.REFERENCIAS
        ]
    }
    
    return required_sections.get(category, [])


# Export all enums and helper functions
__all__ = [
    'CitationType',
    'ClassificationCategory',
    'QualityLevel',
    'SectionType',
    'AnalysisDimension',
    'ValidationStatus',
    'RecommendationPriority',
    'SeverityLevel',
    'get_quality_level_from_score',
    'get_citation_type_from_pattern',
    'classify_section_by_name',
    'get_required_sections_for_category'
]

class SeverityLevel(Enum):
    """Severity levels for validation issues"""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"