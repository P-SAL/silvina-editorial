"""
models.py
Domain models for Silvina Editorial Assistant v0.7
Defines core data structures used throughout the application.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional
from datetime import datetime
from .enums import CitationType, ClassificationCategory, QualityLevel, SectionType
from domain.enums import ArticleType, ArticleSize


@dataclass
class Citation:
    """Represents a citation found in the document."""
    text: str
    citation_type: CitationType
    location: int  # Paragraph index where found
    author: Optional[str] = None
    year: Optional[str] = None
    
    def __str__(self):
        return f"Citation({self.text[:50]}...)"


@dataclass
class Reference:
    """Represents a reference in the bibliography."""
    text: str
    authors: Optional[str] = None
    year: Optional[str] = None
    title: Optional[str] = None
    source: Optional[str] = None
    
    def __str__(self):
        return f"Reference({self.authors}, {self.year})"


@dataclass
class DocumentContent:
    """Represents the extracted content of a document."""
    word_count: int
    char_count: int
    paragraph_count: int = 0  

    title: Optional[str] = None
    authors: Optional[str] = None
    abstract: Optional[str] = None
    keywords: List[str] = field(default_factory=list)
    references: List[Reference] = field(default_factory=list)
    paragraphs: List[str] = field(default_factory=list)
    sections: Dict[str, str] = field(default_factory=dict)

    def __post_init__(self):
        """Calculate word count if not provided."""
        if self.word_count == 0 and self.paragraphs:
            self.word_count = sum(len(p.split()) for p in self.paragraphs)


@dataclass
class ClassificationResult:
    """Result of article classification."""
    article_type: ArticleType  
    article_size: ArticleSize  
    confidence: Optional[float]
    reasoning: str
    timestamp: datetime = field(default_factory=datetime.now)
    
    def __str__(self):
        conf = f"{self.confidence:.1%}" if self.confidence is not None else "—"
        return (
            f"Classification: {self.article_type.value} | "
            f"Size: {self.article_size.value} | "
            f"Confidence: {conf}"
        )
    

@dataclass
class QualityAnalysisResult:
    overall_score: float
    quality_level: QualityLevel
    dimension_scores: Dict[str, Dict]
       
@dataclass
class QualityResult:
    """Result of quality analysis."""
    overall_score: float
    quality_level: QualityLevel
    dimension_scores: Dict[str, Dict[str, any]] = field(default_factory=dict)
    timestamp: datetime = field(default_factory=datetime.now)
    
    def __str__(self):
        return f"Quality: {self.overall_score:.1f}/10 ({self.quality_level.value})"


@dataclass
class StructureValidationResult:
    """Result of structure validation."""
    is_valid: bool
    missing_sections: List[str] = field(default_factory=list)
    section_details: Dict[str, Dict[str, any]] = field(default_factory=dict)
    timestamp: datetime = field(default_factory=datetime.now)
    
    def __str__(self):
        status = "Valid" if self.is_valid else f"Invalid ({len(self.missing_sections)} missing)"
        return f"Structure: {status}"


@dataclass
class CitationAnalysisResult:
    """Result of citation analysis."""
    total_citations: int
    total_references: int
    matched_count: int
    unmatched_count: int
    citations_by_type: Dict[str, int] = field(default_factory=dict)
    unmatched_citations: List[str] = field(default_factory=list)
    timestamp: datetime = field(default_factory=datetime.now)
    
    def __str__(self):
        match_rate = (self.matched_count / self.total_citations * 100 
                     if self.total_citations > 0 else 0)
        return f"Citations: {self.total_citations} ({match_rate:.1f}% matched)"


@dataclass
class AnalysisResult:
    """Complete analysis result for a document."""
    filename: str
    document_content: DocumentContent
    classification: ClassificationResult
    quality: QualityResult
    structure: StructureValidationResult
    citations: CitationAnalysisResult
    timestamp: datetime = field(default_factory=datetime.now)
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            'filename': self.filename,
            'timestamp': self.timestamp.isoformat(),
            'classification': {
                'category': self.classification.article_type.value,
                'confidence': self.classification.confidence,
                'reasoning': self.classification.reasoning
            },
            'quality': {
                'overall_score': self.quality.overall_score,
                'quality_level': self.quality.quality_level.value,
                'dimension_scores': self.quality.dimension_scores
            },
            'structure': {
                'is_valid': self.structure.is_valid,
                'missing_sections': self.structure.missing_sections,
                'section_details': self.structure.section_details
            },
            'citations': {
                'total_citations': self.citations.total_citations,
                'total_references': self.citations.total_references,
                'matched_count': self.citations.matched_count,
                'unmatched_count': self.citations.unmatched_count,
                'citations_by_type': self.citations.citations_by_type,
                'unmatched_citations': self.citations.unmatched_citations
            }
        }
    
    def __str__(self):
        return f"""
Analysis Result for {self.filename}:
  {self.classification}
  {self.quality}
  {self.structure}
  {self.citations}
        """.strip()


# Helper functions for creating instances

def create_empty_document() -> DocumentContent:
    """Create an empty document content instance."""
    return DocumentContent()


def create_classification_result(
    category: ClassificationCategory,
    confidence: float,
    reasoning: str
) -> ClassificationResult:
    """Create a classification result."""
    return ClassificationResult(
        category=category,
        confidence=confidence,
        reasoning=reasoning
    )


def create_quality_result(
    overall_score: float,
    quality_level: QualityLevel,
    dimension_scores: Dict[str, Dict[str, any]]
) -> QualityResult:
    """Create a quality result."""
    return QualityResult(
        overall_score=overall_score,
        quality_level=quality_level,
        dimension_scores=dimension_scores
    )

@dataclass
class Section:
    """Represents a section in an academic document"""
    title: str
    content: str
    section_type: Optional[SectionType] = None
    start_position: int = 0
    end_position: int = 0
    level: int = 1  # Heading level (1, 2, 3, etc.)
    
    def __post_init__(self):
        """Validate section data after initialization"""
        if not self.title:
            raise ValueError("Section title cannot be empty")
        
        # Auto-detect section type if not provided
        if self.section_type is None:
            from domain.enums import classify_section_by_name
            self.section_type = classify_section_by_name(self.title)
    
    def get_word_count(self) -> int:
        """Get word count of section content"""
        return len(self.content.split())
    
    def is_empty(self) -> bool:
        """Check if section has no content"""
        return len(self.content.strip()) == 0