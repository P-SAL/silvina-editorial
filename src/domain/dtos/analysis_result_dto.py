from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResultDTO
from src.domain.dtos.classification_result_dto import ClassificationResultDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.dtos.structure_validation_result_dto import StructureValidationResultDTO


@dataclass(frozen=True)
class AnalysisResultDTO(BaseDTO):
    """Immutable aggregate output DTO for a complete document analysis."""

    filename: str
    document_content: DocumentContentDTO
    classification: ClassificationResultDTO
    quality: QualityResultDTO
    structure: StructureValidationResultDTO
    citations: CitationAnalysisResultDTO
    timestamp: datetime = field(default_factory=datetime.now)

    def to_dict(self) -> dict[str, Any]:
        """Return the flattened analysis result as a plain dictionary.

        Key structure is byte-compatible with the legacy shape consumed by
        report_formatter, word_exporter, and Gradio. Slice 0 preserves the
        exact legacy contract: the 'classification' sub-dict key stays
        'category' (its value is the article type).
        """
        return {
            "filename": self.filename,
            "timestamp": self.timestamp.isoformat(),
            "classification": {
                "category": self.classification.article_type.value,
                "confidence": self.classification.confidence,
                "reasoning": self.classification.reasoning,
            },
            "quality": {
                "overall_score": self.quality.overall_score,
                "quality_level": self.quality.quality_level.value,
                "dimension_scores": self.quality.dimension_scores,
                "editorial_suitability": (
                    self.quality.editorial_suitability.as_dict()
                    if self.quality.editorial_suitability is not None
                    else None
                ),
            },
            "structure": {
                "is_valid": self.structure.is_valid,
                "missing_sections": self.structure.missing_sections,
                "section_details": self.structure.section_details,
            },
            "citations": {
                "total_citations": self.citations.total_citations,
                "total_references": self.citations.total_references,
                "matched_count": self.citations.matched_count,
                "unmatched_count": self.citations.unmatched_count,
                "citations_by_type": self.citations.citations_by_type,
                "unmatched_citations": self.citations.unmatched_citations,
            },
        }

    def __str__(self) -> str:
        """Return human-readable analysis summary."""
        suitability_line = ""
        if self.quality.editorial_suitability is not None:
            suitability_line = f"\n  {self.quality.editorial_suitability}"
        return (
            f"Analysis Result for {self.filename}:\n"
            f"  {self.classification}\n"
            f"  {self.quality}\n"
            f"  {self.structure}\n"
            f"  {self.citations}"
            f"{suitability_line}"
        )
