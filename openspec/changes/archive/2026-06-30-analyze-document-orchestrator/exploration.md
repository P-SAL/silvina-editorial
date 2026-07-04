# Exploration: analyze-document-orchestrator

**Change**: analyze-document-orchestrator — Slice 13 hexagonal orchestration of `analyze_document`
**Date**: 2026-06-30
**Status**: Proposed

---

## 1. Context & Objectives

The goal of Slice 13 is to implement the top-level orchestration for complete document analysis, migrating the logic from legacy `SilvinaEditorialAssistant.analyze_document()` and `_generate_recommendations()` in [main.py](file:///E:/Python/silvina-editorial/main.py) into the clean hexagonal architecture codebase:
1. **`RecommendationDTO`**: Immutable DTO representing a single quality or formatting recommendation.
2. **`RecommendationBuilder`**: Domain service that encapsulates recommendation generation logic based on analysis inputs.
3. **`AnalyzeDocumentUseCase`**: Application orchestrator coordinating reading, content extraction, citation parsing, APA validation, grammar checking, classification, quality analysis, structure validation, citation matching, and EUMIC format verification.
4. **`AnalyzeDocumentUseCaseWiring`**: Assembly factory wiring all sub-use cases and dependencies.

---

## 2. Domain Enum & DTO Changes

### A. Extend `RecommendationPriority`

The existing [recommendation_priority.py](file:///E:/Python/silvina-editorial/src/domain/enums/recommendation_priority.py) only has:
- `HIGH = "alta"`
- `MEDIUM = "media"`
- `LOW = "baja"`

However, the legacy recommendation system uses three additional publication-readiness levels: `"critica"`, `"advertencia"`, and `"aprobado"`. To encapsulate this in the domain enum, we will extend `RecommendationPriority`:

File: `src/domain/enums/recommendation_priority.py`
```python
from enum import Enum


class RecommendationPriority(Enum):
    """Priority levels for recommendations."""

    HIGH = "alta"
    MEDIUM = "media"
    LOW = "baja"
    CRITICAL = "critica"
    WARNING = "advertencia"
    APPROVED = "aprobado"
```

We will update [test_recommendation_priority.py](file:///E:/Python/silvina-editorial/src/domain/tests/enums/test_recommendation_priority.py) to assert all 6 enum values and the member count of 6.

### B. New `RecommendationDTO`

We will create a new DTO under `src/domain/dtos/recommendation_dto.py`:

File: `src/domain/dtos/recommendation_dto.py`
```python
from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.recommendation_priority import RecommendationPriority


@dataclass(frozen=True)
class RecommendationDTO(BaseDTO):
    """Immutable data transfer object representing a single recommendation."""

    priority: RecommendationPriority
    message: str
```

### C. Update `ReportInputDTO`

We will replace the placeholder type `list` for the `recommendations` field in [report_input_dto.py](file:///E:/Python/silvina-editorial/src/domain/dtos/report_input_dto.py):

File: `src/domain/dtos/report_input_dto.py`
```python
from dataclasses import dataclass

from src.domain.dtos.apa_validation_result_dto import ApaValidationResultDTO
from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResultDTO
from src.domain.dtos.classification_result_dto import ClassificationResultDTO
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.dtos.structure_validation_result_dto import StructureValidationResultDTO

_PUBLISH_THRESHOLD = 7.0


@dataclass(frozen=True)
class ReportInputDTO(BaseDTO):
    """Aggregates all analysis results required to generate an export report."""

    filename: str
    document_content: DocumentContentDTO
    classification: ClassificationResultDTO
    quality: QualityResultDTO
    grammar: GrammarCheckResultDTO
    structure: StructureValidationResultDTO
    citations: CitationAnalysisResultDTO
    apa_validation: ApaValidationResultDTO
    recommendations: list[RecommendationDTO]
```

---

## 3. Domain Service: RecommendationBuilder

The recommendation generation is pure business logic, which must reside in the domain layer. We will create the domain service `RecommendationBuilder` class inside `src/domain/recommendation/recommendation_builder.py`.

We will adhere to strict **no-abbreviations** guidelines:
- Avoid loops using `rec`, use `recommendation` or `missing_section` instead.
- Avoid using `dim`, use `dimension_name` and `dimension_data`.

File: `src/domain/recommendation/recommendation_builder.py`
```python
from src.domain.dtos.classification_result_dto import ClassificationResultDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.dtos.structure_validation_result_dto import StructureValidationResultDTO
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResultDTO
from src.domain.dtos.apa_validation_result_dto import ApaValidationResultDTO
from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.dtos.recommendation_dto import RecommendationDTO
from src.domain.enums.recommendation_priority import RecommendationPriority


class RecommendationBuilder:
    """Domain service for generating editorial and quality recommendations for documents."""

    def build(
        self,
        classification: ClassificationResultDTO,
        quality: QualityResultDTO,
        structure: StructureValidationResultDTO,
        citations: CitationAnalysisResultDTO,
        apa_validation: ApaValidationResultDTO,
        grammar: GrammarCheckResultDTO,
    ) -> list[RecommendationDTO]:
        """Generate a list of recommendations based on document analysis results."""
        recommendations: list[RecommendationDTO] = []

        # 1. Quality recommendations (semantic only)
        if quality.overall_score < 7.0:
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.HIGH,
                    message=(
                        f"La calidad semántica ({quality.overall_score:.1f}/10) necesita mejorar. "
                        "Revise las dimensiones con puntuación baja."
                    ),
                )
            )

        # 2. Grammar recommendations
        if grammar.score < 7.0:
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.HIGH,
                    message=f"Gramática ({grammar.score:.1f}/10) requiere corrección.",
                )
            )

        # 3. Check individual dimensions
        for dimension_name, dimension_data in quality.dimension_scores.items():
            score = dimension_data.get("score", 0.0)
            if score < 6.0:
                feedback = dimension_data.get("feedback", "Requiere atención.")
                recommendations.append(
                    RecommendationDTO(
                        priority=RecommendationPriority.MEDIUM,
                        message=f'Dimensión "{dimension_name}" tiene puntuación baja ({score:.1f}). {feedback}',
                    )
                )

        # 4. Structure recommendations
        if not structure.is_valid:
            for missing_section in structure.missing_sections:
                recommendations.append(
                    RecommendationDTO(
                        priority=RecommendationPriority.HIGH,
                        message=(
                            f'Falta la sección requerida: "{missing_section}". '
                            "Complete esta sección según las normas EUMIC."
                        ),
                    )
                )

        # 5. Citation recommendations
        total_citations = citations.total_citations
        matched_count = citations.matched_count
        match_rate = (matched_count / total_citations * 100.0) if total_citations > 0 else 100.0

        if match_rate < 90.0:
            unmatched_list = citations.unmatched_citations[:10]
            unmatched_string = "; ".join(unmatched_list) if unmatched_list else ""
            message = (
                f"Tasa de coincidencia de citas baja ({match_rate:.1f}%). "
                f"{citations.unmatched_count} citas no tienen referencia correspondiente."
            )
            if unmatched_string:
                message += f" Citas sin referencia: {unmatched_string}"
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.HIGH,
                    message=message,
                )
            )
        elif citations.unmatched_count > 0:
            unmatched_list = citations.unmatched_citations[:10]
            unmatched_string = "; ".join(unmatched_list) if unmatched_list else ""
            message = f"{citations.unmatched_count} citas no tienen referencia correspondiente."
            if unmatched_string:
                message += f" Citas sin referencia: {unmatched_string}"
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.MEDIUM,
                    message=message,
                )
            )

        if total_citations < 10:
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.MEDIUM,
                    message=(
                        f"Número bajo de citas ({total_citations}). "
                        "Considere ampliar el marco teórico con más referencias."
                    ),
                )
            )

        # 6. Classification confidence
        if classification.confidence is not None and classification.confidence < 0.7:
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.LOW,
                    message=(
                        f"La clasificación tiene confianza baja ({classification.confidence:.1%}). "
                        "Verifique que el documento siga la estructura típica de su categoría."
                    ),
                )
            )

        # 7. Final publication recommendation
        has_critical_issues = False
        has_warnings = False

        if quality.overall_score < 5.0:
            has_critical_issues = True
        if grammar.score < 5.0:
            has_critical_issues = True
        if not structure.is_valid:
            has_critical_issues = True
        if match_rate < 50.0:
            has_critical_issues = True

        if quality.overall_score < 7.0 or grammar.score < 7.0:
            has_warnings = True
        if match_rate < 90.0:
            has_warnings = True
        if len(apa_validation.violations) > 0:
            has_warnings = True

        if has_critical_issues:
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.CRITICAL,
                    message="❌ NO APTO PARA PUBLICACIÓN. El documento presenta errores críticos que deben corregirse.",
                )
            )
        elif total_citations == 0:
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.CRITICAL,
                    message="❌ NO APTO PARA PUBLICACIÓN. No se detectaron citas APA en el texto. Verifique el formato de citación según normas APA 7.",
                )
            )
        elif has_warnings:
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.WARNING,
                    message="⚠️ REQUIERE REVISIÓN antes de publicación. Corrija los problemas identificados.",
                )
            )
        else:
            recommendations.append(
                RecommendationDTO(
                    priority=RecommendationPriority.APPROVED,
                    message="✅ APTO PARA PUBLICACIÓN. El documento cumple con los estándares de calidad.",
                )
            )

        return recommendations
```

---

## 4. Use Case Design: AnalyzeDocumentUseCase

This orchestrator coordinates reading paragraphs, content extraction, citations/references extraction, APA validation, grammar checking, classification, quality analysis, structure validation, citation matching, and recommendations building.

It receives the document path, delegates steps, builds the domain entities, and returns the aggregated output in a `ReportInputDTO`.

File: `src/application/analyze_document_use_case.py`
```python
from pathlib import Path

from src.application.analyze_quality_use_case import AnalyzeQualityUseCase
from src.application.check_grammar_use_case import CheckGrammarUseCase
from src.application.classify_article_use_case import ClassifyArticleUseCase
from src.application.extract_citations_use_case import ExtractCitationsUseCase
from src.application.extract_content_use_case import ExtractContentUseCase
from src.application.match_citations_use_case import MatchCitationsUseCase
from src.application.read_document_use_case import ReadDocumentUseCase
from src.application.validate_apa_use_case import ValidateApaUseCase
from src.application.validate_structure_use_case import ValidateStructureUseCase
from src.application.verify_eumic_use_case import VerifyEumicUseCase
from src.domain.dtos.report_input_dto import ReportInputDTO
from src.domain.enums.article_type import ArticleType
from src.domain.enums.citation_type import CitationType
from src.domain.enums.section_name import SectionName
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.recommendation.recommendation_builder import RecommendationBuilder


class AnalyzeDocumentUseCase:
    """Orchestrates the entire document analysis pipeline."""

    def __init__(
        self,
        read_document_use_case: ReadDocumentUseCase,
        extract_content_use_case: ExtractContentUseCase,
        extract_citations_use_case: ExtractCitationsUseCase,
        validate_apa_use_case: ValidateApaUseCase,
        check_grammar_use_case: CheckGrammarUseCase,
        classify_article_use_case: ClassifyArticleUseCase,
        analyze_quality_use_case: AnalyzeQualityUseCase,
        validate_structure_use_case: ValidateStructureUseCase,
        match_citations_use_case: MatchCitationsUseCase,
        verify_eumic_use_case: VerifyEumicUseCase,
        recommendation_builder: RecommendationBuilder,
    ) -> None:
        self._read_document_use_case = read_document_use_case
        self._extract_content_use_case = extract_content_use_case
        self._extract_citations_use_case = extract_citations_use_case
        self._validate_apa_use_case = validate_apa_use_case
        self._check_grammar_use_case = check_grammar_use_case
        self._classify_article_use_case = classify_article_use_case
        self._analyze_quality_use_case = analyze_quality_use_case
        self._validate_structure_use_case = validate_structure_use_case
        self._match_citations_use_case = match_citations_use_case
        self._verify_eumic_use_case = verify_eumic_use_case
        self._recommendation_builder = recommendation_builder

    @generic_error_handler
    def execute(self, document_path: str) -> ReportInputDTO:
        """Run all analysis steps for a document and return aggregated results."""
        # 1. Read document paragraphs
        paragraphs = self._read_document_use_case.execute(path=document_path)

        # 2. Extract structured content
        document_content = self._extract_content_use_case.execute(
            paragraphs=paragraphs, path=document_path
        )

        # 3. Extract citations and references
        citation_extraction = self._extract_citations_use_case.execute(docx_path=document_path)

        # 4. Validate APA citations (only validate AUTHOR_YEAR citation types)
        citation_tuples = [
            (
                citation.text,
                citation.location,
                paragraphs[citation.location] if citation.location < len(paragraphs) else "",
            )
            for citation in citation_extraction.citations
            if citation.citation_type == CitationType.AUTHOR_YEAR
        ]
        apa_validation = self._validate_apa_use_case.execute(citations=citation_tuples)

        # 5. Check grammar
        grammar = self._check_grammar_use_case.execute(paragraphs=paragraphs)

        # 6. Classify article type
        classification = self._classify_article_use_case.execute(
            document_content=document_content
        )

        # 7. Analyze quality
        quality = self._analyze_quality_use_case.execute(
            document_content=document_content, article_type=classification.article_type
        )

        # 8. Validate structure
        # Apply the IMRyD scientific classification override logic
        is_imryd_classified = classification.article_type == ArticleType.CIENTIFICO and (
            "IMRyD" in (classification.reasoning or "")
        )
        effective_article_type = (
            classification.article_type
            if is_imryd_classified
            else (
                ArticleType.DIVULGACION
                if classification.article_type == ArticleType.CIENTIFICO
                else classification.article_type
            )
        )
        structure = self._validate_structure_use_case.execute(
            document_content=document_content,
            article_type=effective_article_type,
            has_references=len(citation_extraction.references) > 0,
        )

        # 9. Match citations
        try:
            matched_section_type = SectionName(citation_extraction.section_type)
        except ValueError:
            matched_section_type = SectionName.REFERENCES

        citations_analysis = self._match_citations_use_case.execute(
            citations=citation_extraction.citations,
            references=citation_extraction.references,
            section_type=matched_section_type,
        )

        # 10. Verify EUMIC compliance
        self._verify_eumic_use_case.execute(
            docx_path=document_path, word_count=document_content.word_count
        )

        # 11. Build recommendations
        recommendations = self._recommendation_builder.build(
            classification=classification,
            quality=quality,
            structure=structure,
            citations=citations_analysis,
            apa_validation=apa_validation,
            grammar=grammar,
        )

        # 12. Aggregate in ReportInputDTO
        return ReportInputDTO(
            filename=Path(document_path).name,
            document_content=document_content,
            classification=classification,
            quality=quality,
            grammar=grammar,
            structure=structure,
            citations=citations_analysis,
            apa_validation=apa_validation,
            recommendations=recommendations,
        )
```

---

## 5. Wiring Design: AnalyzeDocumentUseCaseWiring

Every use case has a wiring file inside `src/infrastructure/wirings/`. We will create the wiring factory class for `AnalyzeDocumentUseCase` which instantiates and injects all required sub-use cases and the domain service `RecommendationBuilder`.

File: `src/infrastructure/wirings/analyze_document_use_case_wiring.py`
```python
from src.application.analyze_document_use_case import AnalyzeDocumentUseCase
from src.domain.recommendation.recommendation_builder import RecommendationBuilder
from src.infrastructure.wirings.analyze_quality_use_case_wiring import AnalyzeQualityUseCaseWiring
from src.infrastructure.wirings.check_grammar_use_case_wiring import CheckGrammarUseCaseWiring
from src.infrastructure.wirings.classify_article_use_case_wiring import (
    ClassifyArticleUseCaseWiring,
)
from src.infrastructure.wirings.extract_citations_use_case_wiring import (
    ExtractCitationsUseCaseWiring,
)
from src.infrastructure.wirings.extract_content_use_case_wiring import (
    ExtractContentUseCaseWiring,
)
from src.infrastructure.wirings.match_citations_use_case_wiring import (
    MatchCitationsUseCaseWiring,
)
from src.infrastructure.wirings.read_document_use_case_wiring import ReadDocumentUseCaseWiring
from src.infrastructure.wirings.validate_apa_wiring import ValidateApaWiring
from src.infrastructure.wirings.validate_structure_wiring import ValidateStructureWiring
from src.infrastructure.wirings.verify_eumic_use_case_wiring import VerifyEumicUseCaseWiring


class AnalyzeDocumentUseCaseWiring:
    """Factory for building a ready-to-use AnalyzeDocumentUseCase."""

    def create_use_case(self) -> AnalyzeDocumentUseCase:
        """Assemble and return a fully wired AnalyzeDocumentUseCase instance."""
        return AnalyzeDocumentUseCase(
            read_document_use_case=ReadDocumentUseCaseWiring().create_use_case(),
            extract_content_use_case=ExtractContentUseCaseWiring().create_use_case(),
            extract_citations_use_case=ExtractCitationsUseCaseWiring().create_use_case(),
            validate_apa_use_case=ValidateApaWiring().create_use_case(),
            check_grammar_use_case=CheckGrammarUseCaseWiring().create_use_case(),
            classify_article_use_case=ClassifyArticleUseCaseWiring().create_use_case(),
            analyze_quality_use_case=AnalyzeQualityUseCaseWiring().create_use_case(),
            validate_structure_use_case=ValidateStructureWiring().create_use_case(),
            match_citations_use_case=MatchCitationsUseCaseWiring().create_use_case(),
            verify_eumic_use_case=VerifyEumicUseCaseWiring().create_use_case(),
            recommendation_builder=RecommendationBuilder(),
        )
```

---

## 6. Adapter & Fixtures Refactoring

### A. Modify `DocxReportAdapter`

We must change dict-based access to attribute-based access in [docx_report_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_adapter.py) for the recommendations:

```python
    def _add_recommendations(self, doc, report_input: ReportInputDTO) -> None:
        if not report_input.recommendations:
            return

        heading = doc.add_heading("💡 RECOMENDACIONES", 1)
        for run in heading.runs:
            run.font.color.rgb = RGBColor(*self._settings.heading_color_rgb)

        recommendations = report_input.recommendations

        final_recommendations = [
            recommendation
            for recommendation in recommendations
            if recommendation.priority in [
                RecommendationPriority.CRITICAL,
                RecommendationPriority.WARNING,
                RecommendationPriority.APPROVED,
            ]
        ]
        final_priority_colors = {
            RecommendationPriority.CRITICAL: self._settings.reject_color_rgb,
            RecommendationPriority.WARNING: self._settings.reject_color_rgb,
            RecommendationPriority.APPROVED: self._settings.publishable_color_rgb,
        }
        if final_recommendations:
            recommendation = final_recommendations[0]
            paragraph = doc.add_paragraph()
            paragraph.add_run(recommendation.message).bold = True
            paragraph.runs[0].font.size = Pt(self._settings.recommendation_font_size_pt)
            paragraph.runs[0].font.color.rgb = RGBColor(*final_priority_colors[recommendation.priority])

        priority_icons = {
            RecommendationPriority.HIGH: "🔴",
            RecommendationPriority.MEDIUM: "🟡",
            RecommendationPriority.LOW: "🟢",
        }
        other_recommendations = [
            recommendation
            for recommendation in recommendations
            if recommendation.priority not in [
                RecommendationPriority.CRITICAL,
                RecommendationPriority.WARNING,
                RecommendationPriority.APPROVED,
            ]
        ]
        if other_recommendations:
            doc.add_paragraph("Recomendaciones específicas:").bold = True
            for recommendation in other_recommendations:
                icon = priority_icons.get(recommendation.priority, "⚪")
                doc.add_paragraph(f"{icon} {recommendation.message}", style="List Bullet")
```

### B. Update `ReportFixtures`

We will modify [fixtures.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/adapters/report/fixtures.py) to set a default list of `RecommendationDTO` instances instead of raw dictionaries for `recommendations`.

---

## 7. Testing Strategy

### A. Domain Tests: `RecommendationBuilder`

File: `src/domain/tests/recommendation/test_recommendation_builder.py`
We will write a complete suite test utilizing `unittest.TestCase` checking:
- Standard case where overall quality and grammar scores are above 7.0 (should produce APPROVED publication recommendation).
- Low quality score (< 7.0) (should yield HIGH priority quality recommendation + WARNING publication recommendation).
- Critical quality score (< 5.0) (should yield CRITICAL publication recommendation).
- Low grammar score (< 7.0) (should yield HIGH priority grammar recommendation + WARNING publication recommendation).
- Critical grammar score (< 5.0) (should yield CRITICAL publication recommendation).
- Quality dimensions score (< 6.0) (should yield MEDIUM priority dimension-specific recommendations).
- Non-valid structure (should yield HIGH priority missing-section recommendations + CRITICAL publication recommendation).
- Low citation match rate (< 90%) (should yield HIGH priority citation matching recommendation + WARNING publication recommendation).
- Critical citation match rate (< 50%) (should yield CRITICAL publication recommendation).
- Low citation count (< 10) (should yield MEDIUM priority theoretical frame recommendation).
- Low classification confidence (< 0.7) (should yield LOW priority classification check recommendation).

### B. Application Tests: `AnalyzeDocumentUseCase`

File: `src/application/tests/test_analyze_document_use_case.py`
We will mock or construct double implementations of the sub-use cases using `unittest.mock.MagicMock` to isolate the orchestration:
- GIVEN valid paragraphs, citations, structure, and classification inputs.
- WHEN `execute()` is called.
- THEN verify that all sub-use case methods (`execute()`) are invoked with expected args.
- THEN check that the returned object is a `ReportInputDTO` containing expected aggregations.

### C. Wiring Tests: `AnalyzeDocumentUseCaseWiring`

File: `src/infrastructure/tests/test_analyze_document_use_case_wiring.py`
We will write tests to verify wiring:
- GIVEN `AnalyzeDocumentUseCaseWiring().create_use_case()` is called.
- THEN verify that the returned instance is of type `AnalyzeDocumentUseCase`.
- THEN check that all underlying use case properties are set and match their expected types.

---

## 8. Risks and Mitigations

1. **Sub-use Case Coupling**: Any change in sub-use cases could break the orchestrator.
   - *Mitigation*: We rely strictly on standard stable DTO contracts (e.g. `QualityResultDTO`, `GrammarCheckResultDTO`) for all inputs, isolating the orchestrator from implementation detail changes.
2. **Circular Dependencies / Wrong Imports**: `clean-architecture` guidelines mandate strict layer boundaries.
   - *Mitigation*:
     - `src/domain/` must only import standard modules, enums, exceptions, and DTOs. It must never import anything from `src/application/` or `src/infrastructure/`.
     - `src/application/` must only import from `src/domain/` or standard modules.
     - `src/infrastructure/` can import from any layer.
3. **Abbreviations**: The project enforces a strict no-abbreviation naming policy.
   - *Mitigation*: We will perform self-checks to ensure variables like `rec`, `dim`, `doc`, `exc`, `uc` are completely expanded in all newly written or updated code.

---

## 9. File Inventory

The following files are part of Slice 13 scope:

### New Files:
- `src/domain/dtos/recommendation_dto.py`
- `src/domain/recommendation/recommendation_builder.py`
- `src/application/analyze_document_use_case.py`
- `src/infrastructure/wirings/analyze_document_use_case_wiring.py`
- `src/domain/tests/recommendation/test_recommendation_builder.py`
- `src/application/tests/test_analyze_document_use_case.py`
- `src/infrastructure/tests/test_analyze_document_use_case_wiring.py`

### Modified Files:
- `src/domain/enums/recommendation_priority.py` (add critical, warning, approved enums)
- `src/domain/tests/enums/test_recommendation_priority.py` (update member tests)
- `src/domain/dtos/report_input_dto.py` (update type of `recommendations` list)
- `src/infrastructure/adapters/report/docx_report_adapter.py` (refactor to use `RecommendationDTO` properties)
- `src/infrastructure/tests/adapters/report/fixtures.py` (refactor dummy data to use `RecommendationDTO` instances)
