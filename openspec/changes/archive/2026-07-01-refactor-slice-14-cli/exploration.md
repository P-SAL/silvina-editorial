# Exploration Report: Refactoring main.py CLI Controller (Slice 14)

This report details the exploration for refactoring the main CLI controller ([main.py](file:///E:/Python/silvina-editorial/main.py)) under Slice 14 of the Silvina Editorial Assistant migration. The goal is to fully transition the CLI entry point to use the Hexagonal Architecture UseCases and Wirings, clean up obsolete dependencies, establish structured exception handling with exit codes, and ensure backwards compatibility with legacy consumers (the Gradio interface and E2E tests).

---

## 1. Current State of main.py

The current implementation of [main.py](file:///E:/Python/silvina-editorial/main.py) is a monolith that coordinates the complete analysis pipeline. It acts as both the entry point and the orchestrator:

*   **Manual Orchestration:** The `SilvinaEditorialAssistant` class manually instantiates and chains legacy components:
    *   `WordReader`, `ContentExtractor`, `CitationParser`, and `ReferenceParser` for data access.
    *   `ArticleClassifier`, `QualityAnalyzer`, `CitationMatcher`, and `StructureValidator` for business logic.
    *   `ReportFormatter` and `WordExporter` for report generation.
*   **Ad-hoc Logic:** It contains hardcoded thresholds, direct dependencies on `python-docx` (`docx.Document`), and inline formatting methods (e.g. `_generate_recommendations`, `_format_quality_level`, `_format_category`, `_prepare_for_json`).
*   **Basic Error Handling:** Exceptions are caught globally using a catch-all `except Exception as e` block, which prints raw tracebacks and exits with status code `1` (or `0` on keyboard interrupt).
*   **Consumer Dependencies:**
    *   [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) imports `SilvinaEditorialAssistant` and calls `analyze_document()`, `save_word_report()`, and `save_json_report()`, expecting a specific legacy dictionary structure.
    *   [test_cli_e2e.py](file:///E:/Python/silvina-editorial/tests/e2e/test_cli_e2e.py) performs in-process E2E test runs validating key assertions on the returned dictionary keys like `filename`, `document_info`, `classification`, `quality_analysis`, and `structure_validation`.

---

## 2. Decoupled Architecture Proposal

To decouple [main.py](file:///E:/Python/silvina-editorial/main.py) and align it with Hexagonal Architecture:

1.  **Use Cases and Wirings:**
    *   We will instantiate [AnalyzeDocumentUseCase](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) using [AnalyzeDocumentUseCaseWiring](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py).
    *   We will instantiate [ExportReportUseCase](file:///E:/Python/silvina-editorial/src/application/export_report_use_case.py) using [ExportReportWiring](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/export_report_wiring.py).
2.  **Maintain `SilvinaEditorialAssistant` for Compatibility:**
    *   Instead of deleting `SilvinaEditorialAssistant`, we will refactor it into a thin shim that delegates directly to the Use Cases.
    *   This preserves the interface required by [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) and [test_cli_e2e.py](file:///E:/Python/silvina-editorial/tests/e2e/test_cli_e2e.py).
    *   It will execute the decoupled pipeline and convert the returned `ReportInputDTO` into a legacy-compatible dictionary structure using a private helper method.

### Proposed DTO to Legacy Dictionary Mapping

```python
def _convert_dto_to_legacy_dict(self, dto: ReportInputDTO) -> dict[str, Any]:
    # Formats enum values and constructs equivalent structures for legacy consumers.
    legacy_recommendations = [
        {"priority": rec.priority.value, "message": rec.message}
        for rec in dto.recommendations
    ]
    # Append the final publication verdict/recommendation as the last item
    legacy_recommendations.append({
        "priority": dto.verdict.verdict.value,
        "message": dto.verdict.message
    })

    return {
        "filename": Path(dto.filename).name,
        "citations": [],  # Kept for interface backward-compatibility (not accessed by consumers)
        "document_info": {
            "title": dto.document_content.title,
            "authors": dto.document_content.authors,
            "word_count": dto.document_content.word_count,
            "char_count": dto.document_content.char_count,
            "estimated_pages": dto.document_content.word_count // 250,
        },
        "classification": {
            "category": dto.classification.article_type,
            "article_size": dto.classification.article_size,
            "confidence": dto.classification.confidence,
            "reasoning": dto.classification.reasoning,
        },
        "quality_analysis": {
            "overall_score": dto.quality.overall_score,
            "quality_level": dto.quality.quality_level,
            "gramatica": {
                "score": dto.grammar.score,
                "feedback": dto.grammar.feedback,
                "errors": [
                    {
                        "message": err.message,
                        "context": err.context,
                        "replacements": err.replacements,
                    }
                    for err in dto.grammar.errors
                ],
            },
            "dimensions": dto.quality.dimension_scores,
        },
        "structure_validation": {
            "is_valid": dto.structure.is_valid,
            "missing_sections": dto.structure.missing_sections,
            "details": dto.structure.section_details,
        },
        "citations_analysis": {
            "total_citations": dto.citations.total_citations,
            "total_references": dto.citations.total_references,
            "matched_count": dto.citations.matched_count,
            "unmatched_count": dto.citations.unmatched_count,
            "by_type": dto.citations.citations_by_type,
            "unmatched_citations": dto.citations.unmatched_citations,
            "apa_violations": len(dto.apa_validation.violations),
            "apa_compliant": len(dto.apa_validation.violations) == 0,
        },
        "apa_validation": {
            "violations": [
                {
                    "citation": v.citation_text,
                    "error_type": v.error_type.value,
                    "location": v.location,
                    "explanation": v.explanation,
                    "correction": v.correction,
                }
                for v in dto.apa_validation.violations
            ],
            "report": dto.apa_validation.report,
        },
        "recommendations": legacy_recommendations,
    }
```

---

## 3. Exception Handling and Exit Codes Design

Under Hexagonal Architecture, errors raised by the domain or adapters inherit from [BaseSrcError](file:///E:/Python/silvina-editorial/src/domain/exceptions/base_src_error.py). The CLI controller must catch these exceptions and map them cleanly without displaying raw Python tracebacks.

### Exception Types and Mapped Exit Codes

| Exception Type | Trigger Scenario | Exit Code | Console Message (Spanish) |
|---|---|---|---|
| `KeyboardInterrupt` | User interrupts command execution via Ctrl+C | `0` | `\n\n⚠️ Análisis interrumpido por el usuario` |
| `SrcBaseNotFound` | Input file not found, or missing critical sections / entities | `1` | `❌ Error: {error_message}` (Clean message, no traceback) |
| `BaseSrcError` | Other domain errors (parsing, classification, quality errors) | `1` | `❌ Error de análisis: {error_message}` (Clean message, no traceback) |
| `ValueError` / `TypeError` | Invalid inputs or arguments passed to CLI | `2` | `❌ Argumento o valor inválido: {error}` (Clean error) |
| General `Exception` | Unexpected bug / system crash | `1` | `❌ Error fatal inesperado: {error}` (Prints traceback to stderr for debugging) |

> [!IMPORTANT]
> To comply with architecture constraints, any catch block matching `BaseSrcError` will fetch the error message from the dictionary format returned by the `dict()` method (e.g. `err.dict()["error"]`) and output it cleanly to the console.

---

## 4. Code Cleanup and Simplification

Refactoring [main.py](file:///E:/Python/silvina-editorial/main.py) allows us to remove numerous obsolete parts:

1.  **Remove Legacy Imports:**
    *   Remove imports of `WordReader`, `ContentExtractor`, `CitationParser`, and `ReferenceParser`.
    *   Remove imports of `ArticleClassifier`, `QualityAnalyzer`, `CitationMatcher`, and `StructureValidator`.
    *   Remove imports of `ReportFormatter` and `WordExporter`.
    *   Remove the duplicate/fallback sys-path manipulations and imports of `verify_eumic_compliance` / `validate_apa_citations`.
2.  **Remove Obsolete Methods:**
    *   `SilvinaEditorialAssistant._generate_recommendations()` (receded to `RecommendationBuilder` in domain).
    *   `SilvinaEditorialAssistant._format_category()` (handled by translation rules or serialization).
    *   `SilvinaEditorialAssistant._format_quality_level()` (handled by string processing in presentation).
3.  **JSON Serialization Helper:**
    *   Refactor `_prepare_for_json` to support mapping enums generically by checking `isinstance(data, Enum): return data.value` instead of checking individual enums.
4.  **CLI Report Generation:**
    *   Refactor the `main()` function to directly consume the result from `SilvinaEditorialAssistant().analyze_document()` or output them in a structured way utilizing the same clean representation.

---

## 5. Risk Assessment

*   **Risk:** Changing the `SilvinaEditorialAssistant` return format could break `gradio_app.py` or `test_cli_e2e.py`.
    *   *Mitigation:* The `_convert_dto_to_legacy_dict` mapping method will be thoroughly validated to ensure it exposes exactly the same key/value types, structure, and enums as the legacy controller.
*   **Risk:** Test doubles or mock dependencies inside tests might conflict with the new wiring.
    *   *Mitigation:* Verification tests (TDD) will be run in the next phase to ensure the mocks (like `ollama.Client`) function correctly with the refactored code.
