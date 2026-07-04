# Exploration Report: Refactoring gradio_app.py Web Controller (Slice 15)

This report details the exploration for refactoring the Gradio web interface controller ([gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py)) under Slice 15 of the Silvina Editorial Assistant migration. The goal is to transition the Gradio driving adapter to consume the Clean Hexagonal Architecture use cases and wirings directly, rather than relying on the legacy shim provided by the CLI controller ([main.py](file:///E:/Python/silvina-editorial/main.py)).

---

## 1. Current State and Coupling of gradio_app.py

Currently, [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) acts as a monolithic Gradio web interface. It depends on the CLI orchestrator shim to invoke domain logic:

*   **Inter-Controller Coupling:**
    *   It imports `SilvinaEditorialAssistant` from [main.py](file:///E:/Python/silvina-editorial/main.py) (line 21).
    *   This violates clean architecture principles: driving adapters (Gradio UI and CLI) should be decoupled from one another and independently consume application services (use cases).
*   **Pipeline Execution:**
    *   In `process_document(uploaded_file)` (lines 36–90), it instantiates `silvina = SilvinaEditorialAssistant()` and calls `results = silvina.analyze_document(uploaded_file.name)`.
    *   It generates Word reports via `silvina.save_word_report(results, ...)` and JSON reports via `silvina.save_json_report(results, ...)`.
*   **Legacy Dict Consumption:**
    *   `create_results_display(results)` (lines 95–237) manually unpacks a legacy dictionary structure: `results['document_info']`, `results['classification']`, `results['quality_analysis']`, etc.
    *   It reads the final verdict by accessing the last item in the recommendations list (`recommendations[-1]`), expecting a dict with `priority` and `message` keys.
*   **Errors:**
    *   Domain exceptions from the underlying layers are not caught specifically; a generic `except Exception as e` handles failures, which prints tracebacks and returns a generic error string.

---

## 2. Decoupled Architecture Proposal

To decouple [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py), we will migrate it to consume application layer use cases directly via their infrastructure wirings, operating on strongly typed DTOs:

1.  **Direct Use Cases Instantiation:**
    *   We will import [AnalyzeDocumentUseCaseWiring](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py) and [ExportReportWiring](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/export_report_wiring.py).
    *   At the module level or inside controllers, we will instantiate the use cases:
        ```python
        analyze_document_use_case = AnalyzeDocumentUseCaseWiring().create_use_case()
        export_report_use_case = ExportReportWiring().create_use_case()
        ```
2.  **DTO Property Access:**
    *   `create_results_display` will receive a [ReportInputDTO](file:///E:/Python/silvina-editorial/src/domain/dtos/report_input_dto.py) instead of a dictionary.
    *   We will refactor lookups to use properties of the DTOs instead of nested dictionary keys:
        *   `results['document_info']` $\rightarrow$ `report.document_content` (type `DocumentContentDTO`)
        *   `results['classification']` $\rightarrow$ `report.classification` (type `ClassificationResultDTO`)
        *   `results['quality_analysis']` $\rightarrow$ `report.quality` (type `QualityResultDTO`)
        *   `results['structure_validation']` $\rightarrow$ `report.structure` (type `StructureValidationResultDTO`)
        *   `results['citations_analysis']` $\rightarrow$ `report.citations` (type `CitationAnalysisResultDTO`)
        *   `results['recommendations']` $\rightarrow$ `report.recommendations` (type `list[RecommendationDTO]`)
    *   Instead of extracting the final verdict from `recommendations[-1]`, we will directly consume `report.verdict` (type `PublicationVerdictDTO`), checking `report.verdict.verdict` against [PublicationVerdict](file:///E:/Python/silvina-editorial/src/domain/enums/publication_verdict.py) enum values.

### Proposed DTO to Gradio HTML Property Mapping

| UI Section | Legacy Dictionary Access | Refactored DTO Property Access |
|---|---|---|
| **Document Title** | `doc_info.get('title', 'Sin título')` | `report.document_content.title or "Sin título"` |
| **Authors** | `doc_info.get('authors', 'No especificado')` | `report.document_content.authors or "No especificado"` |
| **Word Count** | `doc_info['word_count']` | `report.document_content.word_count` |
| **Article Type** | `classification['category'].value.upper()` | `report.classification.article_type.value.upper()` |
| **Final Status** | `final_rec['priority']` | `report.verdict.verdict` (check against `PublicationVerdict.APPROVED` etc.) |
| **Final Message** | `final_rec['message']` | `report.verdict.message` |
| **Grammar Score** | `quality['gramatica']['score']` | `report.grammar.score` |
| **Grammar Feedback** | `quality['gramatica']['feedback']` | `report.grammar.feedback` |
| **Semantic Score** | `quality['overall_score']` | `report.quality.overall_score` |
| **Quality Level** | `quality['quality_level'].value` | `report.quality.quality_level.value` |
| **Grammar Errors** | `len(quality['gramatica'].get('errors', []))` | `len(report.grammar.errors)` |
| **APA Errors** | `citations.get('apa_violations', 0)` | `report.apa_validation.violation_count` |
| **Unmatched Citations** | `citations.get('unmatched_count', 0)` | `report.citations.unmatched_count` |
| **Missing Sections** | `len(structure['missing_sections'])` | `len(report.structure.missing_sections)` |
| **Semantic Dimensions** | `quality['dimensions'].items()` | `report.quality.dimension_scores.items()` |
| **Critical Issues** | `rec['priority'] in ['alta', 'critica']` | `rec.priority == RecommendationPriority.HIGH` (from `RecommendationPriority`) |

---

## 3. Exception Handling and UI Resiliency

To make the Gradio app resilient and prevent unhandled server failures:

1.  **Domain Exception Catching:**
    *   We will specifically catch [BaseSrcError](file:///E:/Python/silvina-editorial/src/domain/exceptions/base_src_error.py).
    *   If a domain error occurs, we will extract the clean error message using `exc.dict().get("error", str(exc))` and display it as the status message in the Gradio textbox, preventing traceback printouts to the user.
2.  **Unknown Errors:**
    *   Unexpected exceptions will be caught, printed to standard error for developer debugging, and returned to the UI with a generic error message (e.g. `❌ Error al procesar el documento: {str(e)}`).

---

## 4. Code Cleanup

During refactoring, the following cleanup will be performed:

*   **Remove Legacy Imports:** The import `from main import SilvinaEditorialAssistant` will be removed.
*   **Specific Imports:** Replace any wildcard or module-level imports with specific name imports as required by the `clean-architecture` skill conventions (e.g. import `PublicationVerdict` from `src.domain.enums.publication_verdict` and `RecommendationPriority` from `src.domain.enums.recommendation_priority`).
*   **JSON Serialization:** Keep a localized `_prepare_for_json` utility function within `gradio_app.py` to recursively map DTO dictionaries (from `report.as_dict()`) and translate Enum fields to their values.

---

## 5. Risk Assessment

*   **JSON Serialization compatibility:** DTO `as_dict()` returns recursively nested dataclasses as dictionaries, but retains Enum members. The local JSON serializer must correctly extract `.value` from Enums to prevent crash during `json.dump()`.
*   **Visual Regression:** The HTML returned by `create_results_display` must render exactly the same tags, classes, and styles. All variables must be correctly mapped to prevent layout breaking.
