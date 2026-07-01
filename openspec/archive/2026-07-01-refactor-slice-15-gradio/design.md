# Technical Design: Refactor Gradio Web Controller (Slice 15)

## 1. Technical Approach

This design describes the migration of the Gradio web controller ([gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py)) to directly consume clean hexagonal architecture use cases, decoupling it from the CLI driving adapter ([main.py](file:///E:/Python/silvina-editorial/main.py)).

Instead of importing `SilvinaEditorialAssistant` from `main.py`, the controller will:
1. Instantiate the application layer use cases at startup using their respective infrastructure wirings.
2. Coordinate the analysis execution and output export using strongly typed Data Transfer Objects (DTOs).
3. Bind UI rendering to the fields of [ReportInputDTO](file:///E:/Python/silvina-editorial/src/domain/dtos/report_input_dto.py).
4. Save the JSON report by directly serializing [AnalysisResultDTO](file:///E:/Python/silvina-editorial/src/domain/dtos/analysis_result_dto.py) with a localized serializer helper.
5. Capture domain exceptions (`BaseSrcError`) and present clean, localized messages in the UI.

---

## 2. Architecture Decisions

### Decision 1: Startup Instantiation of Use Cases via Wirings
* **Choice**: Import and instantiate `AnalyzeDocumentUseCase` and `ExportReportUseCase` once at module scope during startup.
  ```python
  from src.infrastructure.wirings.analyze_document_use_case_wiring import AnalyzeDocumentUseCaseWiring
  from src.infrastructure.wirings.export_report_wiring import ExportReportWiring

  analyze_document_use_case = AnalyzeDocumentUseCaseWiring().create_use_case()
  export_report_use_case = ExportReportWiring().create_use_case()
  ```
* **Alternatives Considered**: Instantiating wirings dynamically inside `process_document` on every request.
* **Rationale**: Module-scope instantiation acts as a fast-fail check, ensuring wiring issues are caught immediately when the server starts or during import in automated test environments.

### Decision 2: Direct DTO Binding in UI Display
* **Choice**: Refactor `create_results_display` to receive a `ReportInputDTO` instance instead of the legacy dictionary format.
* **Alternatives Considered**: Maintaining the legacy dict mapping in the controller and converting the DTO back to a dictionary.
* **Rationale**: Eliminating dictionary mappings simplifies the codebase, improves performance, and enables editor static-type verification for UI bindings.

### Decision 3: Decoupling Recommendations from Final Verdict
* **Choice**: Unpack final publication status and message directly from `report.verdict` (type `PublicationVerdictDTO`) and filter critical recommendations using `rec.priority == RecommendationPriority.HIGH` over the entire `report.recommendations` list.
* **Alternatives Considered**: Continuing to append the final verdict as the last item in the recommendations list (`recommendations[-1]`).
* **Rationale**: The domain layer now explicitly distinguishes between editorial recommendations (specific improvements) and the overall verdict (system decision). Handling these fields separately reflects the clean architecture boundaries.

### Decision 4: Localized Serialization for JSON Report
* **Choice**: Map `ReportInputDTO` to `AnalysisResultDTO` fields and serialize it with a recursive serializer helper `_prepare_for_json` that converts `Enum` values to strings and `datetime` objects to ISO strings.
* **Alternatives Considered**: Using standard `json.dumps()` with a custom JSON encoder class.
* **Rationale**: A localized helper function is lightweight, does not require configuring global encoders, and matches the formatting conventions established in the codebase.

### Decision 5: Domain Exception Mapping
* **Choice**: Explicitly catch `BaseSrcError` and retrieve its message via `exc.dict().get("error", str(exc))` to update the UI status. Unexpected exceptions are caught, logged to standard error, and returned as a generic error string.
* **Alternatives Considered**: Letting exceptions propagate raw to the UI.
* **Rationale**: Displaying technical stack traces to non-technical editorial staff degrades user experience. Intercepting domain exceptions allows presenting clear, Spanish-language error messages while keeping debugging details in the server logs.

---

## 3. Data Flow

```
[ gradio_app.py (UI Upload) ]
             │
             ▼
[ process_document(uploaded_file) ]
             │
             ▼
[ AnalyzeDocumentUseCase.execute() ] ── (via AnalyzeDocumentUseCaseWiring)
             │
             ▼
      [ ReportInputDTO ]
       ┌─────┼──────────────────────────────────┐
       │     │                                  │
       ▼     ▼                                  ▼
   [ create_results_display() ]    [ ExportReportUseCase.execute() ]  [ AnalysisResultDTO ]
       │                                        │                              │
       ▼                                        ▼                              ▼
 [ HTML Rendering ]                      [ .docx Report ]            [ _prepare_for_json() ]
                                                                               │
                                                                               ▼
                                                                        [ .json Report ]
```

---

## 4. UI Properties and Field Bindings

The table below describes the direct mappings from the `ReportInputDTO` (and its nested DTO properties) to the Gradio HTML results panel:

| HTML Display Property | Target DTO Field / Expression | Source Type | Description |
|:---|:---|:---|:---|
| **Document Title** | `report.document_content.title or "Sin título"` | `str` | Title extracted from document metadata. |
| **Authors** | `report.document_content.authors or "No especificado"` | `str` | Document author metadata. |
| **Word Count** | `report.document_content.word_count` | `int` | Total number of words analyzed. |
| **Article Type** | `report.classification.article_type.value.upper()` | `ArticleType` (Enum) | Classified article category (e.g. *CIENTIFICO*). |
| **Verdict Status** | `report.verdict.verdict` | `PublicationVerdict` (Enum) | Evaluated status: *aprobado*, *advertencia*, *critica*. |
| **Verdict Message** | `report.verdict.message` | `str` | Final verdict explanation. |
| **Grammar Score** | `report.grammar.score` | `float` | Quality score for spelling and grammar. |
| **Grammar Feedback** | `report.grammar.feedback` | `str` | Textual grammar summary feedback. |
| **Semantic Score** | `report.quality.overall_score` | `float` | Aggregated semantic quality score. |
| **Quality Level** | `report.quality.quality_level.value` | `QualityLevel` (Enum) | Quality tier (e.g. *excelente*). |
| **Grammar Errors** | `len(report.grammar.errors)` | `int` | Total number of detected grammatical violations. |
| **APA Errors** | `report.apa_validation.violation_count` | `int` | Count of formatting violations. |
| **Unmatched Citations** | `report.citations.unmatched_count` | `int` | Citations referencing missing bibliography items. |
| **Missing Sections** | `len(report.structure.missing_sections)` | `int` | Count of missing template sections. |
| **Semantic Dimensions** | `report.quality.dimension_scores.items()` | `dict[str, dict[str, Any]]` | Loop key/value for dimension scores (e.g., *coherencia*). |
| **Critical Issues** | `rec.priority == RecommendationPriority.HIGH` | `RecommendationPriority` | Filter for critical issues display. |

---

## 5. Localized JSON Serialization Helper

To convert nested DTOs (using `.as_dict()`), `Enum` values, and `datetime` types into a pure dictionary suitable for `json.dump`, the following recursive helper will be added to `gradio_app.py`:

```python
from datetime import datetime
from enum import Enum
from typing import Any
from src.domain.dtos.base_dto import BaseDTO

def _prepare_for_json(data: Any) -> Any:
    """Recursively convert enums, datetimes, and DTOs into JSON-serializable structures."""
    if isinstance(data, dict):
        return {k: _prepare_for_json(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [_prepare_for_json(item) for item in data]
    elif isinstance(data, Enum):
        return data.value
    elif isinstance(data, datetime):
        return data.isoformat()
    elif isinstance(data, BaseDTO):
        return _prepare_for_json(data.as_dict())
    else:
        return data
```

---

## 6. Exception Handling Blocks

The execution block inside `process_document` will catch domain exceptions separately:

```python
from src.domain.exceptions.base_src_error import BaseSrcError
import traceback

def process_document(uploaded_file):
    # ... check file ...
    try:
        # 1. Run Analysis Use Case
        report = analyze_document_use_case.execute(uploaded_file.name)

        # 2. Setup Report Paths
        base_name = Path(uploaded_file.name).stem
        output_dir = Path.home() / "Documents" / "Silvina" / "reports"
        output_dir.mkdir(parents=True, exist_ok=True)
        word_report_path = output_dir / f"{base_name}_analisis.docx"
        json_report_path = output_dir / f"{base_name}_analisis.json"

        # 3. Export Word Report via ExportReportUseCase
        export_report_use_case.execute(
            report_input=report,
            output_path=str(word_report_path)
        )

        # 4. Instantiate AnalysisResultDTO & Save JSON
        analysis_result = AnalysisResultDTO(
            filename=report.filename,
            document_content=report.document_content,
            classification=report.classification,
            quality=report.quality,
            structure=report.structure,
            citations=report.citations
        )
        json_data = _prepare_for_json(analysis_result)
        with open(json_report_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)

        # 5. Build HTML UI Display
        results_html = create_results_display(report)
        success_msg = "✅ Análisis completado exitosamente"

        return (
            success_msg,
            results_html,
            str(word_report_path),
            str(json_report_path),
            str(word_report_path),
            gr.Button(interactive=True)
        )

    except BaseSrcError as exc:
        # Extract clean domain message to display in UI without traceback
        error_msg = f"❌ Error de validación: {exc.dict().get('error', str(exc))}"
        print(f"\n[Domain Error] {error_msg}")
        return (error_msg, "", None, None, "", gr.Button(interactive=True))

    except Exception as e:
        # Log generic unexpected runtime errors
        error_msg = f"❌ Error al procesar el documento: {str(e)}"
        print(f"\n[System Error] {error_msg}")
        traceback.print_exc()
        return (error_msg, "", None, None, "", gr.Button(interactive=True))
```

---

## 7. Testing Strategy

| Layer / Level | Target component | Verification / Assertion method |
|:---|:---|:---|
| **Compilation** | Gradio Blocks Interface | Assert that `gradio_app.py` exposes a module-level `demo` variable (the `gr.Blocks` instance) without starting the web server. |
| **Unit** | UI HTML Generation | Write unit tests calling `create_results_display(mock_report_input_dto)` and check that the resulting string contains the correct HTML elements, classes, and styles. |
| **Unit** | Serializer Helper | Test `_prepare_for_json` using nested dicts containing custom `Enum`, `datetime`, and dummy DTO instances, asserting that they serialize cleanly. |
| **E2E Integration** | E2E Gradio Pipeline | Execute [test_gradio_e2e.py](file:///E:/Python/silvina-editorial/tests/e2e/test_gradio_e2e.py) to run mock analysis files through the test client and verify results. |
