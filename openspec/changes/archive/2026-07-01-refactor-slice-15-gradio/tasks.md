# Tasks: Refactor Gradio Web Controller (Slice 15)

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~150-250 lines |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Refactor gradio_app.py to use use cases and DTOs | PR 1 | Base branch; E2E and unit tests included |

## Phase 1: Imports & Startup Wiring Instantiation

- [x] 1.1 Remove import of `SilvinaEditorialAssistant` from `main.py` in [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py).
- [x] 1.2 Import `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring` in [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py).
- [x] 1.3 Instantiate `analyze_document_use_case` and `export_report_use_case` at module scope using their respective wirings.

## Phase 2: DTO Bindings & UI HTML Generation Refactoring

- [x] 2.1 Refactor `create_results_display` signature and body to accept `ReportInputDTO` instead of a dictionary.
- [x] 2.2 Bind HTML fields for document title, author, and word count to properties of `report.document_content`.
- [x] 2.3 Bind classification to `report.classification` and display verdict status and message from `report.verdict`.
- [x] 2.4 Bind grammar, quality score, grammar errors, APA violations, and unmatched citations to their respective nested DTO fields.
- [x] 2.5 Filter critical recommendations using `rec.priority == RecommendationPriority.HIGH` over the `report.recommendations` list.

## Phase 3: Serialization Helper & JSON/Word Output

- [x] 3.1 Implement recursive helper `_prepare_for_json` to handle DTOs, Enums, and datetimes.
- [x] 3.2 Update `process_document` to run `analyze_document_use_case.execute` and retrieve `ReportInputDTO`.
- [x] 3.3 Invoke `export_report_use_case.execute` inside `process_document` to write the Word report.
- [x] 3.4 Instantiate `AnalysisResultDTO` from `ReportInputDTO` fields, convert it using `_prepare_for_json`, and write the JSON report.

## Phase 4: Exception Handling & UI Resiliency

- [x] 4.1 Catch `BaseSrcError` in `process_document`, log it, and return a user-friendly Spanish message to the UI status.
- [x] 4.2 Catch generic `Exception` in `process_document`, print traceback, and return a clean system error message in Spanish.

## Phase 5: Verification & Tests

- [x] 5.1 Run [test_gradio_e2e.py](file:///E:/Python/silvina-editorial/tests/e2e/test_gradio_e2e.py) to verify the server loads the `demo` block successfully.
- [x] 5.2 Add unit tests for `_prepare_for_json` and `create_results_display` in [test_gradio_e2e.py](file:///E:/Python/silvina-editorial/tests/e2e/test_gradio_e2e.py) using `unittest.TestCase`.
