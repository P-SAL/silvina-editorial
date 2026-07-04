# Proposal: Refactor Gradio Web Controller (Slice 15)

## Intent

Remove inter-controller coupling by refactoring [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) to directly use clean hexagonal architecture use cases ([AnalyzeDocumentUseCase](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) and [ExportReportUseCase](file:///E:/Python/silvina-editorial/src/application/export_report_use_case.py)) via their wirings, instead of relying on the legacy [main.py](file:///E:/Python/silvina-editorial/main.py) shim. We will bind the UI display to DTO properties and save the JSON report in the new [AnalysisResultDTO](file:///E:/Python/silvina-editorial/src/domain/dtos/analysis_result_dto.py) structure serialized directly.

## Scope

### In Scope
- Remove the import of `SilvinaEditorialAssistant` from [main.py](file:///E:/Python/silvina-editorial/main.py).
- Instantiate use cases via [AnalyzeDocumentUseCaseWiring](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py) and [ExportReportWiring](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/export_report_wiring.py).
- Refactor `create_results_display` to read from [ReportInputDTO](file:///E:/Python/silvina-editorial/src/domain/dtos/report_input_dto.py) instead of the legacy dictionary format, cleaning up internal HTML/CSS string construction while preserving exact visual parity.
- Extract publication status and verdict messages from `report.verdict` properties.
- Serialize the JSON report directly from [AnalysisResultDTO](file:///E:/Python/silvina-editorial/src/domain/dtos/analysis_result_dto.py) using `.as_dict()` (and a local Enum/datetime serializer helper) instead of matching the legacy dictionary structure.
- Catch [BaseSrcError](file:///E:/Python/silvina-editorial/src/domain/exceptions/base_src_error.py) to display Spanish messages in the UI without tracebacks.

### Out of Scope
- Modifying core domain models, entities, or use cases.
- Changing the styling, layout, or CSS of the Gradio interface.
- Refactoring the expert feedback save mechanism.

## Capabilities

### New Capabilities
- None

### Modified Capabilities
- None

## Approach

- Import `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring`.
- Use `AnalyzeDocumentUseCase` to perform analysis, returning `ReportInputDTO`.
- Pass `ReportInputDTO` to `ExportReportUseCase` to generate the Word report.
- Instantiate `AnalysisResultDTO` from `ReportInputDTO` fields, then serialize it using `.as_dict()` processed by a recursive converter helper to clean up Enums and datetimes for writing the JSON report.
- Clean up html-building in `create_results_display` using `ReportInputDTO` properties.
- Catch `BaseSrcError` and other unexpected exceptions in the processing loop.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) | Modified | Replace legacy controller shim with wirings, bind UI to DTOs, direct JSON serialization, and add clean error handling. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| JSON serialization of Enums/Datetimes fails | Medium | Implement recursive parser to extract Enum values and ISO-format datetimes. |
| Visual layout regression in HTML render | Low | Rigorously check HTML output structure and inline CSS properties to guarantee visual parity. |

## Rollback Plan

Run `git checkout -- gradio_app.py` to revert all edits to the original on-disk state.

## Dependencies

- Completion of Slice 14, and availability of `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring`.

## Success Criteria

- [ ] [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) starts and runs without importing `SilvinaEditorialAssistant`.
- [ ] Analysis completes and displays results with exact visual parity to the legacy interface.
- [ ] Word report is generated successfully via `ExportReportUseCase`.
- [ ] JSON report is saved containing the direct serialization of `AnalysisResultDTO`.
- [ ] Domain exceptions are caught and reported as user-friendly Spanish errors.
