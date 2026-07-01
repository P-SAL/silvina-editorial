# Report Export — Hexagonal Layer Specification
## Slice 12 (export-report)

## Purpose

Defines the port, DTO, exception, adapter, use case, wiring, and fake double that expose `.docx` report export as a testable hexagonal service, while leaving `presentation/word_exporter.py` untouched.

---

## Requirements

### Requirement: ReportExportPort Abstract Interface

`ReportExportPort(ABC)` MUST reside in `src/domain/report/` and declare one abstract method: `export(report_input: ReportInputDTO, path: str) -> bool`. The domain MUST NOT import from infrastructure or application layers.

#### Scenario: Concrete implementation satisfies the contract

- GIVEN a class implementing `export(report_input, path)`
- WHEN it is instantiated and passed as a `ReportExportPort`
- THEN no `TypeError` is raised

#### Scenario: Abstract class cannot be instantiated directly

- GIVEN `ReportExportPort` is abstract
- WHEN client code calls `ReportExportPort()`
- THEN Python raises `TypeError`

---

### Requirement: ReportInputDTO Is Frozen

`ReportInputDTO` MUST be a frozen dataclass in `src/domain/dtos/report_input_dto.py` with fields:
- `filename: str`
- `document_content: DocumentContentDTO`
- `classification: ClassificationResultDTO`
- `quality: QualityResultDTO`
- `grammar: GrammarCheckResultDTO`
- `structure: StructureValidationResultDTO`
- `citations: CitationAnalysisResultDTO`
- `apa_validation: ApaValidationResultDTO`
- `recommendations: list[RecommendationDTO]`
- `verdict: PublicationVerdictDTO`
- `eumic_violations: list[EumicViolationDTO]`

> Updated in Slice 13: `recommendations` is now `list[RecommendationDTO]` (was `list[dict]`). `verdict: PublicationVerdictDTO` and `eumic_violations: list[EumicViolationDTO]` were added.

#### Scenario: DTO is immutable after construction

- GIVEN a constructed `ReportInputDTO`
- WHEN any field is reassigned
- THEN Python raises `FrozenInstanceError`

#### Scenario: DTO holds typed verdict and recommendation lists

- GIVEN a valid instance of `ReportInputDTO`
- WHEN `recommendations`, `verdict`, and `eumic_violations` are read
- THEN `recommendations` is `list[RecommendationDTO]`, `verdict` is `PublicationVerdictDTO`, and `eumic_violations` is `list[EumicViolationDTO]`

---

### Requirement: ReportExportUnavailable Exception

`ReportExportUnavailable(SrcBaseWarning)` MUST be defined in `src/domain/exceptions/report_errors.py` with `MESSAGE = "The report export service is unavailable (python-docx not installed)."`.

#### Scenario: Exception carries the expected message

- GIVEN `ReportExportUnavailable` is raised and caught
- WHEN its `MESSAGE` attribute is read
- THEN it equals the defined string literal

---

### Requirement: DocxReportAdapter Hard-Fails Without python-docx

`DocxReportAdapter.__init__` MUST raise `ReportExportUnavailable` at construction time when `DOCX_AVAILABLE` is `False`. The system SHALL NOT proceed to serve requests without python-docx installed.

#### Scenario: Adapter raises at construction when python-docx is absent

- GIVEN `DOCX_AVAILABLE = False`
- WHEN `DocxReportAdapter(logo_path=None)` is called
- THEN `ReportExportUnavailable` is raised before `__init__` returns

#### Scenario: Adapter initializes normally when python-docx is present

- GIVEN `DOCX_AVAILABLE = True`
- WHEN `DocxReportAdapter(logo_path=None)` is called
- THEN no exception is raised

---

### Requirement: DocxReportAdapter Produces Functionally Equivalent DOCX

`DocxReportAdapter.export()` MUST produce a `.docx` containing the same 11 sections as `WordExporter`, in order: title page, executive summary, document info, classification, quality analysis, grammar, APA validation, structure validation, citations, recommendations, footer. All 13 private rendering methods MUST remain private. Error wrapping is handled by `@generic_error_handler` at the use case layer, not the adapter.

The `_add_recommendations` method MUST render the publication verdict from `report_input.verdict` (a `PublicationVerdictDTO`) before the specific recommendations list. Color mapping: `PublicationVerdict.CRITICAL` and `PublicationVerdict.WARNING` → `reject_color_rgb`; `PublicationVerdict.APPROVED` → `publishable_color_rgb`. Specific recommendations use `RecommendationPriority` icons (`HIGH` → 🔴, `MEDIUM` → 🟡, `LOW` → 🟢) and attribute access (`rec.priority`, `rec.message`).

> Updated in Slice 13: `_add_recommendations` was refactored from dict-key access to attribute access. Verdict rendering was separated from the recommendations list (uses `report_input.verdict` directly, not a special entry in `recommendations`).

#### Scenario: Successful export returns True

- GIVEN a valid `ReportInputDTO` and a writable `path`
- WHEN `DocxReportAdapter.export(report_input, path)` is called
- THEN it returns `True` and a `.docx` file exists at `path`

#### Scenario: All 11 sections are present in the output

- GIVEN a valid `ReportInputDTO`
- WHEN `export()` completes
- THEN the output `.docx` contains: title page, executive summary, document info, classification, quality analysis, grammar, APA validation, structure validation, citations, recommendations, footer

#### Scenario: IO error propagates to caller

- GIVEN a path that cannot be written
- WHEN `DocxReportAdapter.export(report_input, path)` is called
- THEN the underlying `OSError` propagates; `@generic_error_handler` on the use case wraps it as `SrcGenericError`

---

### Requirement: ExportReportUseCase Delegates to Port

`ExportReportUseCase` MUST accept `ReportExportPort` via constructor injection. `execute(report_input, path) -> bool` MUST be decorated with `@generic_error_handler` and MUST delegate directly to the injected port.

#### Scenario: Use case returns port result

- GIVEN `FakeReportExportPort` configured to return `True`
- WHEN `ExportReportUseCase.execute(report_input, path)` is called
- THEN it returns `True`

#### Scenario: Use case propagates port exceptions

- GIVEN `FakeReportExportPort` configured to raise `ReportExportUnavailable`
- WHEN `execute(report_input, path)` is called
- THEN `ReportExportUnavailable` propagates to the caller

---

### Requirement: FakeReportExportPort Enables Test Isolation

`FakeReportExportPort(ReportExportPort)` MUST reside in `src/domain/tests/report/` and support a configurable return value and an optional exception raise. It MUST satisfy the full port contract.

#### Scenario: Fake returns configured boolean

- GIVEN `FakeReportExportPort(return_value=False)`
- WHEN `export(report_input, path)` is called
- THEN it returns `False`

#### Scenario: Fake raises configured exception

- GIVEN `FakeReportExportPort(raise_error=ReportExportUnavailable)`
- WHEN `export(report_input, path)` is called
- THEN `ReportExportUnavailable` is raised

---

### Requirement: ExportReportWiring Assembles the Full Object Graph

`ExportReportWiring.create_use_case() -> ExportReportUseCase` MUST resolve `logo_path` relative to the project root, instantiate `DocxReportAdapter`, and return a fully wired `ExportReportUseCase`. It MUST raise `ReportExportUnavailable` at startup if python-docx is absent.

#### Scenario: Wiring returns a fully wired use case

- GIVEN python-docx is installed
- WHEN `ExportReportWiring.create_use_case()` is called
- THEN it returns an `ExportReportUseCase` whose port is a `DocxReportAdapter`

#### Scenario: Wiring fails at startup without python-docx

- GIVEN python-docx is not installed
- WHEN `ExportReportWiring.create_use_case()` is called
- THEN `ReportExportUnavailable` is raised

---

## Constraints

| Constraint | Detail |
|------------|--------|
| Zero existing file modifications | 15 new files only; `presentation/word_exporter.py` stays untouched |
| Coexistence invariant | No `src/` file imports from `presentation/`; both layers coexist until Slice 16 |
| Callers out of scope | `main.py` and `gradio_app.py` continue using `WordExporter` until Slice 16 |
| Recommendations | `list[RecommendationDTO]` — typed and validated (TODO resolved in Slice 13) |
| Hard-fail at startup | `ReportExportUnavailable` raised in `__init__` — not deferred |
| Domain isolation | Domain imports nothing from infrastructure or application |
| One class per file | No multi-class modules |
| Test coverage | All new files covered by `unittest.TestCase` tests following port/fake/integration pattern |
