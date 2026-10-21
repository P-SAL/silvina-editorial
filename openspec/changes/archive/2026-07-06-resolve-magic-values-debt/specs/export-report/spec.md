# Delta Spec — export-report

**Change**: resolve-magic-values-debt

## MODIFIED Requirements

### Requirement: DocxReportAdapter Hard-Fails Without python-docx

`DocxReportAdapter.__init__` MUST raise `ReportExportUnavailable` at construction time when `DOCX_AVAILABLE` is `False`. The constructor MUST accept an optional `settings: DocxReportSettings` parameter (defaulting to a default-constructed `DocxReportSettings` if not provided). The system SHALL NOT proceed to serve requests without python-docx installed.

(Previously: `DocxReportAdapter.__init__` only accepted `logo_path: str | None` and did not accept `settings`.)

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

The `DocxReportAdapter` MUST utilize configuration settings from `DocxReportSettings` instead of hardcoded magic values:
- Use `settings.words_per_page` for estimating pages (`estimated_pages = word_count // words_per_page`).
- Use `settings.max_errors_displayed` to limit the number of displayed grammar errors and APA validation violations.
- Use `settings.context_truncation_limit` to truncate grammar error context.
- Use `settings.max_replacements` to limit the number of replacements shown per grammar error.

The `_add_recommendations` method MUST render the publication verdict from `report_input.verdict` (a `PublicationVerdictDTO`) before the specific recommendations list. Color mapping: `PublicationVerdict.CRITICAL` and `PublicationVerdict.WARNING` → `reject_color_rgb`; `PublicationVerdict.APPROVED` → `publishable_color_rgb`. Specific recommendations use `RecommendationPriority` icons (`HIGH` → 🔴, `MEDIUM` → 🟡, `LOW` → 🟢) and attribute access (`rec.priority`, `rec.message`).

(Previously: layout dimensions and limit parameters like 250 words per page, maximum of 5 displayed errors, 150 character context truncation limit, and maximum of 3 replacements were hardcoded.)

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

#### Scenario: Settings are respected during rendering
- GIVEN a `DocxReportSettings` with `words_per_page` = 100, `max_errors_displayed` = 2, `context_truncation_limit` = 10, `max_replacements` = 1
- WHEN `DocxReportAdapter.export()` is called with this configuration
- THEN the exported report uses page estimation based on 100 words per page
- AND at most 2 grammar/APA errors are listed
- AND error contexts are truncated to 10 characters
- AND at most 1 replacement is shown per error

---

### Requirement: ExportReportWiring Assembles the Full Object Graph

`ExportReportWiring.create_use_case() -> ExportReportUseCase` MUST resolve `logo_path` relative to the project root, instantiate `DocxReportSettings` (loading configuration from the environment), instantiate `DocxReportAdapter` with that settings object, and return a fully wired `ExportReportUseCase`. It MUST raise `ReportExportUnavailable` at startup if python-docx is absent.

(Previously: `ExportReportWiring.create_use_case()` did not instantiate or inject a configured `DocxReportSettings` into `DocxReportAdapter`.)

#### Scenario: Wiring returns a fully wired use case
- GIVEN python-docx is installed
- WHEN `ExportReportWiring.create_use_case()` is called
- THEN it returns an `ExportReportUseCase` whose port is a `DocxReportAdapter`
- AND the adapter's settings are loaded from environment variables `REPORT_WORDS_PER_PAGE`, `REPORT_MAX_ERRORS_DISPLAYED`, `REPORT_CONTEXT_TRUNCATION_LIMIT`, and `REPORT_MAX_REPLACEMENTS`

#### Scenario: Wiring fails at startup without python-docx
- GIVEN python-docx is not installed
- WHEN `ExportReportWiring.create_use_case()` is called
- THEN `ReportExportUnavailable` is raised
