# Delta for Report Export (centralize-configuration)

## MODIFIED Requirements

### Requirement: DocxReportSettings Configuration Object

`DocxReportSettings` MUST be a frozen dataclass at `src/infrastructure/adapters/report/docx_report_settings.py`. Its 8 deployment-config fields are **required** (no defaults) — the wiring layer MUST supply them from `EnvConfig`:
- `app_name: str`
- `app_version: str`
- `score_high_threshold: float`
- `score_medium_threshold: float`
- `words_per_page: int`
- `max_errors_displayed: int`
- `context_truncation_limit: int`
- `max_replacements: int`

Its remaining fields (Word template visual constants: fonts, colors, sizes, table style, layout dimensions) keep static defaults and are NOT sourced from `EnvConfig` — they are design constants, not deployment config.

`DocxReportSettings` MUST NOT read environment variables directly.

(Previously: All 8 deployment fields had static defaults (`words_per_page: int = 250`, etc.), allowing `DocxReportSettings()` to be instantiated with zero arguments. Defaults were removed to fail fast if the wiring layer omits a deployment value, mirroring the `RecommendationSettingsDTO` no-defaults pattern.)

#### Scenario: Instantiation without deployment fields raises TypeError

- GIVEN no arguments are passed
- WHEN `DocxReportSettings()` is instantiated
- THEN Python raises a `TypeError` due to missing required arguments

#### Scenario: Visual template fields use static defaults when only deployment fields are provided

- GIVEN all 8 deployment fields are provided and no visual fields are provided
- WHEN `DocxReportSettings(...)` is instantiated
- THEN `settings.font_name`, `settings.table_style`, `settings.heading_color_rgb`, and other visual fields match their static defaults

#### Scenario: Settings accepts arguments passed during instantiation

- GIVEN all 8 deployment fields are provided, with `words_per_page=300` and `max_errors_displayed=10` overriding the rest
- WHEN `DocxReportSettings(...)` is instantiated
- THEN `settings.words_per_page == 300` AND `settings.max_errors_displayed == 10`

---

### Requirement: DocxReportAdapter Hard-Fails Without python-docx

`DocxReportAdapter.__init__` MUST raise `ReportExportUnavailable` at construction time when `DOCX_AVAILABLE` is `False`. The constructor MUST accept `settings: DocxReportSettings` as a **required** parameter (no default, no fallback construction) and an optional `logo_path`. The system SHALL NOT proceed to serve requests without python-docx installed.

(Previously: `settings` defaulted to `None` and fell back to a default-constructed `DocxReportSettings()`. This fallback is no longer possible because `DocxReportSettings` has no defaults for its 8 deployment-config fields.)

#### Scenario: Adapter raises at construction when python-docx is absent

- GIVEN `DOCX_AVAILABLE = False` and a valid `settings` object
- WHEN `DocxReportAdapter(logo_path=None, settings=settings)` is called
- THEN `ReportExportUnavailable` is raised before `__init__` returns

#### Scenario: Adapter initializes normally when python-docx is present

- GIVEN `DOCX_AVAILABLE = True` and a valid `settings` object
- WHEN `DocxReportAdapter(logo_path=None, settings=settings)` is called
- THEN no exception is raised

#### Scenario: Adapter construction fails without settings

- GIVEN no `settings` argument is provided
- WHEN `DocxReportAdapter(logo_path=None)` is called
- THEN Python raises a `TypeError` due to the missing required argument

---

### Requirement: ExportReportWiring Assembles the Full Object Graph

`ExportReportWiring.create_use_case()` MUST resolve `logo_path` relative to the project root, instantiate `EnvConfig`, construct `DocxReportSettings` using values from the `EnvConfig` instance, instantiate `DocxReportAdapter` with that settings object, and return a fully wired `ExportReportUseCase`. It MUST raise `ReportExportUnavailable` at startup if python-docx is absent.

(Previously: Instantiated `DocxReportSettings` which read environment variables dynamically from `os.environ` via default factories.)

#### Scenario: Wiring returns a fully wired use case

- GIVEN python-docx is installed and `REPORT_WORDS_PER_PAGE=300` is set in the environment
- WHEN `ExportReportWiring.create_use_case()` is called
- THEN it returns an `ExportReportUseCase` whose port is a `DocxReportAdapter`
- AND the adapter's settings has `words_per_page == 300`

#### Scenario: Wiring fails at startup without python-docx

- GIVEN python-docx is not installed
- WHEN `ExportReportWiring.create_use_case()` is called
- THEN `ReportExportUnavailable` is raised
