# Tasks: Resolve Magic Values Debt

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~150 lines |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | single-pr |
| Chain strategy | size-exception |

Decision needed before apply: Yes
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Complete magic value parameterization, wiring, and tests | PR 1 | Base branch; tests and configuration included |

## Phase 1: Foundation & Settings

- [x] 1.1 Add new configurations to `.env` and `.env.example` (`STRUCTURE_MAX_HEADER_LENGTH`, `CITATION_MAX_AUTHOR_NAME_LENGTH`, `GRAMMAR_MAX_REPLACEMENTS`, `REPORT_WORDS_PER_PAGE`, `REPORT_MAX_ERRORS_DISPLAYED`, `REPORT_CONTEXT_TRUNCATION_LIMIT`, `REPORT_MAX_REPLACEMENTS`). Manually added by maintainer (agent file-access permissions denied `.env`/`.env.example`).
- [x] 1.2 Modify [docx_report_settings.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_settings.py) to add `words_per_page`, `max_errors_displayed`, `context_truncation_limit`, and `max_replacements` as fields with default factories reading from environment.

## Phase 2: Domain and Adapters Parameterization

- [x] 2.1 Parameterize [structure_validator.py](file:///E:/Python/silvina-editorial/src/domain/structure/structure_validator.py) with `max_header_length: int = 100` in `__init__` and use it in `_extract_present_sections`.
- [x] 2.2 Parameterize [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py) with `max_author_name_length: int = 100` in `__init__` and use it in `_collect_multi_author`.
- [x] 2.3 Import `IGNORECASE` and `DOTALL` from `re` in [docx_reference_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_reference_adapter.py) and replace raw integer flags `2 | 16` with symbolic flags.
- [x] 2.4 Parameterize [language_tool_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/grammar/language_tool_adapter.py) with a required `max_replacements: int` in `__init__` (no default — the sole default lives in the wiring's env-var read). Use it in `_map_to_dto`. Keep the existing manual try/except in `check()`/`_initialize_tool_if_needed` raising `GrammarCheckUnavailable` — the `@generic_error_handler` decorator is not applied here, since it's scoped to the use-case boundary, not adapters.
- [x] 2.5 Conditionally import `docx` in [docx_report_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_adapter.py) using try-except, setting `DOCX_AVAILABLE`. Raise `ReportExportUnavailable` in `__init__` if `DOCX_AVAILABLE` is `False`. Use settings parameters for layout checks.

## Phase 3: Wiring & Integration

- [x] 3.1 Update [analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py) to fetch `STRUCTURE_MAX_HEADER_LENGTH`, `CITATION_MAX_AUTHOR_NAME_LENGTH`, and `GRAMMAR_MAX_REPLACEMENTS` from env and pass to respective constructors.
- [x] 3.2 Update [export_report_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/export_report_wiring.py) to pass `DocxReportSettings` to `DocxReportAdapter`. No `load_dotenv` call added here — redundant, since `AnalyzeDocumentUseCaseWiring`'s module-level `load_dotenv()` already runs first in every real entry point.

## Phase 4: Testing & Verification

- [x] 4.1 Write unit tests in `src/domain/tests/structure/test_structure_validator_aliases.py` (or a new file) verifying the custom header length parameter works.
- [x] 4.2 Write unit tests in `src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py` verifying custom author name length works.
- [x] 4.3 Write unit tests in `src/infrastructure/tests/test_language_tool_adapter.py` verifying custom max replacements.
- [x] 4.4 Update `src/infrastructure/tests/test_analyze_document_use_case_wiring.py` and `src/infrastructure/tests/test_export_report_wiring.py` to assert correct setting propagation from mocked env vars.
- [x] 4.5 Update `src/infrastructure/tests/adapters/report/test_docx_report_adapter_init.py` to verify that `DOCX_AVAILABLE = False` raises `ReportExportUnavailable` on adapter initialization.
- [x] 4.6 Run the test suite (`pytest`) to ensure all tests pass successfully.
