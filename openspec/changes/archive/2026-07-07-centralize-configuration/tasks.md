# Tasks: Centralize Configuration

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 350-450 |
| 400-line budget risk | Medium |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | single-pr |
| Chain strategy | size-exception |

Decision needed before apply: Yes
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Medium

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Centralize config, refactor adapters/wirings, and clean up | PR 1 | Base branch; tests and cleanup included |

## Phase 1: Foundation / Infrastructure

- [x] 1.1 Create `src/infrastructure/env_config.py` defining `EnvConfig` with all environment variables. Ensure NO `import os` is allowed; use specific name imports (e.g., `from os import getenv`).
- [x] 1.2 Add environment variables to `.env` and `.env.example`, organized by theme with section comments (renamed to drop the `RECOMMENDATION_` prefix; reopened `SILVINA_APP_NAME`/`SILVINA_VERSION`/`REPORT_SCORE_HIGH_THRESHOLD`/`REPORT_SCORE_MEDIUM_THRESHOLD` for env injection). Done manually by the user (tooling permissions block `.env*` access).

## Phase 2: Core Implementation

- [x] 2.1 Refactor `DocxReportSettings` in `src/infrastructure/adapters/report/docx_report_settings.py` to use static defaults and remove direct environment reads. Use PEP 604 union types; do NOT use `import os` or `from os import environ`.
- [x] 2.2 Refactor `AnalyzeDocumentUseCaseWiring` in `src/infrastructure/wirings/analyze_document_use_case_wiring.py` to instantiate `EnvConfig` and inject settings. Memoize and share a single `LlmGeneratorPort`. Do NOT use `import os`.
- [x] 2.3 Refactor `ExportReportWiring` in `src/infrastructure/wirings/export_report_wiring.py` to instantiate `EnvConfig` and construct `DocxReportSettings` using configuration values. Check for `python-docx` availability, raising `ReportExportUnavailable` if missing. Do NOT use `import os`.

## Phase 3: Testing / Verification

- [x] 3.1 Create `src/infrastructure/tests/test_env_config.py` to test defaults, overrides, casting, and DTO construction. Use pure Python and `unittest.TestCase`. Do NOT use `import os`.
- [x] 3.2 Update `src/infrastructure/tests/test_analyze_document_use_case_wiring.py` to verify wiring dependencies and LLM generator sharing. Do NOT use `import os`.
- [x] 3.3 Update `src/infrastructure/tests/adapters/report/test_docx_report_settings.py` to assert static defaults and constructor arguments without environment patching. Do NOT use `import os`.
- [x] 3.4 Update `src/infrastructure/tests/test_export_report_wiring.py` to verify wiring behavior and the raising of `ReportExportUnavailable` if docx is absent. Do NOT use `import os`.

## Phase 4: Cleanup

- [x] 4.1 Delete the deprecated config file `src/infrastructure/config/recommendation_config.py`.
- [x] 4.2 Run formatting and style check to ensure compliance with Clean Architecture rules: no abbreviations, specific name imports only, and single class per file.
