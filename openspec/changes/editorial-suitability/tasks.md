# Tasks: Editorial Suitability Analysis

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 500-600 |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 (Domain) → PR 2 (Integration) → PR 3 (Report/UI) |
| Delivery strategy | ask-on-risk |
| Chain strategy | feature-branch-chain |

Decision needed before apply: Yes
Chained PRs recommended: Yes
Chain strategy: feature-branch-chain
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Focused test command | Runtime harness | Rollback boundary |
|------|------|-----------|----------------------|-----------------|-------------------|
| 1 | Core DTOs and prompts | PR 1 | `.venv/Scripts/pytest src/domain/tests/quality/test_editorial_suitability_parser.py` | N/A: Domain logic validated by unit tests | Delete new domain files and prompts |
| 2 | Orchestration and wiring | PR 2 | `.venv/Scripts/pytest src/domain/tests/quality/test_quality_analyzer.py` | N/A: Orchestrator integration validated by unit tests | Revert QualityAnalyzer and wiring changes |
| 3 | Report rendering and UI | PR 3 | `.venv/Scripts/pytest src/infrastructure/tests/test_export_report_wiring.py` | `python gradio_app.py` | Revert docx adapter and gradio app changes |

## Phase 1: Foundation (DTOs, Prompts, Enums)

- [x] 1.1 Create `src/domain/dtos/editorial_suitability_dto.py` with 6 string fields matching design.
- [x] 1.2 Add optional `editorial_suitability` field to `src/domain/dtos/quality_result_dto.py`.
- [x] 1.3 Update `to_dict()` and `__str__()` in `src/domain/dtos/analysis_result_dto.py` to support suitability.
- [x] 1.4 Create `src/infrastructure/resources/prompts/quality/contribution_prompt.txt` with template.
- [x] 1.5 Create `src/infrastructure/resources/prompts/quality/alignment_prompt.txt` with template.
- [x] 1.6 **[RED]** Write unit tests in `src/domain/tests/enums/test_quality_level.py` for `get_quality_level_from_score`.
- [x] 1.7 **[GREEN]** Implement `get_quality_level_from_score()` and threshold constants in `src/domain/enums/quality_level.py`.

## Phase 2: Domain Services & Parsers (TDD)

- [x] 2.1 **[RED]** Write unit tests in `src/domain/tests/quality/test_editorial_suitability_parser.py` covering case insensitivity, boundary truncation, trailing `…`, and verdict consistency.
- [x] 2.2 **[GREEN]** Implement `EditorialSuitabilityParser` in `src/domain/quality/editorial_suitability_parser.py` parsing raw texts.
- [x] 2.3 **[RED]** Write unit tests in `src/domain/tests/quality/test_editorial_suitability_analyzer.py` verifying temperature, num_predict, and exactly 2 LLM generator calls.
- [x] 2.4 **[GREEN]** Implement `EditorialSuitabilityAnalyzer` in `src/domain/quality/editorial_suitability_analyzer.py` executing contribution and alignment LLM calls.
- [x] 2.5 **[RED]** Update constructor tests in `test_quality_analyzer.py` to expect 5 collaborators.
- [x] 2.6 **[GREEN]** Refactor `QualityAnalyzer` constructor in `src/domain/quality/quality_analyzer.py` to accept 5 collaborators, call enum function directly, and delegate suitability analysis.
- [x] 2.7 **[REFACTOR]** Delete deprecated `src/domain/quality/quality_level_resolver.py` and its tests in `src/domain/tests/quality/test_quality_level_resolver.py`.
- [x] 2.8 **[RED]** Write unit tests for `FileGatewayAdapter.read()`/`write()` in `src/infrastructure/tests/adapters/gateway/test_file_gateway_adapter.py` covering UTF-8 content (accented Spanish text).
- [x] 2.9 **[GREEN]** Fix `src/infrastructure/adapters/gateway/file_gateway_adapter.py` to open files with `encoding="utf-8"` in `read()` and `write()`.
- [x] 2.10 Create `src/infrastructure/resources/prompts/quality/research_lines.txt` with the 7 FMC research lines (content pending editorial validation).
- [x] 2.11 **[RED]** Update `test_editorial_suitability_analyzer.py` to expect `research_lines` as an injected constructor string instead of the module-level `_RESEARCH_LINES` constant.
- [x] 2.12 **[GREEN]** Refactor `EditorialSuitabilityAnalyzer` to accept `research_lines: str` in its constructor and remove the `_RESEARCH_LINES` module constant.
- [x] 2.13 Update `_get_editorial_suitability_analyzer()` in `analyze_document_use_case_wiring.py` to read `research_lines.txt` via `FileGatewayAdapter` and pass it to `EditorialSuitabilityAnalyzer`.

## Phase 3: Wiring, Adapters & Gradio UI

- [x] 3.1 **[RED]** Write integration test in `src/infrastructure/tests/test_docx_report_adapter.py` verifying Word report formatting and rendering of editorial suitability.
- [x] 3.2 **[GREEN]** Update `docx_report_adapter.py` to add `_add_editorial_suitability(doc, report_input)` and call it in `export()` pipeline.
- [x] 3.3 Update `analyze_document_use_case_wiring.py` to wire `EditorialSuitabilityAnalyzer` and remove `QualityLevelResolver`.
- [x] 3.4 Update `gradio_app.py` results view HTML to render editorial suitability results.
