# SDD Tasks — analyze-document-orchestrator

**Change**: analyze-document-orchestrator
**Phase**: tasks
**Date**: 2026-06-30
**Status**: active
**TDD**: STRICT (RED → GREEN → REFACTOR)
**Test runner**: `python -m pytest src/`

---

## Phase 1 — SCAFFOLD (parallel, no tests)

### T-01 [SCAFFOLD] Create `src/domain/recommendation/` package
- Create `src/domain/recommendation/__init__.py` (empty)
- Spec ref: Proposal §Files Affected, Design §Architecture Overview

### T-02 [SCAFFOLD] Create `src/domain/tests/recommendation/` package
- Create `src/domain/tests/recommendation/__init__.py` (empty)
- Spec ref: Proposal §Files Affected, Design §Architecture Overview

---

## Phase 2 — DOMAIN ENUM & DTO EXTENSIONS (sequential TDD loop)

### T-03 [RED] Write failing tests for extended `RecommendationPriority` values
- File: `src/domain/tests/enums/test_recommendation_priority.py`
- Add tests to assert that `CRITICAL`, `WARNING`, and `APPROVED` are members of `RecommendationPriority` and map to `"critica"`, `"advertencia"`, and `"aprobado"` respectively.
- Assert that `len(RecommendationPriority)` is 6.
- Run tests and verify failure.

### T-04 [GREEN] Extend `RecommendationPriority` enum
- File: `src/domain/enums/recommendation_priority.py`
- Add:
  - `CRITICAL = "critica"`
  - `WARNING = "advertencia"`
  - `APPROVED = "aprobado"`
- Run `python -m pytest src/domain/tests/enums/test_recommendation_priority.py` → green.

### T-05 [RED] Write failing tests for IMRyD classification override property
- File: `src/domain/tests/dtos/test_classification_result.py`
- Assert `effective_structure_type` property behavior:
  - `article_type = ArticleType.CIENTIFICO` and `"IMRyD"` in `reasoning` → returns `ArticleType.CIENTIFICO`
  - `article_type = ArticleType.CIENTIFICO` and `"IMRyD"` NOT in `reasoning` → returns `ArticleType.DIVULGACION`
  - Other article types (e.g. `OPINION`, `DIVULGACION`) → returned as-is.
- Run tests and verify failure.

### T-06 [GREEN] Implement `effective_structure_type` property on `ClassificationResultDTO`
- File: `src/domain/dtos/classification_result_dto.py`
- Add read-only `@property` `effective_structure_type` to `ClassificationResultDTO`.
- Run `python -m pytest src/domain/tests/dtos/test_classification_result.py` → green.

### T-07 [SCAFFOLD] Define DTO classes `RecommendationDTO` and `RecommendationSettings`
- File: `src/domain/dtos/recommendation_dto.py`
  - Define `RecommendationDTO` inheriting from `BaseDTO`, `frozen=True`.
  - Fields: `priority: RecommendationPriority`, `message: str`.
- File: `src/domain/recommendation/recommendation_settings.py`
  - Define `RecommendationSettings` inheriting from `BaseDTO`, `frozen=True`.
  - Fields and default float/int values for publish, quality, grammar, dimension, citation_match, critical_citation_match, citation_count, and classification_confidence thresholds.

### T-08 [RED] Write failing tests for `RecommendationBuilder`
- File: `src/domain/tests/recommendation/test_recommendation_builder.py`
- One `TestCase` class: `TestRecommendationBuilder`
- Scenarios covered:
  - Default settings initialization.
  - Quality score below threshold triggers `HIGH` priority recommendation.
  - Grammar score below threshold triggers `HIGH` priority recommendation.
  - Individual dimension scores below threshold trigger `MEDIUM` priority recommendations.
  - Missing structural sections trigger `HIGH` priority recommendations.
  - Citation match rates below threshold triggers `HIGH` priority recommendation.
  - Citation match rates with unmatched count > 0 triggers `MEDIUM` priority recommendation.
  - Citation count below threshold triggers `MEDIUM` priority recommendation.
  - Classification confidence below threshold triggers `LOW` priority recommendation.
  - Publication recommendation decisions: `CRITICAL` (for critical issues or 0 citations), `WARNING` (for warning issues/violations), `APPROVED` (if clean).
- Run tests and verify failure.

### T-09 [GREEN] Implement `RecommendationBuilder`
- File: `src/domain/recommendation/recommendation_builder.py`
- Implement validation rules in the `build()` method per Design §Component Interfaces.
- Run `python -m pytest src/domain/tests/recommendation/test_recommendation_builder.py` → green.

### T-10 [REFACTOR] Update `ReportInputDTO`
- File: `src/domain/dtos/report_input_dto.py`
- Update `recommendations` type hint to `list[RecommendationDTO]`.
- Add `eumic_violations: list[EumicViolationDTO]` field.
- Verify `is_publishable` and `publishability_reason` logic is intact.
- Run `python -m pytest src/domain/tests/dtos/test_analysis_result.py` → green.

---

## Phase 3 — APPLICATION LAYER (sequential TDD loop)

### T-11 [RED] Write failing tests for `AnalyzeDocumentUseCase`
- File: `src/application/tests/test_analyze_document_use_case.py`
- One `TestCase` class: `TestAnalyzeDocumentUseCase`
- Scenarios:
  - Executing complete pipeline using mock use case dependencies.
  - Verify sequence of calls is correct.
  - Verify citation filtering (only `CitationType.AUTHOR_YEAR` are validated via `ValidateApaUseCase`).
  - Verify paragraph bounds mapping for citation validation location indices.
  - Verify EUMIC violations are correctly captured and added to `ReportInputDTO` without halting execution.
  - Verify that `effective_structure_type` of classification is passed to `ValidateStructureUseCase`.
- Run tests and verify failure.

### T-12 [GREEN] Implement `AnalyzeDocumentUseCase`
- File: `src/application/analyze_document_use_case.py`
- Implement the orchestrator coordinating:
  - `ReadDocumentUseCase`, `ExtractContentUseCase`, `ExtractCitationsUseCase`, `ValidateApaUseCase`, `CheckGrammarUseCase`, `ClassifyArticleUseCase`, `AnalyzeQualityUseCase`, `ValidateStructureUseCase`, `MatchCitationsUseCase`, `VerifyEumicUseCase`, and `RecommendationBuilder`.
- Use the `@generic_error_handler` decorator on `execute`.
- Implement citation filtering and mapping to paragraph text index before passing to `ValidateApaUseCase` per Design §ADR-5.
- Run `python -m pytest src/application/tests/test_analyze_document_use_case.py` → green.

---

## Phase 4 — INFRASTRUCTURE ADAPTERS & WIRING (sequential TDD loop)

### T-13 [REFACTOR] Update mock fixtures in report tests
- File: `src/infrastructure/tests/adapters/report/fixtures.py`
- Update `ReportFixtures.make_report_input_dto` to include real `RecommendationDTO` and `EumicViolationDTO` instances, adapting mocks to use attribute-based models.

### T-14 [REFACTOR] Update `DocxReportAdapter`
- File: `src/infrastructure/adapters/report/docx_report_adapter.py`
- Refactor the `_add_recommendations` method to use attribute-based access (e.g. `rec.priority` and `rec.message`) rather than dictionary lookup.
- Update lookup maps to handle the new `RecommendationPriority` enums instead of Spanish strings directly if needed, or mapping `rec.priority` value / checking if priority is in `[RecommendationPriority.CRITICAL, RecommendationPriority.WARNING, RecommendationPriority.APPROVED]`.
- Verify report generation works by running existing report adapter tests.

### T-15 [RED] Write failing tests for `AnalyzeDocumentUseCaseWiring`
- File: `src/infrastructure/tests/test_analyze_document_use_case_wiring.py`
- One `TestCase` class: `TestAnalyzeDocumentUseCaseWiring`
- Scenarios:
  - Successful factory dependency resolution with `create_use_case()`.
  - Environmental variables override defaults for recommendation thresholds.
- Assert that calling `create_use_case()` returns an instance of `AnalyzeDocumentUseCase`.
- Run tests and verify failure.

### T-16 [GREEN] Implement `AnalyzeDocumentUseCaseWiring`
- File: `src/infrastructure/wirings/analyze_document_use_case_wiring.py`
- Implement the factory class initializing settings from `os.getenv` with the default thresholds, constructing `RecommendationSettings`, `RecommendationBuilder`, and wiring the 10 use cases to `AnalyzeDocumentUseCase`.
- Run `python -m pytest src/infrastructure/tests/test_analyze_document_use_case_wiring.py` → green.

---

## Phase 5 — FULL VERIFICATION

### T-17 [VERIFY] Run full test suite
- Command: `python -m pytest src/`
- Assertions:
  - All new tests pass.
  - Legacy tests run without errors.
  - Zero regressions on report generation or structural validation.

---

## Dependency Graph

```
T-01 ────────────────────────────────────────────────────────┐
T-02 ────────────────────────────────────────────────────────┼─► T-08 ──► T-09 ──┐
T-03 ──► T-04 ───────────────────────────────────────────────┤                   │
T-05 ──► T-06 ───────────────────────────────────────────────┤                   ├──► T-10 ──► T-11 ──► T-12 ──► T-13 ──► T-14 ──► T-15 ──► T-16 ──► T-17
T-07 ────────────────────────────────────────────────────────┘
```

---

## Review Workload Forecast

| Category | Files | Est. Lines |
|----------|-------|------------|
| Scaffold `__init__.py` | 2 | ~2 |
| `recommendation_priority.py` (modify) | 1 | ~10 |
| `test_recommendation_priority.py` (modify) | 1 | ~15 |
| `classification_result_dto.py` (modify) | 1 | ~10 |
| `report_input_dto.py` (modify) | 1 | ~25 |
| `recommendation_dto.py` (create) | 1 | ~15 |
| `recommendation_settings.py` (create) | 1 | ~25 |
| `recommendation_builder.py` (create) | 1 | ~65 |
| `test_recommendation_builder.py` (create) | 1 | ~120 |
| `analyze_document_use_case.py` (create) | 1 | ~90 |
| `test_analyze_document_use_case.py` (create) | 1 | ~180 |
| `analyze_document_use_case_wiring.py` (create) | 1 | ~60 |
| `test_analyze_document_use_case_wiring.py` (create) | 1 | ~40 |
| `docx_report_adapter.py` (modify) | 1 | ~30 |
| `fixtures.py` (modify) | 1 | ~15 |
| **Total** | **16** | **~702 lines** |

**Chained PRs recommended**: No
**Chain strategy**: size-exception (approved by user)
**400-line budget risk**: Medium (slice implemented as single development unit)
**Decision needed before apply**: No
