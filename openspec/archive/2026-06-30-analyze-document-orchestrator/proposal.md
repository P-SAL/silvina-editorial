# SDD Proposal — analyze-document-orchestrator

> **Status**: proposed
> **Change**: analyze-document-orchestrator
> **Date**: 2026-06-30
> **Artifact store**: hybrid (engram + openspec/)

---

## Intent

Implement the top-level orchestration for complete document analysis, migrating the logic from legacy `SilvinaEditorialAssistant.analyze_document()` and `_generate_recommendations()` in [main.py](file:///E:/Python/silvina-editorial/main.py) into the clean hexagonal architecture codebase:
1. **`RecommendationDTO`**: Immutable DTO representing a single quality or formatting recommendation.
2. **`RecommendationSettings`**: Configuration settings for generating recommendations, loaded from environment variables.
3. **`RecommendationBuilder`**: Domain service that encapsulates recommendation generation logic based on analysis inputs and settings.
4. **`AnalyzeDocumentUseCase`**: Application orchestrator coordinating reading, content extraction, citation parsing, APA validation, grammar checking, classification, quality analysis, structure validation, citation matching, and EUMIC format verification.
5. **`AnalyzeDocumentUseCaseWiring`**: Assembly factory wiring all sub-use cases and dependencies, injecting settings from environment variables.

This slice consolidates all document analysis sub-use cases under a single orchestrator, providing a unified `ReportInputDTO` suitable for document generation adapters (such as `DocxReportAdapter`).

**Success**: The `AnalyzeDocumentUseCase` executes the entire pipeline, maps results cleanly, produces recommendations matching the legacy rules (with customizable thresholds), and returns a validated `ReportInputDTO` containing EUMIC violations.

---

## Scope

### In scope

1. **`src/domain/enums/recommendation_priority.py`** — Extend `RecommendationPriority` enum with English members mapping to Spanish string values for legacy compatibility:
   - `HIGH = "alta"`
   - `MEDIUM = "media"`
   - `LOW = "baja"`
   - `CRITICAL = "critica"`
   - `WARNING = "advertencia"`
   - `APPROVED = "aprobado"`
2. **`src/domain/dtos/recommendation_dto.py`** — Define `RecommendationDTO` (inherits from `BaseDTO`, frozen=True):
   - Fields: `priority: RecommendationPriority`, `message: str`
3. **`src/domain/recommendation/recommendation_settings.py`** — Create a pure domain settings DTO `RecommendationSettings` to hold threshold limits.
4. **`src/domain/recommendation/recommendation_builder.py`** — Domain service generating a list of `RecommendationDTO` based on analysis inputs and `RecommendationSettings`.
5. **`src/domain/dtos/classification_result_dto.py`** — Add a read-only property `effective_structure_type -> ArticleType` to encapsulate the IMRyD scientific classification override logic:
   - If `article_type` is `CIENTIFICO` and `"IMRyD"` is in the classification `reasoning`, return `ArticleType.CIENTIFICO`.
   - If `article_type` is `CIENTIFICO` and `"IMRyD"` is NOT in the `reasoning`, return `ArticleType.DIVULGACION`.
   - For all other cases, return `article_type` as-is.
6. **`src/domain/dtos/report_input_dto.py`** — Update fields:
   - `recommendations: list[RecommendationDTO]`
   - `eumic_violations: list[EumicViolationDTO]`
7. **`src/application/analyze_document_use_case.py`** — Implement the orchestrator coordinating:
   - `ReadDocumentUseCase`, `ExtractContentUseCase`, `ExtractCitationsUseCase`, `ValidateApaUseCase`, `CheckGrammarUseCase`, `ClassifyArticleUseCase`, `AnalyzeQualityUseCase`, `ValidateStructureUseCase`, `MatchCitationsUseCase`, `VerifyEumicUseCase`, and `RecommendationBuilder`.
   - Uses `classification.effective_structure_type` for structure validation.
   - Retains legacy EUMIC behavior: run verification, collect violations, and return them without throwing exceptions.
8. **`src/infrastructure/wirings/analyze_document_use_case_wiring.py`** — Wired class loading settings from `os.environ` via `os.getenv` with sensible defaults to build `RecommendationSettings`, instantiating the builder and orchestrator.
9. **`src/infrastructure/adapters/report/docx_report_adapter.py`** — Refactor to use attribute-based access on `recommendation` objects (`recommendation.priority` and `recommendation.message`) instead of key-based dict access.
10. **`src/infrastructure/tests/adapters/report/fixtures.py`** — Update mock fixtures to use `RecommendationDTO` and `EumicViolationDTO` instances.
11. **Unit and Integration Test Suites**:
    - `src/domain/tests/enums/test_recommendation_priority.py`: assert new enum members.
    - `src/domain/tests/recommendation/test_recommendation_builder.py`: test builder with different settings and scores.
    - `src/application/tests/test_analyze_document_use_case.py`: orchestrator flow tests using mocks/doubles.
    - `src/infrastructure/tests/test_analyze_document_use_case_wiring.py`: factory dependency resolution test.

### Explicitly out of scope

- Deleting any legacy files under `business_logic/` or modifying `main.py` in this slice (wiring of `main.py` to the new orchestrator is deferred to a future integration slice).
- Raising fatal exceptions on EUMIC violation (violations are returned in `ReportInputDTO` but do not block pipeline execution).
- Performing validation or checks for non-APA citation types (only `CitationType.AUTHOR_YEAR` is processed for APA violations).

---

## Behavioral Contracts

### Extended RecommendationPriority
The priority enum must map English identifier names to legacy Spanish string representation:
- `RecommendationPriority.HIGH -> "alta"`
- `RecommendationPriority.MEDIUM -> "media"`
- `RecommendationPriority.LOW -> "baja"`
- `RecommendationPriority.CRITICAL -> "critica"`
- `RecommendationPriority.WARNING -> "advertencia"`
- `RecommendationPriority.APPROVED -> "aprobado"`

### IMRyD Override Logic (Domain-level property)
`ClassificationResultDTO.effective_structure_type` property:
```python
@property
def effective_structure_type(self) -> ArticleType:
    if self.article_type == ArticleType.CIENTIFICO:
        if "IMRyD" in (self.reasoning or ""):
            return ArticleType.CIENTIFICO
        return ArticleType.DIVULGACION
    return self.article_type
```

### EUMIC Verification Behavior
`AnalyzeDocumentUseCase.execute` runs EUMIC validation:
```python
eumic_violations = self._verify_eumic_use_case.execute(
    docx_path=document_path, word_count=document_content.word_count
)
```
Any violations found are passed directly into the final `ReportInputDTO`. The process does not abort, reflecting the legacy behavior of printing the violations without raising errors.

### Configurable Thresholds Settings
`RecommendationSettings` fields are loaded from the environment with these defaults:
- `RECOMMENDATION_PUBLISH_THRESHOLD` (Default: `7.0`)
- `RECOMMENDATION_QUALITY_THRESHOLD` (Default: `7.0`)
- `RECOMMENDATION_GRAMMAR_THRESHOLD` (Default: `7.0`)
- `RECOMMENDATION_DIMENSION_THRESHOLD` (Default: `6.0`)
- `RECOMMENDATION_CITATION_MATCH_THRESHOLD` (Default: `90.0`)
- `RECOMMENDATION_CRITICAL_CITATION_MATCH_THRESHOLD` (Default: `50.0`)
- `RECOMMENDATION_CITATION_COUNT_THRESHOLD` (Default: `10`)
- `RECOMMENDATION_CLASSIFICATION_CONFIDENCE_THRESHOLD` (Default: `0.7`)

---

## Approach and Rationale

### Why put IMRyD logic on ClassificationResultDTO?
The IMRyD classification override determines how an article classified as CIENTIFICO should have its structure validated. Rather than hardcoding this string-matching and conversion logic inside the `AnalyzeDocumentUseCase` orchestrator, placing it as a read-only property `effective_structure_type` on the `ClassificationResultDTO` domain class keeps the orchestrator clean, respects SOLID principles, and exposes it as a clear business policy.

### Why use English identifiers for Spanish strings?
Clean hexagonal architecture conventions dictate that all code identifiers, enums, classes, and comments be in English. However, the downstream report generator (`DocxReportAdapter`) and legacy expectations use specific Spanish terms for priorities ("critica", "advertencia", "aprobado"). Mapping English enum members to these Spanish strings maintains code compliance with architecture conventions while ensuring backward-compatible behavior.

### Why RecommendationSettings is in the domain?
Moving the thresholds to a pure domain dataclass (`RecommendationSettings`) allows the recommendation building logic to remain a pure domain service (`RecommendationBuilder`) without environment variable retrieval (`os.environ`) side-effects. The wiring layer manages the interaction with the environment (`os.getenv`) and injects the populated settings into the builder.

---

## Files Affected

### Created

| Path | Description |
|------|-------------|
| `src/domain/dtos/recommendation_dto.py` | Immutable recommendation DTO |
| `src/domain/recommendation/__init__.py` | Package marker |
| `src/domain/recommendation/recommendation_settings.py` | Configurable thresholds settings class |
| `src/domain/recommendation/recommendation_builder.py` | Pure domain service for recommendation building |
| `src/application/analyze_document_use_case.py` | Top-level pipeline orchestrator use case |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Orchestrator factory and environment wiring |
| `src/domain/tests/recommendation/__init__.py` | Package marker |
| `src/domain/tests/recommendation/test_recommendation_builder.py` | Unit tests for RecommendationBuilder and Settings |
| `src/application/tests/test_analyze_document_use_case.py` | Orchestrator pipeline integration tests using mocks |
| `src/infrastructure/tests/test_analyze_document_use_case_wiring.py` | Factory wiring tests |

### Modified

| Path | Description |
|------|-------------|
| `src/domain/enums/recommendation_priority.py` | Add CRITICAL, WARNING, APPROVED members |
| `src/domain/tests/enums/test_recommendation_priority.py` | Update enum validation tests |
| `src/domain/dtos/classification_result_dto.py` | Add `effective_structure_type` property |
| `src/domain/dtos/report_input_dto.py` | Add type safety for recommendations and include `eumic_violations` |
| `src/infrastructure/adapters/report/docx_report_adapter.py` | Access recommendation attributes instead of keys |
| `src/infrastructure/tests/adapters/report/fixtures.py` | Adapt mock report fixtures to use real DTOs |

---

## Dependencies

- **Prior Slices**: All sub-use cases (`ReadDocumentUseCase`, `ExtractContentUseCase`, etc.) and their respective DTOs/ports must be fully implemented and wired.
- **Environment config**: `.env` file support (values are retrieved using `os.getenv` within the wiring file).

---

## Non-Goals

- Refactoring any sub-use cases (e.g. `ExtractContentUseCase`) or changes to their interface.
- Deleting the legacy `SilvinaEditorialAssistant` class or modifying the legacy `main.py` file.
- Modifying EUMIC check rules (just calling the existing `VerifyEumicUseCase` interface).

---

## Risks

1. **Incorrect Wiring / Circular Imports**: Introducing many use case dependencies into `AnalyzeDocumentUseCase` can cause circular imports if any sub-use case import is improperly structured.
   - *Mitigation*: Adhere strictly to the clean architecture import rules (application only imports domain; infrastructure imports application and domain).
2. **Missing Environment Variables**: If environment variables are missing from `.env` in production, recommendations could use wrong values.
   - *Mitigation*: Provide sensible, safe default values in the `os.getenv()` calls inside the wiring layer.
