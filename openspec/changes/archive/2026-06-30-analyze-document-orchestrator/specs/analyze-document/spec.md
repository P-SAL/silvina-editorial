# Specification: Document Analysis Orchestrator & Recommendations (Slice 13 + Refactor)

> Normative guide: `.agent/skills/clean-architecture/SKILL.md`
> Parent proposal: `openspec/changes/analyze-document-orchestrator/proposal.md`
> Migration plan reference: `docs/plan-migracion-hexagonal.md`
> **Refactor note**: Slice 14 integration observations caused intentional divergences from the original Slice 13 spec. This document reflects the final implemented state.

---

## Purpose

This specification defines the behavior of the document analysis pipeline orchestrator, the domain service that generates editorial recommendations, and the supporting DTOs and configuration settings in compliance with clean hexagonal architecture guidelines.

---

## Requirements

### Requirement: RecommendationPriority Enum (3 members)

The `RecommendationPriority` enum MUST live in `src/domain/enums/recommendation_priority.py` and carry exactly **three** values representing severity levels for specific editorial recommendations:

- `HIGH` maps to `"alta"`
- `MEDIUM` maps to `"media"`
- `LOW` maps to `"baja"`

> **Rationale**: `CRITICAL`, `WARNING`, and `APPROVED` were removed from this enum because they represent a *publication verdict* — a semantically distinct concept. They live in `PublicationVerdict` (see below).

#### Scenario: RecommendationPriority contains exactly three members

- GIVEN the `RecommendationPriority` enum
- WHEN the list of members is queried
- THEN it has exactly 3 members: `HIGH`, `MEDIUM`, and `LOW`

#### Scenario: RecommendationPriority values match Spanish strings

- GIVEN the `RecommendationPriority` members
- WHEN their values are read
- THEN the values map to `"alta"`, `"media"`, and `"baja"` respectively

---

### Requirement: PublicationVerdict Enum

`PublicationVerdict` MUST live in `src/domain/enums/publication_verdict.py` and carry exactly **three** values representing the final editorial verdict for a document:

- `CRITICAL` maps to `"critica"`
- `WARNING` maps to `"advertencia"`
- `APPROVED` maps to `"aprobado"`

#### Scenario: PublicationVerdict contains exactly three members

- GIVEN the `PublicationVerdict` enum
- WHEN the list of members is queried
- THEN it has exactly 3 members: `CRITICAL`, `WARNING`, and `APPROVED`

---

### Requirement: PublicationVerdictDTO

`PublicationVerdictDTO` MUST be a frozen data transfer object inheriting from `BaseDTO` in `src/domain/dtos/publication_verdict_dto.py` with fields:
- `verdict: PublicationVerdict`
- `message: str`

#### Scenario: PublicationVerdictDTO is immutable

- GIVEN a constructed `PublicationVerdictDTO`
- WHEN any field is modified or reassigned
- THEN Python raises `FrozenInstanceError`

---

### Requirement: Immutable RecommendationDTO

`RecommendationDTO` MUST be a frozen data transfer object inheriting from `BaseDTO` in `src/domain/dtos/recommendation_dto.py` with fields:
- `priority: RecommendationPriority`
- `message: str`

#### Scenario: RecommendationDTO is immutable

- GIVEN a constructed `RecommendationDTO`
- WHEN any field is modified or reassigned
- THEN Python raises `FrozenInstanceError`

---

### Requirement: ClassificationResultDTO IMRyD Override Property

`ClassificationResultDTO` (in `src/domain/dtos/classification_result_dto.py`) MUST expose a read-only property `effective_structure_type: ArticleType` using guard clauses:

1. If `article_type` is NOT `ArticleType.CIENTIFICO`, return `article_type` immediately.
2. If `"IMRyD"` is in `(self.reasoning or "")` (case-sensitive), return `ArticleType.CIENTIFICO`.
3. Otherwise return `ArticleType.DIVULGACION`.

#### Scenario: Scientific article with IMRyD reasoning

- GIVEN a `ClassificationResultDTO` with `article_type = ArticleType.CIENTIFICO` and `reasoning` containing `"IMRyD"`
- WHEN `effective_structure_type` is called
- THEN it returns `ArticleType.CIENTIFICO`

#### Scenario: Scientific article without IMRyD reasoning

- GIVEN a `ClassificationResultDTO` with `article_type = ArticleType.CIENTIFICO` and `reasoning` not containing `"IMRyD"`
- WHEN `effective_structure_type` is called
- THEN it returns `ArticleType.DIVULGACION`

#### Scenario: Non-scientific article type

- GIVEN a `ClassificationResultDTO` with `article_type = ArticleType.DIVULGACION`
- WHEN `effective_structure_type` is called
- THEN it returns `ArticleType.DIVULGACION` unchanged

---

### Requirement: RecommendationSettings DTO (no defaults)

`RecommendationSettings` MUST be a frozen DTO inheriting from `BaseDTO` in `src/domain/recommendation/recommendation_settings.py`. All fields are **required** (no default values). Default values are the responsibility of the infrastructure config layer (`RecommendationConfig`):

- `publish_threshold: float`
- `quality_threshold: float`
- `grammar_threshold: float`
- `dimension_threshold: float`
- `citation_match_threshold: float`
- `critical_citation_match_threshold: float`
- `citation_count_threshold: int`
- `classification_confidence_threshold: float`

> **Rationale**: Defaults in domain DTOs introduce hidden dependencies on infrastructure decisions. Defaults belong in `RecommendationConfig` (infrastructure config layer) and are resolved at wiring time.

---

### Requirement: RecommendationConfig Infrastructure Config Class

`RecommendationConfig` MUST reside in `src/infrastructure/config/recommendation_config.py`. It MUST expose a classmethod `build_settings() -> RecommendationSettings` that reads environment variables **at call time** (not at import time) with the following defaults:

| Env var | Default |
|---------|---------|
| `RECOMMENDATION_PUBLISH_THRESHOLD` | `7.0` |
| `RECOMMENDATION_QUALITY_THRESHOLD` | `7.0` |
| `RECOMMENDATION_GRAMMAR_THRESHOLD` | `7.0` |
| `RECOMMENDATION_DIMENSION_THRESHOLD` | `6.0` |
| `RECOMMENDATION_CITATION_MATCH_THRESHOLD` | `90.0` |
| `RECOMMENDATION_CRITICAL_CITATION_MATCH_THRESHOLD` | `50.0` |
| `RECOMMENDATION_CITATION_COUNT_THRESHOLD` | `10` |
| `RECOMMENDATION_CLASSIFICATION_CONFIDENCE_THRESHOLD` | `0.7` |

> **Rationale**: Reading `os.getenv` at class-attribute definition time (module import) makes env var overrides in tests impossible without module reloading. A classmethod reads env at call time, making it trivially patchable with `patch.dict(os.environ, ...)`.

#### Scenario: Environment variable overrides default

- GIVEN `RECOMMENDATION_QUALITY_THRESHOLD=6.5` is set in the environment
- WHEN `RecommendationConfig.build_settings()` is called
- THEN the returned `RecommendationSettings.quality_threshold` equals `6.5`

---

### Requirement: AnalysisContext Value Object

`AnalysisContext` MUST reside in `src/domain/recommendation/analysis_context.py` as a frozen dataclass grouping all analysis inputs and settings required by recommendation rules:

- `classification: ClassificationResultDTO`
- `quality: QualityResultDTO`
- `structure: StructureValidationResultDTO`
- `citations: CitationAnalysisResultDTO`
- `apa_validation: ApaValidationResultDTO`
- `grammar: GrammarCheckResultDTO`
- `settings: RecommendationSettings`

It MUST expose a computed property `citation_match_rate: float` returning `100.0` when `total_citations == 0`, otherwise `matched_count / total_citations * 100.0`.

---

### Requirement: RecommendationRule Abstract Base

`RecommendationRule` MUST reside in `src/domain/recommendation/recommendation_rule.py` as an abstract base class with a single abstract method:

```python
def evaluate(self, context: AnalysisContext) -> list[RecommendationDTO]: ...
```

---

### Requirement: Concrete Recommendation Rules

Seven concrete rule classes MUST reside in `src/domain/recommendation/rules.py`, each implementing `RecommendationRule`:

1. **`QualityRule`** — `HIGH` when `quality.overall_score < settings.quality_threshold`
2. **`GrammarRule`** — `HIGH` when `grammar.score < settings.grammar_threshold`
3. **`DimensionRule`** — `MEDIUM` for each dimension score below `settings.dimension_threshold`
4. **`StructureRule`** — `HIGH` for each missing section when `structure.is_valid is False`
5. **`CitationMatchRule`** — `HIGH` when match rate below `citation_match_threshold`; `MEDIUM` when unmatched count > 0 but above threshold
6. **`CitationCountRule`** — `MEDIUM` when `total_citations < settings.citation_count_threshold`
7. **`ConfidenceRule`** — `LOW` when `classification.confidence` is not `None` and below `settings.classification_confidence_threshold`

---

### Requirement: PublicationVerdictEvaluator Domain Service

`PublicationVerdictEvaluator` MUST reside in `src/domain/recommendation/publication_verdict_evaluator.py`. Its method `evaluate(context: AnalysisContext) -> PublicationVerdictDTO` determines the final publication verdict:

- **`CRITICAL`** if any of: `quality < 5.0`, `grammar < 5.0`, `structure.is_valid is False`, `match_rate < critical_citation_match_threshold`, OR `total_citations == 0`
- **`WARNING`** if any of: `quality < publish_threshold`, `grammar < publish_threshold`, `match_rate < citation_match_threshold`, `len(apa_validation.violations) > 0`
- **`APPROVED`** otherwise

---

### Requirement: RecommendationBuilder Domain Service (Rule Pattern)

`RecommendationBuilder` MUST reside in `src/domain/recommendation/recommendation_builder.py`. Its constructor accepts:
- `settings: RecommendationSettings`
- `rules: list[RecommendationRule] | None` (defaults to the 7 concrete rules)
- `verdict_evaluator: PublicationVerdictEvaluator | None` (defaults to a new instance)

Its method `build(...) -> tuple[list[RecommendationDTO], PublicationVerdictDTO]` MUST:
1. Build an `AnalysisContext` from the inputs and settings.
2. Call `rule.evaluate(context)` for each rule and concatenate results.
3. Call `verdict_evaluator.evaluate(context)` to get the verdict.
4. Return `(recommendations, verdict)`.

> **Rationale**: The Rule pattern makes each recommendation rule independently testable and eliminates long `if/elif` chains. The verdict is a distinct return value (not a recommendation) to reflect its semantic difference.

#### Scenario: All thresholds satisfied — APPROVED verdict, empty recommendations

- GIVEN `RecommendationBuilder` with default rules and all analysis inputs satisfying thresholds
- WHEN `build()` is executed
- THEN it returns `([], PublicationVerdictDTO(verdict=PublicationVerdict.APPROVED, ...))`

#### Scenario: Quality below threshold — HIGH recommendation + WARNING verdict

- GIVEN a quality score of `6.5` (above critical but below `quality_threshold=7.0`)
- WHEN `build()` is executed
- THEN the list contains a `HIGH` recommendation for quality AND the verdict is `WARNING`

#### Scenario: Critical quality issue — CRITICAL verdict

- GIVEN a quality score of `4.5` (below critical threshold `5.0`)
- WHEN `build()` is executed
- THEN the verdict is `CRITICAL`

#### Scenario: Zero citations — CRITICAL verdict

- GIVEN `total_citations == 0`
- WHEN `build()` is executed
- THEN the verdict is `CRITICAL` with a message about missing APA citations

---

### Requirement: AnalyzeDocumentUseCase Orchestrator

`AnalyzeDocumentUseCase` MUST live in `src/application/analyze_document_use_case.py` and coordinate the document analysis steps. It accepts its 11 dependencies via constructor injection. Method `execute(document_path: str) -> ReportInputDTO` MUST be wrapped with `@generic_error_handler` and perform:

1. Load paragraphs via `read_document_use_case`.
2. Extract content via `extract_content_use_case`.
3. Extract citations via `extract_citations_use_case`.
4. Filter `AUTHOR_YEAR` citations and pass tuples `(text, location, paragraph_text)` to `validate_apa_use_case`.
5. Grammar check via `check_grammar_use_case`.
6. Classify article via `classify_article_use_case`.
7. Analyze quality via `analyze_quality_use_case` using `classification.article_type`.
8. Validate structure using `classification.effective_structure_type` and `has_references = len(references) > 0`.
9. Parse `section_type` to `SectionName`; fallback to `SectionName.REFERENCES` on `ValueError`.
10. Match citations via `match_citations_use_case`.
11. Verify EUMIC via `verify_eumic_use_case` (no fatal exception on violations).
12. Call `recommendation_builder.build()` and **unpack the tuple**: `recommendations, verdict = builder.build(...)`.
13. Construct and return `ReportInputDTO` with all results including `verdict`.

#### Scenario: Orchestrator executes all pipeline steps sequentially

- GIVEN a valid `document_path`
- WHEN `execute(document_path)` is called
- THEN each sub-use case is called exactly once and a `ReportInputDTO` is returned

#### Scenario: Only AUTHOR_YEAR citations are validated

- GIVEN a document with 2 `AUTHOR_YEAR` and 1 `NUMERIC` citations
- WHEN the orchestrator runs
- THEN only the 2 `AUTHOR_YEAR` citations are sent to `ValidateApaUseCase`

#### Scenario: Structure validation uses effective structure type

- GIVEN a `CIENTIFICO` article without `"IMRyD"` in reasoning
- WHEN structure validation is invoked
- THEN `ValidateStructureUseCase` receives `ArticleType.DIVULGACION`

---

### Requirement: AnalyzeDocumentUseCaseWiring Assembly Factory

`AnalyzeDocumentUseCaseWiring` MUST reside in `src/infrastructure/wirings/analyze_document_use_case_wiring.py`. It MUST follow the **private-method wiring pattern** established in other wirings (e.g. `AnalyzeQualityUseCaseWiring`):

- `create_use_case()` delegates each dependency to a `_get_xxx()` private method.
- Each `_get_xxx()` method instantiates the corresponding sub-wiring or builds the object directly.
- `_get_recommendation_builder()` calls `RecommendationConfig.build_settings()` to obtain settings.

#### Scenario: Wiring constructs correct dependency graph

- GIVEN the wiring configuration
- WHEN `AnalyzeDocumentUseCaseWiring().create_use_case()` is called
- THEN it returns a valid `AnalyzeDocumentUseCase` with all 11 dependencies injected

#### Scenario: Environment variable overrides threshold at wiring time

- GIVEN `RECOMMENDATION_QUALITY_THRESHOLD=6.5` is set before calling `create_use_case()`
- WHEN `create_use_case()` is called
- THEN `recommendation_builder._settings.quality_threshold` equals `6.5`
