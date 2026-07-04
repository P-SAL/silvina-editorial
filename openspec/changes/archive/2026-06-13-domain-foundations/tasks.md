# Task Checklist: Domain Foundations (Slice 0)

> Generated from: spec `sdd/domain-foundations/spec` + design `sdd/domain-foundations/design`
> Runner: `python -m pytest src/`
> TDD order: failing-first per type (red → green per work unit)
> Coexistence rule: no legacy file is modified; all tasks are purely additive under `src/domain/`.

---

## Review Workload Forecast

| Metric | Estimate |
|---|---|
| Production files created | 19 (10 enums + 9 DTOs) + 0 `__init__.py` for entity folders (no entity folders) |
| Test files created | 19 (10 enum tests + 9 DTO tests) |
| Estimated lines — prod (avg ~25 lines/enum, ~55/DTO) | ~250 + ~495 = ~745 |
| Estimated lines — tests (avg ~30 lines/enum test, ~70/DTO test) | ~300 + ~630 = ~930 |
| **Total estimated changed lines** | **~1 690** |
| 400-line budget risk | **High** |
| Chained PRs recommended | **Yes** |
| Decision needed before apply | **Yes** |

### Proposed PR slices (stacked to main — each independently mergeable)

| PR | Scope | Prod lines | Test lines | Slice total | Depends on |
|---|---|---|---|---|---|
| PR 1 | 10 enums + `src/domain/tests/enums/__init__.py` | ~250 | ~300 | ~550 | nothing (no src/ deps) |
| PR 2 | 4 base DTOs (`Citation`, `Reference`, `Section`, `DocumentContent`) | ~220 | ~280 | ~500 | PR 1 (DTOs import enums) |
| PR 3 | 5 result DTOs (`ClassificationResult`, `QualityResult`, `StructureValidationResult`, `CitationAnalysisResult`, `AnalysisResult`) | ~275 | ~350 | ~625 | PR 1 + PR 2 (`AnalysisResult` refs `DocumentContent`) |

> PR 3 alone is ~625 lines. If the team requires strict ≤400/PR, split it further: PR 3a = first 4 result DTOs, PR 3b = `AnalysisResult`. Confirm with orchestrator before apply.

---

## Prerequisites (already in place — verify before starting)

- [x] `src/domain/entities/base_entity.py` exists (`BaseEntity`, unused in this project)
- [x] `src/domain/dtos/base_dto.py` exists (`BaseDTO`, `frozen=True`)
- [x] `src/domain/tests/dtos/` exists with `__init__.py`
- [x] `src/domain/tests/enums/` with `__init__.py` — created in Task 0
- [ ] Verify no entity folders (`citation/`, `reference/`, `section/`, `document/`) under `src/domain/` (all types are DTOs)

---

## Task 0 — Scaffold: create missing `__init__.py` files

**Spec**: REQ-STRUCTURE-1, REQ-TEST-1
**Parallel**: No (must run first; everything depends on folder existence)

- [x] Create `src/domain/tests/enums/__init__.py` (empty)

> No entity folders needed — all types are DTOs under `src/domain/dtos/`.

**Verify**: `python -m pytest src/` passes (no new test yet, just no import errors)

---

## PR 1 — Enums (Tasks 1–10)

> Tasks 1–10 are **fully parallel** once Task 0 is done. Each enum has zero intra-enum dependencies.

### Task 1 — `ArticleType` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 2–10, after Task 0)

- [x] Write failing test `src/domain/tests/enums/test_article_type.py`
  - `test_members_and_values` — asserts `CIENTIFICO`, `DIVULGACION`, `OPINION`, `UNKNOWN` with exact Spanish values
  - `test_importable_independently` — asserts import succeeds in isolation
  - `test_no_domain_imports` — (static assertion documented; actual check via grep in CI)
- [x] Create `src/domain/enums/article_type.py` with `ArticleType(Enum)` (only `from enum import Enum`)
- [x] Run `python -m pytest src/domain/tests/enums/test_article_type.py` — green

**Work unit commit**: `feat(domain/enums): add ArticleType enum and tests`

---

### Task 2 — `ArticleSize` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 1, 3–10)

- [x] Write failing test `src/domain/tests/enums/test_article_size.py`
  - `test_members_and_values` — asserts `LARGO`, `CORTO`, `NO_DEFINIDO`, `FUERA_RANGO` with exact values
- [x] Create `src/domain/enums/article_size.py` with `ArticleSize(Enum)`
- [x] Run `python -m pytest src/domain/tests/enums/test_article_size.py` — green

**Work unit commit**: `feat(domain/enums): add ArticleSize enum and tests`

---

### Task 3 — `CitationType` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 1–2, 4–10)

- [x] Write failing test `src/domain/tests/enums/test_citation_type.py`
  - `test_members_and_values` — asserts `AUTHOR_YEAR`, `NUMERIC`, `FOOTNOTE`, `UNKNOWN` with exact values
- [x] Create `src/domain/enums/citation_type.py` with `CitationType(Enum)`
- [x] Run `python -m pytest src/domain/tests/enums/test_citation_type.py` — green

**Work unit commit**: `feat(domain/enums): add CitationType enum and tests`

---

### Task 4 — `ClassificationCategory` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 1–3, 5–10)

- [x] Write failing test `src/domain/tests/enums/test_classification_category.py`
  - `test_members_and_values` — asserts `RESEARCH_ARTICLE`, `REVIEW_ARTICLE`, `REFLECTION_ARTICLE`, `SHORT_ARTICLE`, `CASE_REPORT`, `UNKNOWN`
- [x] Create `src/domain/enums/classification_category.py` with `ClassificationCategory(Enum)`
- [x] Run `python -m pytest src/domain/tests/enums/test_classification_category.py` — green

**Work unit commit**: `feat(domain/enums): add ClassificationCategory enum and tests`

---

### Task 5 — `QualityLevel` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 1–4, 6–10)

- [x] Write failing test `src/domain/tests/enums/test_quality_level.py`
  - `test_members_and_values` — asserts `EXCELLENT="Excelente"`, `GOOD="Bueno"`, `ACCEPTABLE="Aceptable"`, `NEEDS_IMPROVEMENT="Requiere mejoras"`, `POOR="Deficiente"`
- [x] Create `src/domain/enums/quality_level.py` with `QualityLevel(Enum)`
- [x] Run `python -m pytest src/domain/tests/enums/test_quality_level.py` — green

**Work unit commit**: `feat(domain/enums): add QualityLevel enum and tests`

---

### Task 6 — `SectionType` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 1–5, 7–10)

- [x] Write failing test `src/domain/tests/enums/test_section_type.py`
  - `test_member_count_is_23` — asserts `len(SectionType) == 23` (23 members, not 22; APPENDIX+ANEXO both present — verified against legacy)
  - `test_all_bilingual_section_names_present` — spot-checks `TITLE`, `RESUMEN`, `ABSTRACT`, `INTRODUCCION`, `INTRODUCTION`, `METODOLOGIA`, `METHODOLOGY`, `CONCLUSIONES`, `CONCLUSIONS`, `REFERENCIAS`, `REFERENCES`
- [x] Create `src/domain/enums/section_type.py` with `SectionType(Enum)` (all 23 members verbatim from legacy)
- [x] Run `python -m pytest src/domain/tests/enums/test_section_type.py` — green

**Work unit commit**: `feat(domain/enums): add SectionType enum and tests`

---

### Task 7 — `AnalysisDimension` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 1–6, 8–10)

- [x] Write failing test `src/domain/tests/enums/test_analysis_dimension.py`
  - `test_members_and_values` — asserts all 8 members (`ACADEMIC_RIGOR`, `METHODOLOGICAL_CLARITY`, `ARGUMENTATION`, `LITERATURE_REVIEW`, `ORIGINALITY`, `WRITING_QUALITY`, `STRUCTURE`, `CITATION_QUALITY`)
- [x] Create `src/domain/enums/analysis_dimension.py` with `AnalysisDimension(Enum)`
- [x] Run `python -m pytest src/domain/tests/enums/test_analysis_dimension.py` — green

**Work unit commit**: `feat(domain/enums): add AnalysisDimension enum and tests`

---

### Task 8 — `ValidationStatus` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 1–7, 9–10)

- [x] Write failing test `src/domain/tests/enums/test_validation_status.py`
  - `test_members_and_values` — asserts `PASSED`, `FAILED`, `WARNING`, `NOT_APPLICABLE`
- [x] Create `src/domain/enums/validation_status.py` with `ValidationStatus(Enum)`
- [x] Run `python -m pytest src/domain/tests/enums/test_validation_status.py` — green

**Work unit commit**: `feat(domain/enums): add ValidationStatus enum and tests`

---

### Task 9 — `RecommendationPriority` enum + test

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3
**Parallel**: Yes (with Tasks 1–8, 10)

- [x] Write failing test `src/domain/tests/enums/test_recommendation_priority.py`
  - `test_members_and_values` — asserts `HIGH="alta"`, `MEDIUM="media"`, `LOW="baja"`
- [x] Run failing first — create `src/domain/enums/recommendation_priority.py`
- [x] Run `python -m pytest src/domain/tests/enums/test_recommendation_priority.py` — green

**Work unit commit**: `feat(domain/enums): add RecommendationPriority enum and tests`

---

### Task 10 — `SeverityLevel` enum + test (bug #1 regression guard)

**Spec**: REQ-ENUM-1, REQ-ENUM-2, REQ-ENUM-3, REQ-TEST-3 (bug #1)
**Parallel**: Yes (with Tasks 1–9)

- [x] Write failing test `src/domain/tests/enums/test_severity_level.py`
  - `test_severity_level_importable_independently` — imports `SeverityLevel` in isolation (no other enum imported first); asserts `SeverityLevel.CRITICAL.value == "critical"` (direct regression guard for bug #1)
  - `test_members_and_values` — asserts `INFO`, `WARNING`, `ERROR`, `CRITICAL` with exact values
- [x] Create `src/domain/enums/severity_level.py` with `SeverityLevel(Enum)`
  - Only `from enum import Enum` import; no `__all__`; class defined at top of file
- [x] Run `python -m pytest src/domain/tests/enums/test_severity_level.py` — green

**Work unit commit**: `feat(domain/enums): add SeverityLevel enum and tests (bug-1 fixed)`

---

### PR 1 Integration Task

- [x] Run `python -m pytest src/domain/tests/enums/` — all 10 enum test files green
- [ ] Run `python -m pytest tests/` — legacy suite still passes (coexistence REQ-COEXISTENCE-1)
- [ ] Open PR 1 targeting `main` (or tracker branch per chosen chain strategy)

---

## PR 2 — Base DTOs (Tasks 11–14)

> Tasks 11–13 (`Citation`, `Reference`, `Section`) are parallel. Task 14 (`DocumentContent`)
> depends on Task 11 (`Reference`) because `DocumentContent.references: list[Reference]`.
> All types are frozen DTOs under `src/domain/dtos/`.

### Task 11 — `Citation` DTO + test

**Spec**: REQ-DTO-CITATION-1, REQ-DTO-CITATION-2, REQ-DTO-CITATION-3
**Parallel**: Yes (with Tasks 12, 13; must complete before Task 14 can start)

- [x] Write failing test `src/domain/tests/dtos/test_citation.py`
  - `test_citation_is_subclass_of_base_dto`
  - `test_citation_instantiation_with_required_fields_only` — `author` and `year` are `None`
  - `test_citation_is_immutable` — assignment raises `FrozenInstanceError`
  - `test_citation_as_dict_contains_expected_keys` — keys: `text`, `citation_type`, `location`, `author`, `year`
  - `test_citation_str_truncates_at_50_chars` — result starts with `"Citation("` and contains `"..."` for long text
  - `test_citation_type_hints_use_modern_syntax` — no `Optional[`, no `List[`, no `Dict[`
- [x] Create `src/domain/dtos/citation_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`; imports: `from dataclasses import dataclass`, `from src.domain.dtos.base_dto import BaseDTO`, `from src.domain.enums.citation_type import CitationType`
  - Fields: `text: str`, `citation_type: CitationType`, `location: int`, `author: str | None = None`, `year: str | None = None`
  - `__str__` method returning `"Citation(<text[:50]>...)"` when text > 50 chars
- [x] Run `python -m pytest src/domain/tests/dtos/test_citation.py` — green

**Work unit commit**: `feat(domain/dtos): add Citation DTO and tests`

---

### Task 12 — `Reference` DTO + test

**Spec**: REQ-DTO-REFERENCE-1, REQ-DTO-REFERENCE-2
**Parallel**: Yes (with Tasks 11, 13; must complete before Task 14)

- [x] Write failing test `src/domain/tests/dtos/test_reference.py`
  - `test_reference_is_subclass_of_base_dto`
  - `test_reference_instantiation_with_required_field_only` — all optional fields are `None`
  - `test_reference_is_immutable` — assignment raises `FrozenInstanceError`
  - `test_reference_str_returns_formatted_string` — `Reference(text="...", authors="Smith", year="2020")` → `"Reference(Smith, 2020)"`
  - `test_reference_str_when_authors_and_year_are_none` — no crash; handles gracefully
- [x] Create `src/domain/dtos/reference_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`
  - Fields: `text: str`, `authors: str | None = None`, `year: str | None = None`, `title: str | None = None`, `source: str | None = None`
  - `__str__` returning `f"Reference({self.authors}, {self.year})"` (or equivalent graceful form)
- [x] Run `python -m pytest src/domain/tests/dtos/test_reference.py` — green

**Work unit commit**: `feat(domain/dtos): add Reference DTO and tests`

---

### Task 13 — `Section` DTO + test

**Spec**: REQ-DTO-SECTION-1, REQ-DTO-SECTION-2, REQ-DTO-SECTION-3, REQ-TEST-3 (bug #7)
**Parallel**: Yes (with Tasks 11, 12; no deps on them)

- [x] Write failing test `src/domain/tests/dtos/test_section.py`
  - `test_section_is_subclass_of_base_dto`
  - `test_section_is_immutable` — assignment raises `FrozenInstanceError`
  - `test_section_with_empty_title_raises_value_error`
  - `test_section_without_section_type_has_section_type_none` — bug #7 regression guard: no local import crash, `section_type is None`
  - `test_section_with_explicit_section_type_preserves_it`
- [x] Create `src/domain/dtos/section_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`
  - Imports at top only: `from dataclasses import dataclass`, `from src.domain.dtos.base_dto import BaseDTO`, `from src.domain.enums.section_type import SectionType`
  - Fields: `title: str`, `content: str`, `section_type: SectionType | None = None`, `start_position: int = 0`, `end_position: int = 0`, `level: int = 1`
  - `__post_init__`: only raises `ValueError` if `title` is empty; NO local imports; NO `classify_section_by_name` call; NO `get_word_count`/`is_empty` methods (YAGNI)
- [x] Run `python -m pytest src/domain/tests/dtos/test_section.py` — green

**Work unit commit**: `feat(domain/dtos): add Section DTO and tests (bug-7 fixed)`

---

### Task 14 — `DocumentContent` DTO + test

**Spec**: REQ-DTO-DOCCONTENT-1, REQ-DTO-DOCCONTENT-2
**Parallel**: No — depends on Task 12 (Reference) being green first
**Sequential after**: Task 12

- [x] Write failing test `src/domain/tests/dtos/test_document_content.py`
  - `test_document_content_is_subclass_of_base_dto`
  - `test_document_content_is_immutable` — assignment raises `FrozenInstanceError`
  - `test_document_content_accepts_word_count_as_required_field` — `word_count=42, char_count=100` → `doc.word_count == 42`
  - `test_document_content_field_types_use_modern_syntax` — no `List[`, `Dict[`, `Optional[`
  - `test_document_content_references_is_list_of_reference_dtos` — `references` field accepts `list[Reference]`
- [x] Create `src/domain/dtos/document_content_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`
  - Imports: `from dataclasses import dataclass, field`, `from src.domain.dtos.base_dto import BaseDTO`, `from src.domain.dtos.reference_dto import Reference`
  - Fields with legacy names: `word_count: int`, `char_count: int`, `paragraph_count: int = 0`, `title: str | None = None`, `authors: str | None = None`, `abstract: str | None = None`, `keywords: list[str] = field(default_factory=list)`, `references: list[Reference] = field(default_factory=list)`, `paragraphs: list[str] = field(default_factory=list)`, `sections: dict[str, str] = field(default_factory=dict)`
  - NO `__post_init__` that assigns to `self.*` (frozen); `word_count` is caller-supplied
- [x] Run `python -m pytest src/domain/tests/dtos/test_document_content.py` — green

**Work unit commit**: `feat(domain/dtos): add DocumentContent DTO and tests`

---

### PR 2 Integration Task

- [x] Run `python -m pytest src/domain/tests/dtos/test_citation.py src/domain/tests/dtos/test_reference.py src/domain/tests/dtos/test_section.py src/domain/tests/dtos/test_document_content.py` — all green
- [x] Run `python -m pytest tests/` — legacy suite still passes (148 passed, 3 skipped)
- [ ] Open PR 2 targeting `main` (after PR 1 merged, or on stacked branch if using feature-branch-chain)

---

## PR 3 — DTOs (Tasks 15–19)

> Tasks 15–18 are **parallel** (they only import enums from PR 1). Task 19 (`AnalysisResult`) depends on Tasks 15–18 AND Task 14 (`DocumentContent`).

### Task 15 — `ClassificationResult` DTO + test (includes `create()` factory + bug #3)

**Spec**: REQ-DTO-CLASSIFICATION-1, REQ-DTO-CLASSIFICATION-2, REQ-DTO-CLASSIFICATION-3
**Parallel**: Yes (with Tasks 16, 17, 18; after PR 1 enums)

- [x] Write failing test `src/domain/tests/dtos/test_classification_result.py`
  - `test_classification_result_is_subclass_of_base_dto`
  - `test_classification_result_instantiation_with_correct_fields`
  - `test_classification_result_is_immutable` — assignment raises `FrozenInstanceError`
  - `test_create_factory_builds_valid_instance` — correct fields, `timestamp` auto-set
  - `test_create_factory_with_none_confidence`
  - `test_create_factory_result_is_frozen`
  - `test_str_contains_enum_values_and_confidence_percentage` — output contains `"científico"` and `"largo"` and `"%"`
- [x] Create `src/domain/dtos/classification_result_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`
  - Imports: `from dataclasses import dataclass, field`, `from datetime import datetime`, `from src.domain.dtos.base_dto import BaseDTO`, `from src.domain.enums.article_type import ArticleType`, `from src.domain.enums.article_size import ArticleSize`
  - Fields: `article_type: ArticleType`, `article_size: ArticleSize`, `confidence: float | None`, `reasoning: str`, `timestamp: datetime = field(default_factory=datetime.now)`
  - `@classmethod create(cls, article_type, article_size, confidence, reasoning) -> "ClassificationResult"`
  - `__str__` returning human-readable form with enum `.value` and confidence as percentage
- [x] Run `python -m pytest src/domain/tests/dtos/test_classification_result.py` — green

**Work unit commit**: `feat(domain/dtos): add ClassificationResult DTO with create() factory and tests`

---

### Task 16 — `QualityResult` DTO + test (bug #2 consolidation + bug #4 Any typing)

**Spec**: REQ-DTO-QUALITY-1, REQ-DTO-QUALITY-2, REQ-TEST-3 (bugs #2 and #4)
**Parallel**: Yes (with Tasks 15, 17, 18)

- [x] Write failing test `src/domain/tests/dtos/test_quality_result.py`
  - `test_quality_result_is_subclass_of_base_dto`
  - `test_quality_result_is_immutable`
  - `test_quality_result_str_returns_score_and_level` — `QualityResult(overall_score=8.5, quality_level=QualityLevel.GOOD)` → `"Quality: 8.5/10 (Bueno)"`
  - `test_quality_analysis_result_does_not_exist_in_src` — `from src.domain.dtos.quality_analysis_result_dto import QualityAnalysisResult` raises `ImportError` (bug #2 regression guard)
  - `test_dimension_scores_annotation_uses_typing_any` — inspect `get_type_hints(QualityResult)["dimension_scores"]` and assert it is `dict[str, dict[str, Any]]`, not using builtin `any` (bug #4 regression guard)
- [x] Create `src/domain/dtos/quality_result_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`
  - Imports include `from typing import Any`; no `Dict`, `Optional`
  - Fields: `overall_score: float`, `quality_level: QualityLevel`, `dimension_scores: dict[str, dict[str, Any]] = field(default_factory=dict)`, `timestamp: datetime = field(default_factory=datetime.now)`
  - `__str__` returning `f"Quality: {self.overall_score}/10 ({self.quality_level.value})"`
- [x] Run `python -m pytest src/domain/tests/dtos/test_quality_result.py` — green
- [x] Confirm no `quality_analysis_result_dto.py` file exists under `src/domain/dtos/`

**Work unit commit**: `feat(domain/dtos): add QualityResult DTO and tests (bug-2 and bug-4 fixed)`

---

### Task 17 — `StructureValidationResult` DTO + test (bug #4 Any typing)

**Spec**: REQ-DTO-STRUCTURE-1, REQ-TEST-3 (bug #4)
**Parallel**: Yes (with Tasks 15, 16, 18)

- [x] Write failing test `src/domain/tests/dtos/test_structure_validation_result.py`
  - `test_structure_validation_result_is_subclass_of_base_dto`
  - `test_structure_validation_result_is_immutable`
  - `test_str_for_valid_structure` — `StructureValidationResult(is_valid=True, missing_sections=[])` → `"Structure: Valid"`
  - `test_str_for_invalid_structure_with_two_missing` — → `"Structure: Invalid (2 missing)"`
  - `test_section_details_annotation_uses_typing_any` — bug #4 regression guard
- [x] Create `src/domain/dtos/structure_validation_result_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`
  - `from typing import Any`; fields: `is_valid: bool`, `missing_sections: list[str] = field(default_factory=list)`, `section_details: dict[str, dict[str, Any]] = field(default_factory=dict)`, `timestamp: datetime = field(default_factory=datetime.now)`
  - `__str__` branching on `is_valid`
- [x] Run `python -m pytest src/domain/tests/dtos/test_structure_validation_result.py` — green

**Work unit commit**: `feat(domain/dtos): add StructureValidationResult DTO and tests (bug-4 fixed)`

---

### Task 18 — `CitationAnalysisResult` DTO + test

**Spec**: REQ-DTO-CITATION-ANALYSIS-1
**Parallel**: Yes (with Tasks 15, 16, 17)

- [x] Write failing test `src/domain/tests/dtos/test_citation_analysis_result.py`
  - `test_citation_analysis_result_is_subclass_of_base_dto`
  - `test_citation_analysis_result_is_immutable`
  - `test_str_with_citations` — `total_citations=10, matched_count=8` → `"Citations: 10 (80.0% matched)"`
  - `test_str_with_zero_citations` — `total_citations=0` → `"Citations: 0 (0.0% matched)"` (no division by zero)
- [x] Create `src/domain/dtos/citation_analysis_result_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`
  - Fields: `total_citations: int`, `total_references: int`, `matched_count: int`, `unmatched_count: int`, `citations_by_type: dict[str, int] = field(default_factory=dict)`, `unmatched_citations: list[str] = field(default_factory=list)`, `timestamp: datetime = field(default_factory=datetime.now)`
  - `__str__` computing percentage with zero-safe guard (`self.total_citations or 1` denominator, but display `0.0%` when 0)
- [x] Run `python -m pytest src/domain/tests/dtos/test_citation_analysis_result.py` — green

**Work unit commit**: `feat(domain/dtos): add CitationAnalysisResult DTO and tests`

---

### Task 19 — `AnalysisResult` DTO + test (to_dict contract — `"category"` key)

**Spec**: REQ-DTO-ANALYSIS-1, REQ-DTO-ANALYSIS-2
**Parallel**: No — depends on Tasks 14 (`DocumentContent`), 15 (`ClassificationResult`), 16 (`QualityResult`), 17 (`StructureValidationResult`), 18 (`CitationAnalysisResult`)
**Sequential after**: Tasks 14 + 15 + 16 + 17 + 18

- [x] Write failing test `src/domain/tests/dtos/test_analysis_result.py`
  - `test_analysis_result_is_subclass_of_base_dto`
  - `test_analysis_result_is_immutable`
  - `test_to_dict_returns_all_required_top_level_keys` — keys: `filename`, `timestamp`, `classification`, `quality`, `structure`, `citations`
  - `test_to_dict_classification_uses_category_key` — `result["classification"]` contains `"category"` (NOT `"article_type"`) with enum `.value` string (legacy byte-compatible key per ADR-6)
  - `test_to_dict_timestamp_is_iso8601_string` — value is a string matching `.isoformat()` output
  - Uses a helper `_make_analysis_result()` factory inside the test class to build valid nested objects
- [x] Create `src/domain/dtos/analysis_result_dto.py`
  - `@dataclass(frozen=True)` subclass of `BaseDTO`
  - Imports: all DTO types from their absolute `src.domain.dtos.*` paths; `from typing import Any`; `from datetime import datetime`
  - Fields: `filename: str`, `document_content: DocumentContent` (frozen DTO), `classification: ClassificationResult`, `quality: QualityResult`, `structure: StructureValidationResult`, `citations: CitationAnalysisResult`, `timestamp: datetime = field(default_factory=datetime.now)`
  - `to_dict(self) -> dict[str, Any]`: custom flattened shape; classification sub-dict uses key `"category"` (legacy byte-compatible); uses enum `.value`, ISO timestamp, selected fields per spec contract
- [x] Run `python -m pytest src/domain/tests/dtos/test_analysis_result.py` — green

**Work unit commit**: `feat(domain/dtos): add AnalysisResult DTO with to_dict contract and tests`

---

### PR 3 Integration Task

- [x] Run `python -m pytest src/domain/tests/dtos/` — all 9 DTO tests green
- [x] Run `python -m pytest src/` — entire new test suite green (102 passed)
- [x] Run `python -m pytest tests/` — legacy suite still passes (148 passed, 3 skipped)
- [ ] Open PR 3 targeting `main` (or previous stacked branch)

---

## Cross-Cutting Verification Tasks (run once after all PRs merge)

**Spec**: REQ-IMPORTS-1, REQ-IMPORTS-2, REQ-IMPORTS-3, REQ-IMPORTS-4, REQ-STRUCTURE-2, REQ-COEXISTENCE-1

- [ ] Static scan: no `Optional[`, `List[`, `Dict[` in any file under `src/domain/` created by this slice
- [ ] Static scan: no import statements inside function/method bodies under `src/domain/` (no indented `import`/`from ... import`)
- [ ] Static scan: no `import *` in any migrated file
- [ ] Static scan: all cross-domain imports inside `src/` use `src.domain.*` (no `from domain.`, no `from .`)
- [ ] File ending: every `.py` file created ends with exactly one blank line
- [ ] Run `python -m pytest src/` → all tests pass
- [ ] Run `python -m pytest tests/` → all legacy tests pass

---

## Task Dependency Graph

```
Task 0 (scaffold)
    │
    ├── Tasks 1–10 (enums — all parallel)         ← PR 1
    │       │
    │       ├── Task 11 (Citation DTO)   ─────────┐
    │       ├── Task 12 (Reference DTO)  ─────────┤  ← PR 2
    │       └── Task 13 (Section DTO)    ─────────┤
    │               (all parallel)                 │
    │                                              ▼
    │                                  Task 14 (DocumentContent DTO)
    │                                              │
    │       ┌──── Task 15 (ClassificationResult)  ──────┐
    │       ├──── Task 16 (QualityResult)          ──────┤
    │       ├──── Task 17 (StructureValidationResult) ──┤  ← PR 3
    │       └──── Task 18 (CitationAnalysisResult)  ────┤
    │               (parallel among themselves)          │
    │                                                    ▼
    │                                         Task 19 (AnalysisResult)
    │
    └── Cross-cutting verification (after all PRs merge)
```

---

## Summary

| Phase | Tasks | Sequential? | Estimated lines |
|---|---|---|---|
| Scaffold | T0 | Sequential first | ~5 |
| Enums | T1–T10 | Parallel (post T0) | ~550 |
| Base DTOs | T11–T13 parallel, T14 after T12 | Mixed | ~500 |
| Result DTOs | T15–T18 parallel, T19 after T14+T15+T16+T17+T18 | Mixed | ~625 |
| **Total** | **20 tasks** | — | **~1 680** |
