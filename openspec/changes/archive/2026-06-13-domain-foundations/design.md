# Technical Design: Domain Foundations (Slice 0)

> Slice 0 of the hexagonal migration (`docs/plan-migracion-hexagonal.md`).
> Normative guide: `.agent/skills/clean-architecture/SKILL.md`.
> Scope: PURE domain only — enums, entities, DTOs. No ports, adapters, use cases, wirings.

## 1. Architecture Approach

This slice is **additive population of the existing `src/domain/` skeleton**, not a
restructuring. The architectural pattern is fixed by the SKILL and the master plan:
hexagonal/clean architecture with a pure inner domain ring. Slice 0 fills only the
innermost ring (data types) and adds nothing that performs I/O.

One base class is used by all migrated types:

- `src/domain/dtos/base_dto.py` — `BaseDTO` is `@dataclass(frozen=True, eq=True)`,
  exposes `as_dict()` and `from_dict()`. All 9 migrated data types extend `BaseDTO`.

> `src/domain/entities/base_entity.py` (`BaseEntity`) exists in the skeleton but is
> **not used in this project** — see ADR-1 below.

**Core design decision — all data types are DTOs:**

The project has **no database**. Following the same reasoning as the "no DB → no
Repository / all Port" decision, there is no persistence identity layer. Domain data
simply flows through use cases as immutable records. Therefore:

- ALL 9 migrated data types (`Citation`, `Reference`, `Section`, `DocumentContent`,
  `ClassificationResult`, `QualityResult`, `StructureValidationResult`,
  `CitationAnalysisResult`, `AnalysisResult`) extend `BaseDTO` (`frozen=True`).
- `Section.__post_init__` retains ONLY the empty-title `ValueError` guard (raising
  is allowed in frozen `__post_init__`; only `self.*` assignment is forbidden).
- `DocumentContent.word_count` is a plain required field; paragraph-based
  auto-computation is deferred to a factory in a later slice.

**Boundaries respected:**
- Domain imports only stdlib (`from x import y`) and other `src.domain` modules.
- Absolute imports rooted at `src.domain` — no relative, no local/in-function imports.
- No `src.application` / `src.infrastructure` imports (they do not exist yet anyway).
- Legacy `domain/` is never imported from `src/`; full coexistence.

## 2. Final File Layout (concrete target files)

One class per file; `snake_case` filename = `PascalCase` class. Enums live flat under
`enums/`. Entities live in their **entity folder** (`src/domain/<entity>/`) per SKILL §1
and plan §4.1. DTOs live flat under `dtos/` with the `_dto` filename suffix.

### Enums — `src/domain/enums/`
```
src/domain/enums/article_type.py             ArticleType
src/domain/enums/article_size.py             ArticleSize
src/domain/enums/citation_type.py            CitationType
src/domain/enums/classification_category.py  ClassificationCategory
src/domain/enums/quality_level.py            QualityLevel
src/domain/enums/section_type.py             SectionType
src/domain/enums/analysis_dimension.py       AnalysisDimension
src/domain/enums/validation_status.py        ValidationStatus
src/domain/enums/recommendation_priority.py  RecommendationPriority
src/domain/enums/severity_level.py           SeverityLevel
```

### DTOs — `src/domain/dtos/` (ALL data types)
```
src/domain/dtos/citation_dto.py                     Citation               (BaseDTO, frozen)
src/domain/dtos/reference_dto.py                    Reference              (BaseDTO, frozen)
src/domain/dtos/section_dto.py                      Section                (BaseDTO, frozen)
src/domain/dtos/document_content_dto.py             DocumentContent        (BaseDTO, frozen)
src/domain/dtos/classification_result_dto.py        ClassificationResult   (BaseDTO, frozen)
src/domain/dtos/quality_result_dto.py               QualityResult          (BaseDTO, frozen)
src/domain/dtos/structure_validation_result_dto.py  StructureValidationResult (BaseDTO, frozen)
src/domain/dtos/citation_analysis_result_dto.py     CitationAnalysisResult (BaseDTO, frozen)
src/domain/dtos/analysis_result_dto.py              AnalysisResult         (BaseDTO, frozen)
```

> Class names keep their domain names (no `DTO` suffix on the class) to stay aligned with
> the master-plan catalog (`ClassificationResult`, `QualityResult`, …). The `_dto` token
> is a **filename** marker that the artifact lives in the DTO layer. This matches the plan
> §4.1 target paths exactly (`classification_result_dto.py`, etc.).

### Tests — `src/domain/tests/`
```
src/domain/tests/enums/__init__.py                       (new folder)
src/domain/tests/enums/test_article_type.py
src/domain/tests/enums/test_article_size.py
src/domain/tests/enums/test_citation_type.py
src/domain/tests/enums/test_classification_category.py
src/domain/tests/enums/test_quality_level.py
src/domain/tests/enums/test_section_type.py
src/domain/tests/enums/test_analysis_dimension.py
src/domain/tests/enums/test_validation_status.py
src/domain/tests/enums/test_recommendation_priority.py
src/domain/tests/enums/test_severity_level.py
src/domain/tests/dtos/test_citation.py
src/domain/tests/dtos/test_reference.py
src/domain/tests/dtos/test_section.py
src/domain/tests/dtos/test_document_content.py
src/domain/tests/dtos/test_classification_result.py
src/domain/tests/dtos/test_quality_result.py
src/domain/tests/dtos/test_structure_validation_result.py
src/domain/tests/dtos/test_citation_analysis_result.py
src/domain/tests/dtos/test_analysis_result.py
```
`tests/dtos/` already exists; `tests/enums/` is new (add `__init__.py`).
All data-type tests live under `tests/dtos/` — all 9 types are DTOs.

## 3. Per-Type DTO Classification (applied)

All types are frozen DTOs. The project has no database → no entities.

| Type | File | Rationale |
|---|---|---|
| `Citation` | `src/domain/dtos/citation_dto.py` | No DB → no entity; frozen record; `__str__` is a display helper, not mutation |
| `Reference` | `src/domain/dtos/reference_dto.py` | No DB → no entity; frozen record |
| `Section` | `src/domain/dtos/section_dto.py` | No DB → no entity; `__post_init__` only raises (compatible with frozen); no `get_word_count`/`is_empty` (YAGNI — no callers) |
| `DocumentContent` | `src/domain/dtos/document_content_dto.py` | No DB → no entity; `word_count` plain required field; paragraph-compute deferred to factory |
| `ClassificationResult` | `src/domain/dtos/classification_result_dto.py` | Immutable output of classification; crosses to formatter/exporter. Gets corrected `create()` factory (§4). |
| `QualityResult` | `src/domain/dtos/quality_result_dto.py` | Immutable analysis output. Consolidation target (§5). |
| `StructureValidationResult` | `src/domain/dtos/structure_validation_result_dto.py` | Immutable validation output. |
| `CitationAnalysisResult` | `src/domain/dtos/citation_analysis_result_dto.py` | Immutable analysis output. |
| `AnalysisResult` | `src/domain/dtos/analysis_result_dto.py` | Immutable aggregate; now contains only frozen DTOs. |

### 3.1 `DocumentContent` — frozen and `word_count` field
The legacy `__post_init__` computes `word_count` from `paragraphs` when `word_count == 0`.
That is **in-place mutation** (`self.word_count = ...`), incompatible with `frozen=True`.
**Decision:** `word_count` is a plain required `int` field. Callers supply it directly.
Paragraph-based computation is deferred to a factory in a later slice.
The `references` field is typed `list[Reference]` (absolute import
`from src.domain.dtos.reference_dto import Reference`). `keywords`/`paragraphs` →
`list[str]`, `sections` → `dict[str, str]`.

### 3.2 `Section` — `__post_init__` + out-of-scope helper + YAGNI
Legacy `Section.__post_init__` does two things:
1. Raises `ValueError` if `title` is empty — **kept** (raising in frozen `__post_init__`
   is valid; only `self.*` assignment is forbidden).
2. Auto-detects `section_type` via the **local import** of `classify_section_by_name`.
   That helper is **out of scope** (plan §4.2) and buggy.

**Decision:** `Section.__post_init__` keeps ONLY the empty-title guard. Auto-detection
removed; `section_type: SectionType | None = None` is taken as provided by the caller.
`get_word_count()` and `is_empty()` are **removed** (YAGNI — no callers exist in the
codebase). A test asserts `section_type` stays `None` when not passed.

> This is the only **intentional behavior change** beyond bug fixes. Safe because no
> caller is rewired in Slice 0; the new `src.domain.dtos.section_dto.Section` is not
> yet imported by any production code.

## 4. `ClassificationResult.create(...)` Factory Design

User decision: add a **corrected `@classmethod` factory on the DTO**; leave the broken
legacy `create_classification_result` helper untouched in legacy code (it is out of scope
and lives in `domain/models.py`).

The legacy helper is broken (bug #3): it passes `category=` and only 3 fields, but the
dataclass requires `article_type` + `article_size`. The corrected factory uses the real
field names:

```python
from dataclasses import dataclass, field
from datetime import datetime

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.article_type import ArticleType
from src.domain.enums.article_size import ArticleSize


@dataclass(frozen=True)
class ClassificationResult(BaseDTO):
    """Immutable result of article classification."""
    article_type: ArticleType
    article_size: ArticleSize
    confidence: float | None
    reasoning: str
    timestamp: datetime = field(default_factory=datetime.now)

    @classmethod
    def create(
        cls,
        article_type: ArticleType,
        article_size: ArticleSize,
        confidence: float | None,
        reasoning: str,
    ) -> "ClassificationResult":
        """Build a classification result with the correct domain field names."""
        return cls(
            article_type=article_type,
            article_size=article_size,
            confidence=confidence,
            reasoning=reasoning,
        )

    def __str__(self) -> str:
        confidence_text = f"{self.confidence:.1%}" if self.confidence is not None else "—"
        return (
            f"Classification: {self.article_type.value} | "
            f"Size: {self.article_size.value} | "
            f"Confidence: {confidence_text}"
        )
```

Notes:
- `@dataclass(frozen=True)` is re-declared on the subclass; inheriting a frozen base does
  not automatically make the subclass dataclass frozen unless re-decorated. `BaseDTO` is
  already `frozen=True`; the subclass MUST also be `@dataclass(frozen=True)` so its own
  fields are frozen and `eq` is generated. Frozen + `default_factory` for `timestamp` is
  valid (default_factory runs during `__init__`, before the instance is frozen).
- The factory is the **only** factory introduced in this slice. No other `create_*`
  legacy helper is migrated (the rest belong to later service slices, plan §4.2).
- A test asserts `ClassificationResult.create(...)` returns a valid instance with the
  correct field names and that the instance is frozen (assignment raises
  `FrozenInstanceError`).

## 5. In-Scope Bug Fixes (how each is corrected)

| # | Bug | Fix in this slice |
|---|---|---|
| 1 | `enums.py` `__all__` lists `SeverityLevel` before its definition; omits `ArticleType`/`ArticleSize`. | **Dissolved by construction.** One enum per file means there is no shared `__all__`. Each `src/domain/enums/<enum>.py` contains exactly one `Enum`, no `__all__`. The ordering bug cannot exist. |
| 2 | `QualityResult` vs `QualityAnalysisResult` duplication. | **Consolidate into one DTO: `QualityResult`** (the richer shape — keeps `timestamp` and `__str__`). `QualityAnalysisResult` is NOT migrated. Fields: `overall_score: float`, `quality_level: QualityLevel`, `dimension_scores: dict[str, dict[str, Any]] = field(default_factory=dict)`, `timestamp: datetime = field(default_factory=datetime.now)`. A test asserts the consolidated type carries the richer fields. |
| 4 | `Dict[str, Dict[str, any]]` uses builtin function `any` instead of type `Any` (models.py lines 90, 102, 201). | Migrate as `dict[str, dict[str, Any]]` with `from typing import Any`. Affects `QualityResult.dimension_scores` and `StructureValidationResult.section_details`. A test asserts the resolved type hint is `Any`, not `any`. |
| 7 | Mixed/relative + local imports (`from .enums` vs `from domain.enums`; local imports in `Section.__post_init__` and `get_citation_type_from_pattern`). | All migrated types use absolute `from src.domain.enums.<enum> import <Enum>` at module top. No relative imports, no local/in-function imports. `Section`'s local import is removed entirely (§3.2). |

Bugs #5 and #6 live in out-of-scope helper functions and are NOT touched here (recorded
in the proposal for traceability; fixed in their service slices).

### Typing migration applied uniformly (plan §9 / SKILL §0)
- `Optional[X]` → `X | None` (e.g. `confidence: float | None`, `author: str | None`).
- `List[X]` → `list[X]`; `Dict[K, V]` → `dict[K, V]`.
- `from typing import Any` only (drop `List`, `Dict`, `Optional` imports).
- `datetime` imported as `from datetime import datetime`.

### `AnalysisResult.to_dict()` shape
`AnalysisResult` aggregates the other types (now all frozen DTOs). It extends `BaseDTO`,
so it inherits `as_dict()` (deep `asdict`). The legacy `to_dict()` produced a
**custom, flattened** shape (enum `.value` strings, ISO timestamp, selected fields)
consumed by the formatter/exporter/Gradio. **Decision:** preserve that exact custom
serialization as an explicit `to_dict()` method on the DTO (typed `-> dict[str, Any]`),
separate from the inherited `as_dict()`. Keeping `to_dict()` byte-compatible protects
the downstream contract flagged in plan §10.4 until the later orchestrator slice (13)
formalizes it.

**Critical key:** the classification sub-dict uses key `"category"` (legacy byte-compatible
key), NOT `"article_type"`. The `ClassificationResult` *field* is named `article_type`
(correct domain name), but `to_dict()` serializes it under `"category"` for downstream
compatibility. A test pins the `to_dict()` shape (keys + `"category"` key + enum `.value`
flattening + ISO timestamp).

## 6. Test Strategy (Strict TDD)

- **Framework:** `unittest.TestCase`, MANDATORY (SKILL §6). No bare pytest functions.
- **Runner:** `python -m pytest src/` (pytest discovers `unittest.TestCase`).
- **Order:** failing-first. For each type: write the test (red) → migrate the type
  (green) → next type. One type per red/green cycle.
- **Location:** `src/domain/tests/{enums,entities,dtos}/test_<class>.py`. Add
  `src/domain/tests/enums/__init__.py`.
- **Per-type coverage (minimum):**
  - **Enums:** member set + each `.value`. `SeverityLevel` existence test directly
    documents bug #1 is gone (it imports cleanly with no `__all__`).
  - **DTOs (all 9 types):** construction; frozen (assignment raises `FrozenInstanceError`);
    `as_dict()`/`from_dict()` round-trip where applicable; `Section` empty-title raises
    `ValueError`, `section_type=None` when not provided (§3.2 delta);
    `DocumentContent` accepts `word_count` as required field (no auto-compute);
    `ClassificationResult.create()` factory (§4); `QualityResult` consolidation (bug #2);
    `Any` typing (bug #4); `AnalysisResult.to_dict()` shape with `"category"` key (§5).
- **Purity:** no DB, no network, no `python-docx`/`win32com`/`ollama`. Pure Python only.
- **No docstrings/comments inside test bodies** (SKILL §6); English self-documenting
  method names (e.g. `test_document_content_computes_word_count_from_paragraphs`).

## 7. Coexistence Guarantee

- No legacy file (`domain/enums.py`, `domain/models.py`, or anything outside `src/`) is
  read, modified, or deleted. Slice 0 is purely additive under `src/domain/`.
- Nothing in `src/` imports from legacy `domain/`. All cross-references inside `src/` use
  absolute `src.domain.*` imports (e.g. `DocumentContent.references: list[Reference]`
  imports `src.domain.reference.reference`).
- The legacy pytest suite under `tests/` must still pass unchanged after the slice.
- Rollback is deletion of the new `src/domain/` modules + tests; legacy behavior is
  untouched.

## 8. ADR-Style Decisions

**ADR-1: No entities — all 9 data types are frozen DTOs.**
- Decision: `Citation`, `Reference`, `Section`, `DocumentContent` extend `BaseDTO`
  (frozen), not `BaseEntity`. `BaseEntity` is unused in this project.
- Rationale: the project has no database → no persistence → no identity layer. This
  parallels the "no DB → no Repository / all Port" decision already in effect. Domain data
  flows through use cases as immutable records; there is no lifecycle to track.
- Rejected: using `BaseEntity` for types with behavior — behavior is a `__str__` display
  helper (Citation, Reference) or a raising guard (Section), neither of which requires
  mutability. Frozen DTOs support both patterns.

**ADR-2: Class names keep domain names; `_dto` is a filename marker only.**
- Decision: `ClassificationResult` (not `ClassificationResultDTO`) in file
  `classification_result_dto.py`.
- Rationale: matches plan §4.1 target paths and the use-case catalog (§6) verbatim,
  avoiding churn when later slices reference these names.
- Rejected: appending `DTO` to class names — SKILL §4 lists the suffix convention but the
  plan locked these specific names; consistency with the plan wins for migrated outputs.

**ADR-3: `Section` drops `section_type` auto-detection AND `get_word_count`/`is_empty`.**
- Decision: `__post_init__` keeps only the empty-title guard; auto-detection and behavior
  methods removed.
- Rationale: classifier helper is out of scope (plan §4.2) and buggy. `get_word_count()`
  and `is_empty()` have no callers — YAGNI. Frozen DTO `__post_init__` supports raising.
- Rejected: copying `classify_section_by_name` into the DTO (out of scope, pulls a later
  slice forward) or keeping dead behavior methods (YAGNI violation).

**ADR-4: `ClassificationResult.create()` corrected factory on the DTO; legacy helper
untouched.**
- Decision: per user — add corrected `@classmethod create()`; leave legacy
  `create_classification_result` broken in legacy code.
- Rationale: legacy is frozen for coexistence; the migrated type exposes a correct factory
  using real field names (`article_type`/`article_size`).
- Rejected: fixing the legacy helper (out of scope, would touch legacy) or porting it as a
  standalone function (violates POO/one-class-per-file).

**ADR-5: `QualityResult` is the single consolidated quality DTO.**
- Decision: migrate the richer `QualityResult`; drop `QualityAnalysisResult`.
- Rationale: `QualityResult` is a superset (adds `timestamp`, `__str__`); one concept,
  one type.
- Rejected: keeping both (duplication) or merging into a new name (needless churn for
  downstream references already targeting `QualityResult`).

**ADR-6: Preserve `AnalysisResult.to_dict()` custom shape verbatim; use `"category"` key.**
- Decision: keep the flattened `to_dict()` as an explicit method alongside inherited
  `as_dict()`. The classification sub-dict key is `"category"` (legacy byte-compatible).
- Rationale: formatter/exporter/Gradio depend on the exact flattened shape (plan §10.4);
  changing it now risks downstream breakage before the orchestrator slice.
- Rejected: replacing `to_dict()` with `as_dict()` — different shape, would break the
  downstream contract; using `"article_type"` as key — incompatible with legacy consumers.

## 9. Risks

| Risk | Likelihood | Mitigation |
|---|---|---|
| `Section` frozen `__post_init__` misunderstood as forbidden | Low | Documented ADR-3; raising in `__post_init__` is valid with `frozen=True`; covered by test. |
| `DocumentContent.word_count` callers expect auto-compute | Low | No caller yet (new type); factory deferred to later slice; documented in ADR-1 and spec. |
| `AnalysisResult` now all-DTO composition breaks assumption | Resolved | All sub-fields are now frozen DTOs; cleaner composition than before. |
| `AnalysisResult.to_dict()` `"category"` key surprises readers | Low | Documented ADR-6 + test pins the key explicitly; comment in `to_dict()` explains legacy reason. |
| `Section` `get_word_count`/`is_empty` removal surprises a future slice | Low | Documented ADR-3 + YAGNI rationale; no callers existed. |
| Frozen DTO + `default_factory(timestamp)` misunderstanding | Low | Valid Python; covered by a construction test. |
| Slice exceeds 400-line budget (19 types + 19 tests) | Med | Measured at tasks; split into chained PRs at tasks/apply per delivery strategy if over budget — NOT split in design. |
| Accidental legacy import from `src/` | Low | Invariant enforced by review; all intra-`src` refs use absolute `src.domain.*` paths. |
