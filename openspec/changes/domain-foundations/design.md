# Technical Design: Domain Foundations (Slice 0)

> Slice 0 of the hexagonal migration (`docs/plan-migracion-hexagonal.md`).
> Normative guide: `.agent/skills/clean-architecture/SKILL.md`.
> Scope: PURE domain only — enums, entities, DTOs. No ports, adapters, use cases, wirings.

## 1. Architecture Approach

This slice is **additive population of the existing `src/domain/` skeleton**, not a
restructuring. The architectural pattern is fixed by the SKILL and the master plan:
hexagonal/clean architecture with a pure inner domain ring. Slice 0 fills only the
innermost ring (data types) and adds nothing that performs I/O.

Two base classes already exist and are the contract every migrated type slots into:

- `src/domain/entities/base_entity.py` — `BaseEntity` exposes `as_dict()` via
  `dataclasses.asdict`. Subclasses MUST be `@dataclass`. **Mutable** (no `frozen`).
- `src/domain/dtos/base_dto.py` — `BaseDTO` is `@dataclass(frozen=True, eq=True)`,
  exposes `as_dict()` and `from_dict()`. Subclasses inherit `frozen=True`.

**Core layering decision (Entity vs DTO):** apply the SKILL criterion per type, not a
blanket plan suggestion.

- **Entity** = mutable + has behavior (`__post_init__` that mutates, instance methods).
  Extends `BaseEntity`, decorated `@dataclass` (NOT frozen).
- **DTO** = immutable result that crosses an outward boundary, no mutation.
  Extends `BaseDTO` (already `frozen=True`).

The decisive test for this slice: **can the type be frozen?** If a type mutates itself
in `__post_init__` or exposes mutating behavior, it CANNOT be a frozen `BaseDTO` and MUST
be an Entity. This directly resolves the two flagged cases (`DocumentContent`, `Section`).

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

### Entities — one folder per entity (`src/domain/<entity>/`)
```
src/domain/citation/citation.py              Citation            (BaseEntity, @dataclass)
src/domain/reference/reference.py            Reference           (BaseEntity, @dataclass)
src/domain/section/section.py                Section             (BaseEntity, @dataclass)
src/domain/document/document_content.py      DocumentContent     (BaseEntity, @dataclass)
```
Each entity folder gets an `__init__.py`.

### DTOs — `src/domain/dtos/`
```
src/domain/dtos/classification_result_dto.py        ClassificationResult   (BaseDTO)
src/domain/dtos/quality_result_dto.py               QualityResult          (BaseDTO)
src/domain/dtos/structure_validation_result_dto.py  StructureValidationResult (BaseDTO)
src/domain/dtos/citation_analysis_result_dto.py     CitationAnalysisResult (BaseDTO)
src/domain/dtos/analysis_result_dto.py              AnalysisResult         (BaseDTO)
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
src/domain/tests/entities/test_citation.py
src/domain/tests/entities/test_reference.py
src/domain/tests/entities/test_section.py
src/domain/tests/entities/test_document_content.py
src/domain/tests/dtos/test_classification_result.py
src/domain/tests/dtos/test_quality_result.py
src/domain/tests/dtos/test_structure_validation_result.py
src/domain/tests/dtos/test_citation_analysis_result.py
src/domain/tests/dtos/test_analysis_result.py
```
`tests/entities/` and `tests/dtos/` already exist; `tests/enums/` is new (add
`__init__.py`). Entity tests live under `tests/entities/` even though the source entity
lives in its own `src/domain/<entity>/` folder — the SKILL maps all `src/domain/` tests
under `src/domain/tests/<topic>/`, and `entities` is the topic for entity classes.

## 3. Per-Type Entity vs DTO Decision (applied)

| Type | Decision | Rationale (SKILL criterion) |
|---|---|---|
| `Citation` | **Entity** | Mutable dataclass, has `__str__` behavior, no boundary-crossing immutability requirement. |
| `Reference` | **Entity** | Same as `Citation`: mutable, `__str__` behavior. |
| `Section` | **Entity** | Has `__post_init__` that **raises** on empty title + behavior methods `get_word_count()`, `is_empty()`. Cannot be a frozen DTO. |
| `DocumentContent` | **Entity** | `__post_init__` **mutates** `self.word_count`. A frozen `BaseDTO` forbids attribute assignment after init → would raise `FrozenInstanceError`. MUST be a mutable Entity. |
| `ClassificationResult` | **DTO** | Immutable output of classification; crosses to formatter/exporter. Gets corrected `create()` factory (§4). |
| `QualityResult` | **DTO** | Immutable analysis output. Consolidation target (§5). |
| `StructureValidationResult` | **DTO** | Immutable validation output. |
| `CitationAnalysisResult` | **DTO** | Immutable analysis output. |
| `AnalysisResult` | **DTO** | Immutable aggregate output consumed downstream. |

### 3.1 `DocumentContent` — frozen conflict resolution
The legacy `__post_init__` computes `word_count` from `paragraphs` when `word_count == 0`.
This is **in-place mutation**, incompatible with `frozen=True`. Decision: keep
`DocumentContent` as a **mutable `BaseEntity`** and preserve the `__post_init__` mutation
verbatim (typed). No behavior change. The `references` field is typed
`list[Reference]` referencing the migrated `Reference` entity (absolute import
`from src.domain.reference.reference import Reference`). `keywords`/`paragraphs` →
`list[str]`, `sections` → `dict[str, str]`.

### 3.2 `Section` — `__post_init__` + out-of-scope helper resolution
Legacy `Section.__post_init__` does two things:
1. Raises `ValueError` if `title` is empty — **kept** (still raises `ValueError`; the
   domain exception hierarchy is Slice 1, so we do NOT introduce a `BaseSrcError` subtype
   here to avoid scope creep).
2. Auto-detects `section_type` via the **local import** of `classify_section_by_name`.
   That helper is **out of scope** (it becomes the `SectionClassifier` service in a later
   slice, plan §4.2) and currently has its own bug (returns `None` while annotated
   `-> SectionType`).

**Decision:** in this slice `Section.__post_init__` keeps ONLY the empty-title guard. The
auto-detection branch is **removed**; `section_type: SectionType | None = None` is taken
as provided by the caller. This honors "migrate data types only" — the classification
behavior moves to its own service slice. The removal is documented as an intentional,
test-covered behavior delta (a test asserts `section_type` stays `None` when not passed).
Behavior methods `get_word_count()` and `is_empty()` are preserved verbatim (typed).

> This is the only **intentional behavior change** in the slice beyond bug fixes. It is
> safe because no caller is rewired in Slice 0 (legacy `Section` still auto-detects); the
> new `src.domain.section.Section` is not yet imported by any production code.

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
`AnalysisResult` aggregates the other types. It extends `BaseDTO`, so it inherits
`as_dict()` (deep `asdict`). The legacy `to_dict()` produced a **custom, flattened** shape
(enum `.value` strings, ISO timestamp, selected fields) consumed by the formatter/exporter
/Gradio. **Decision:** preserve that exact custom serialization as an explicit
`to_dict()` method on the DTO (typed `-> dict[str, Any]`), separate from the inherited
`as_dict()`. Keeping `to_dict()` byte-compatible protects the downstream contract flagged
in plan §10.4 until the later orchestrator slice (13) formalizes it. A test pins the
`to_dict()` shape (keys + enum `.value` flattening + ISO timestamp).

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
  - **Entities:** construction; `as_dict()` returns a dict; behavior preserved
    (`Section.get_word_count`, `Section.is_empty`, `Section` empty-title raises
    `ValueError`, `Section` leaves `section_type=None` when not provided —
    documents §3.2 delta; `DocumentContent.__post_init__` computes `word_count`
    from `paragraphs` when `0`).
  - **DTOs:** construction; frozen (assignment raises `FrozenInstanceError`);
    `as_dict()`/`from_dict()` round-trip where applicable; `ClassificationResult.create()`
    factory (§4); `QualityResult` consolidation (bug #2); `Any` typing (bug #4);
    `AnalysisResult.to_dict()` shape (§5).
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

**ADR-1: `DocumentContent` and `Section` are Entities, not DTOs.**
- Decision: both extend `BaseEntity` (mutable).
- Rationale: `DocumentContent.__post_init__` mutates `word_count`; `Section.__post_init__`
  raises and the class exposes behavior. Frozen DTOs forbid both.
- Rejected: forcing them into `BaseDTO` and moving mutation to a factory — would change
  the public construction contract and break the "migrate data types only" boundary.

**ADR-2: Class names keep domain names; `_dto` is a filename marker only.**
- Decision: `ClassificationResult` (not `ClassificationResultDTO`) in file
  `classification_result_dto.py`.
- Rationale: matches plan §4.1 target paths and the use-case catalog (§6) verbatim,
  avoiding churn when later slices reference these names.
- Rejected: appending `DTO` to class names — SKILL §4 lists the suffix convention but the
  plan locked these specific names; consistency with the plan wins for migrated outputs.

**ADR-3: `Section` drops `section_type` auto-detection in Slice 0.**
- Decision: `__post_init__` keeps only the empty-title guard; classification moves to the
  future `SectionClassifier` service.
- Rationale: the classifier helper is out of scope (plan §4.2) and buggy; reproducing it
  would import an out-of-scope, broken function. The new `Section` is not yet wired into
  production, so the delta is inert.
- Rejected: copying `classify_section_by_name` into the entity — violates one-class/one-
  service-per-file and pulls a later slice forward.

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

**ADR-6: Preserve `AnalysisResult.to_dict()` custom shape verbatim.**
- Decision: keep the flattened `to_dict()` as an explicit method alongside inherited
  `as_dict()`.
- Rationale: formatter/exporter/Gradio depend on the exact flattened shape (plan §10.4);
  changing it now risks downstream breakage before the orchestrator slice.
- Rejected: replacing `to_dict()` with `as_dict()` — different shape, would break the
  downstream contract.

## 9. Risks

| Risk | Likelihood | Mitigation |
|---|---|---|
| `Section` auto-detection removal surprises a future slice | Low | Documented ADR-3 + test pins the new behavior; classifier slice re-adds detection as a service. |
| `AnalysisResult.to_dict()` shape drift vs legacy | Low | Test pins keys/values; method copied verbatim and typed. |
| Frozen DTO + `default_factory(timestamp)` misunderstanding | Low | Valid Python; covered by a construction test. |
| Slice exceeds 400-line budget (19 types + 19 tests) | Med | Measured at tasks; split into chained PRs at tasks/apply per delivery strategy if over budget — NOT split in design. |
| Accidental legacy import from `src/` | Low | Invariant enforced by review; all intra-`src` refs use absolute `src.domain.*` paths. |
