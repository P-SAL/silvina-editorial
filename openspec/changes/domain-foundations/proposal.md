# Proposal: Domain Foundations (Slice 0)

> Slice 0 of the hexagonal migration plan (`docs/plan-migracion-hexagonal.md`).
> Normative guide: `.agent/skills/clean-architecture/SKILL.md`.

## Intent

Migrate the **pure domain foundations** — enums, entities, and DTOs — from the
legacy `domain/enums.py` and `domain/models.py` into the existing `src/domain/`
skeleton, with **zero I/O, ports, adapters, or use cases**.

These types are the shared vocabulary that **all 15 remaining slices depend on**.
Until they exist as first-class `src/domain/` artifacts (`BaseEntity` / `BaseDTO`
subclasses, one class per file, `unittest.TestCase` coverage), no downstream
slice (ValidateStructure, ValidateApa, adapters, use cases) can be built on a
clean base. This slice unblocks the rest of the migration.

The legacy `domain/` package keeps working unchanged — old and new code
**coexist**. This slice migrates types into `src/`; it does NOT rewire any
caller. No legacy file is deleted or modified.

### Why now

The migration plan is sequenced from lowest to highest external coupling. The
domain foundations have **no external dependencies**, so they are the safest and
most valuable first move: every later slice imports them. Doing this first also
lets us **fix the latent bugs and duplication** in the legacy domain (see
"Bug Fixes During Migration") before they get copied forward into 15 slices.

### Success looks like

- Every enum, entity, and DTO listed in scope lives in `src/domain/` following
  the one-class-per-file convention, inheriting `BaseEntity` or `BaseDTO`.
- Each migrated type has a `unittest.TestCase` in `src/domain/tests/`, written
  test-first (Strict TDD is active).
- `python -m pytest src/` passes green.
- The documented bugs are corrected in the migrated types (not carried forward).
- Legacy `domain/`, `business_logic/`, `data_access/`, `presentation/`, the
  entry points, and the existing pytest suite all still run exactly as before.

## Scope

### In Scope

- **Enums** from `domain/enums.py`: `ArticleType`, `ArticleSize`, `CitationType`,
  `ClassificationCategory`, `QualityLevel`, `SectionType`, `AnalysisDimension`,
  `ValidationStatus`, `RecommendationPriority`, `SeverityLevel` → one file each
  under `src/domain/enums/`.
- **Domain data types** from `domain/models.py`: `Citation`, `Reference`,
  `Section`, `DocumentContent`, `ClassificationResult`, `QualityResult`,
  `StructureValidationResult`, `CitationAnalysisResult`, `AnalysisResult` →
  migrated as **either** `BaseEntity` (mutable, behavior) **or** `BaseDTO`
  (immutable, crosses boundaries). The per-type Entity-vs-DTO decision is
  **deferred to the spec phase** (see Open Decisions).
- **Consolidation and bug fixes** of each migrated type, documented below.
- **`unittest.TestCase` tests** per migrated type under `src/domain/tests/`.
- **Coexistence** with all legacy code (no legacy file touched).

### Out of Scope

- Ports, adapters, use cases, wirings (later slices 1–16).
- Any I/O: `python-docx`, `win32com`, `ollama`, `language_tool_python`.
- Entry points (`main.py`, `gradio_app.py`) and any caller rewiring.
- Domain services / the loose `enums.py` helpers
  (`classify_article_size`, `get_quality_level_from_score`,
  `classify_section_by_name`, `get_required_sections_for_category`,
  `get_citation_type_from_pattern`) and the `models.py` factory functions.
  These become service classes / `@classmethod` factories in **their own
  later slices** (plan §4.2). This slice migrates **only the data types**.
- Domain exceptions hierarchy population (that is Slice 1).
- Deleting or modifying legacy `domain/enums.py` / `domain/models.py`
  (final cleanup is Slice 16).

## Approach

Migrate the pure data types into the `src/domain/` skeleton **type by type**,
test-first, following the clean-architecture SKILL. The skeleton is not
restructured — only filled (`src/domain/enums/`, `src/domain/entities/`,
`src/domain/dtos/`, `src/domain/tests/` already exist).

**Per-type loop (Strict TDD, runner `python -m pytest src/`):**

1. Write a failing `unittest.TestCase` in `src/domain/tests/<topic>/test_<class>.py`.
2. Migrate the type into its own snake_case file (PascalCase class):
   - Enum → `src/domain/enums/<enum>.py`.
   - Entity → `src/domain/<entity>/<entity>.py`, subclass of `BaseEntity`,
     `@dataclass` (mutable).
   - DTO → `src/domain/dtos/<name>_dto.py`, subclass of `BaseDTO`
     (`@dataclass(frozen=True)`, already provided by `BaseDTO`).
3. Apply the documented bug fix for that type while migrating it.
4. Run `python -m pytest src/` green; move to the next type.

**Conventions enforced** (SKILL + plan §9):
- One class per file; snake_case file = PascalCase class.
- `X | None`, never `Optional[X]`; `list[T]` / `dict[K, V]`, never `List`/`Dict`.
- Specific-name imports only; no wildcard; **no local/in-function imports**.
- No top-level `services/` / `ports/` / `config/` folders.
- Docstrings required on public classes/methods (PEP 257); no inline `#` comments.

### Bug Fixes During Migration (decision: fix, do not carry forward)

The user decided we **correct bugs and duplication as each type is migrated**,
documenting every change. Each fix is verified by the type's `unittest` test.
Confirmed issues in the legacy code:

| # | Source | Bug | Fix on migration |
|---|--------|-----|------------------|
| 1 | `enums.py` `__all__` (line 277) | Lists `SeverityLevel` **before** it is defined (line 284); also omits `ArticleType` / `ArticleSize`. | Each enum is its own module; no shared `__all__`. The ordering bug disappears by construction. |
| 2 | `models.py` `QualityResult` vs `QualityAnalysisResult` | Two overlapping dataclasses for the same concept; `QualityAnalysisResult` is a thinner duplicate with no `__str__`/timestamp. | Consolidate into **one** migrated type (keep the richer `QualityResult` shape). Final name/fields settled in spec. |
| 3 | `models.py` `create_classification_result` | Passes `category=` and only 3 fields to `ClassificationResult`, whose required fields are `article_type` + `article_size` → `TypeError` at runtime (helper is broken). | Helper is **out of this slice's scope** (it is a factory, plan §4.2). The migrated `ClassificationResult` keeps correct field names; the broken legacy helper is left untouched in legacy code and not reproduced in `src/`. |
| 4 | `models.py` (lines 90, 102, 201) | `Dict[str, Dict[str, any]]` uses the builtin **function** `any` instead of the **type** `Any`. | Migrate as `dict[str, dict[str, Any]]` with `Any` imported from `typing`. |
| 5 | `enums.py` `classify_section_by_name` | Return annotated `-> SectionType` but returns `None` on no match. | Helper is out of scope (plan §4.2); when its slice runs the signature becomes `-> SectionType | None`. Noted here so it is not forgotten. |
| 6 | `enums.py` `get_required_sections_for_category` | Annotated bare `-> list`. | Out of scope (helper). When migrated: `-> list[SectionType]`. Noted for traceability. |
| 7 | `models.py` imports (lines 10–11) | Mixes `from .enums import ...` and `from domain.enums import ...`; `Section.__post_init__` and `get_citation_type_from_pattern` use **local** imports. | Migrated types use absolute `from src.domain.enums...` imports at module top; no local imports (SKILL §3). |

> Bugs #5 and #6 live in helper functions that are **out of this slice's
> scope**. They are listed only so they are not lost; they are fixed when their
> service-class slice runs (plan §4.2). This slice fixes only #1, #2, #4, #7,
> which touch the data types being migrated.

## Open Decisions (resolved in the spec phase)

1. **Entity vs DTO mapping — NOT decided here.** The plan §4.1 *suggests* a
   mapping (`Citation`/`Reference`/`Section`/`DocumentContent` → Entity;
   the `*Result` types → DTO) but this proposal **does not lock it**. The spec
   phase decides **case by case** per the SKILL criteria: Entity = mutable +
   behavior (`__post_init__`, `get_word_count`, `is_empty`); DTO = immutable,
   crosses boundaries. Notably `DocumentContent` has mutating `__post_init__`
   logic and `Section` has behavior methods — these need explicit per-type
   judgment in spec. A frozen `BaseDTO` cannot host the current mutating
   `__post_init__`, so that interaction must be resolved in spec.
2. **`QualityResult` consolidation target** — final name and exact field set of
   the single consolidated quality type (decided in spec alongside #1).
3. **`AnalysisResult.to_dict()` shape** — `AnalysisResult` aggregates the other
   result types; its serialization contract (consumed later by formatter /
   exporter / Gradio) is defined in spec, not locked here.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/domain/enums/` | New files | One module per migrated enum. |
| `src/domain/entities/` and/or `src/domain/<entity>/` | New files | Entities migrated as `BaseEntity` subclasses (placement per spec mapping). |
| `src/domain/dtos/` | New files | DTOs migrated as `BaseDTO` subclasses. |
| `src/domain/tests/` | New files | `unittest.TestCase` per migrated type. |
| `domain/enums.py`, `domain/models.py` | Untouched | Legacy kept for coexistence; removed in Slice 16. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Entity-vs-DTO mapping chosen wrong, churns later slices | Med | Defer to spec, decide per SKILL criteria case by case; this proposal locks nothing. |
| `DocumentContent.__post_init__` mutation conflicts with a frozen DTO | Med | Flagged as an explicit spec open decision (#1); resolve placement before implementing the type. |
| Silently re-introducing a legacy bug | Low | Every fix is tied to a `unittest` assertion; bug table is the checklist. |
| Slice exceeds the 400-line budget | Med | Single change; size measured at tasks. If over budget, split into chained PRs at tasks/apply (per delivery strategy) — NOT split now. |
| Breaking legacy coexistence | Low | No legacy file is modified; `src/` only adds new modules; legacy pytest suite must still pass. |

## Rollback Plan

All work is **additive** under `src/domain/` plus new tests. No legacy
production code is modified. Rollback = delete the new `src/domain/` modules and
their tests; legacy behavior is unaffected.

## Dependencies

- Python 3.10+.
- Existing skeleton: `BaseEntity` (`src/domain/entities/base_entity.py`) and
  `BaseDTO` (`src/domain/dtos/base_dto.py`) — already present.
- No external libraries (pure domain).

## Success Criteria

- [ ] All in-scope enums, entities, and DTOs migrated to `src/domain/`,
      one class per file, inheriting `BaseEntity` / `BaseDTO`.
- [ ] Each migrated type has a `unittest.TestCase`; `python -m pytest src/` is green.
- [ ] Documented bugs (#1, #2, #4, #7) corrected in migrated types, each covered by a test.
- [ ] No `Optional`, no `List`/`Dict`, no wildcard imports, no local imports in migrated code.
- [ ] Legacy `domain/`, downstream packages, entry points, and the existing pytest suite still run unchanged.
