# SDD Design — validate-structure

## Executive Summary

A pure-domain hexagonal slice: `StructureValidator` domain service + `RequiredSectionsProvider` domain class + `ValidateStructureUseCase` application use case + production wiring factory. No ports or adapters required. All dependencies are already-migrated domain primitives (Slice 0 DTOs/enums, Slice 1 exceptions).

---

## Architecture Overview

```
src/
  domain/
    structure/
      __init__.py
      structure_validator.py          <- domain service (merged validator + analyzer internals)
      required_sections_provider.py   <- pure domain class: ArticleType -> list[str]
    tests/
      structure/
        __init__.py
        test_structure_validator_cientifico.py
        test_structure_validator_divulgacion.py
        test_structure_validator_opinion.py
        test_structure_validator_unknown.py
        test_structure_validator_aliases.py
        test_required_sections_provider.py
  application/
    __init__.py                       <- already exists (empty)
    validate_structure_use_case.py    <- use case with has_references rule
    tests/
      __init__.py                     <- NEW - scaffold for application-layer tests
      test_validate_structure_use_case.py
  infrastructure/
    wirings/
      __init__.py                     <- already exists
      validate_structure_wiring.py    <- factory, no adapters
```

---

## Component Interfaces

### `RequiredSectionsProvider` — `src/domain/structure/required_sections_provider.py`

```python
class RequiredSectionsProvider:
    @staticmethod
    def get(article_type: ArticleType) -> list[SectionName]:
        ...
```

- Returns `list[SectionName]` — enum members, not raw strings.
- Faithfully ports the hardcoded lists from legacy `validate_structure()`.
- Pure function — no instance state, no I/O.

Section sets (verbatim from legacy `structure_validator.py`):

| ArticleType | Required sections (SectionName members) |
|-------------|----------------------------------------|
| CIENTIFICO | [SUMMARY, INTRODUCTION, METHODOLOGY, RESULTS, DISCUSSION, CONCLUSIONS, REFERENCES] (7) |
| DIVULGACION | [SUMMARY, INTRODUCTION, DEVELOPMENT, CONCLUSIONS, REFERENCES] (5) |
| OPINION | [INTRODUCTION, ARGUMENTATION, CONCLUSIONS] (3) |
| UNKNOWN | [] |

Note: DIVULGACION INCLUDES `SectionName.DEVELOPMENT` at the domain level (faithful port of legacy).
The use case removes it from `missing_sections` unconditionally (faithful port of `main.py:230`).

---

### `StructureValidator` — `src/domain/structure/structure_validator.py`

```python
class StructureValidator:
    _SECTION_ALIASES: dict[SectionName, list[str]] = { ... }  # 9 entries, SectionName keys

    def __init__(self) -> None: ...

    def validate(
        self,
        document_content: DocumentContent,
        article_type: ArticleType,
    ) -> tuple[list[SectionName], list[SectionName]]:
        """Returns (present_sections, missing_sections)."""
        ...

    def _extract_present_sections(self, paragraphs: list[str]) -> list[SectionName]:
        """Returns SectionName members found using 100-char threshold."""
        ...

    def _get_required_sections(self, article_type: ArticleType) -> list[SectionName]:
        """Delegates to RequiredSectionsProvider.get()."""
        ...
```

**Design note**: `validate()` returns a raw `tuple` rather than a `StructureValidationResult` DTO.
This is intentional — the DTO is frozen=True, so post-processing must happen in the use case
before construction. The domain service produces data; the use case constructs the immutable result.

The caller (use case) discards `present_sections` via `_, missing = self._validator.validate(...)`.
`present_sections` is available for future use cases that need it.

---

### `ValidateStructureUseCase` — `src/application/validate_structure_use_case.py`

```python
class ValidateStructureUseCase:
    def __init__(self, validator: StructureValidator) -> None:
        self._validator = validator

    def execute(
        self,
        document_content: DocumentContent,
        article_type: ArticleType,
        has_references: bool = False,
    ) -> StructureValidationResult:
        ...
```

**Exact execution sequence**:

1. Guard: `if not document_content.paragraphs` → raise `DocumentEmpty`
2. `_, missing = self._validator.validate(document_content, article_type)`
   (`present_sections` discarded — use case only needs `missing`)
3. Post-process (order matters):
   a. Always: `missing = [s for s in missing if s != SectionName.DEVELOPMENT]` (port of `main.py:230`)
   b. Conditional: `if has_references: missing = [s for s in missing if s != SectionName.REFERENCES]`
4. Construct and return: `StructureValidationResult(is_valid=len(missing) == 0, missing_sections=list(missing))`

DTO fields used: `is_valid`, `missing_sections` (the DTO also has `section_details` and `timestamp` with defaults — not set here).

---

### `ValidateStructureWiring` — `src/infrastructure/wirings/validate_structure_wiring.py`

```python
class ValidateStructureWiring:
    def create_use_case(self) -> ValidateStructureUseCase:
        return ValidateStructureUseCase(validator=self._get_structure_validator())

    def _get_structure_validator(self) -> StructureValidator:
        return StructureValidator()
```

**Pattern**: instance method (not staticmethod) + one `_get_*` private method per dependency.
This allows test subclasses to override `_get_structure_validator()` to inject mocks without
changing the composition logic. All future wirings in this project MUST follow this pattern.

No ports, no adapters, no config injection.

---

## Header Detection Algorithm

Source: `business_logic/structure_validator.py::_extract_present_sections()` — ported exactly.

```
for each paragraph in document_content.paragraphs:
    text_lower = paragraph.lower().strip()
    is_short_header = len(text_lower) < 100
    is_inline_header = text_lower starts with "keyword:" or "keyword :" for any alias
    if is_short_header OR is_inline_header:
        scan section_map for keyword match -> append canonical name (capitalized)
```

**section_map — must be ported exactly (alias dict is the behavioral contract)**:

```python
section_map = {
    'resumen':       ['resumen', 'abstract'],
    'introducción':  ['introducción', 'introduccion', 'introduction'],
    'metodología':   ['metodología', 'metodologia', 'methodology'],
    'resultados':    ['resultados', 'results'],
    'discusión':     ['discusión', 'discusion', 'discussion'],
    'argumentación': ['argumentación', 'argumentacion', 'argumentation'],
    'desarrollo':    ['desarrollo', 'development'],
    'conclusiones':  ['conclusiones', 'conclusión', 'conclusion'],
    'referencias':   ['referencias', 'bibliografía', 'bibliografia', 'fuentes bibliográficas'],
}
```

Output: `section_name.capitalize()` — e.g., 'resumen' -> "Resumen", 'introducción' -> "Introducción".

---

## Data Flow

```
caller
  |
  v
ValidateStructureWiring.create_use_case()
  |  instantiates
  v
ValidateStructureUseCase
  |  execute(document_content, article_type, has_references)
  |
  +-- guard: raises DocumentEmpty if paragraphs is empty
  |
  +-- StructureValidator.validate(document_content, article_type)
  |     |
  |     +-- _get_required_sections(article_type)
  |     |     └-- RequiredSectionsProvider.get(article_type) -> list[str]
  |     |
  |     └-- _extract_present_sections(paragraphs) -> list[str]
  |           (100-char threshold + section_map alias scan)
  |
  +-- post-process: always remove SectionName.DEVELOPMENT from missing
  |
  +-- post-process: remove SectionName.REFERENCES if has_references is True
  |
  └-- construct StructureValidationResult(is_valid, missing_sections)
        frozen=True, built once, never mutated
```

---

## ADR-1: 100-char threshold over 5-word filter

**Decision**: Use `len(paragraph) < 100` (from `StructureValidator._extract_present_sections`) as the header detection gate, NOT the `1 <= word_count <= 5` filter from `StructureAnalyzer.analyze()`.

**Rationale**:
- The 100-char threshold is the production algorithm that backs the 10 existing behavioral tests. Switching to 5-word would break detection of inline headers like "Resumen: Este artículo..." and multi-word section names like "Fuentes bibliográficas".
- `StructureAnalyzer` was a supplementary IMRyD signal tool (not the canonical section detector). Its conservative 5-word filter is appropriate for pattern detection, not for editorial validation.
- The OR condition (`is_short_header OR is_inline_header`) provides additional coverage that a pure word-count filter cannot.

**Rejected**: 5-word filter from `StructureAnalyzer` — would fail inline headers and break the test suite.

---

## ADR-2: has_references is a bool primitive, not a citation list

**Decision**: `has_references: bool = False` parameter on `execute()`.

**Rationale**:
- The use case's responsibility is structure validation, not citation analysis. Whether references exist is an external fact produced by a different use case.
- Passing a citation list would couple this use case to citation domain types, violating slice isolation.
- A bool is the minimal interface: the use case only needs to know "does the document have references?" not how many or what kind.
- Default `False` = conservative: if the caller doesn't know, we treat References as required (stricter, safer for editorial feedback).

**Rejected**: `list[Reference]` parameter — unnecessary coupling; `int` count — still couples to citation domain concept without adding value over bool.

---

## ADR-3: DEVELOPMENT included at domain level, removed at use-case level (legacy faithful port)

**Decision**: `RequiredSectionsProvider.get(DIVULGACION)` INCLUDES `SectionName.DEVELOPMENT`.
`ValidateStructureUseCase.execute()` removes `SectionName.DEVELOPMENT` from `missing_sections`
unconditionally — regardless of article type, every call strips it.

**Rationale**:
- Legacy `business_logic/structure_validator.py` includes "Desarrollo" in DIVULGACION's required list.
- Legacy `main.py` line 230 removes it unconditionally after every `validate_structure()` call.
- The domain service is a faithful port of `structure_validator.py`; the use case is a faithful port
  of `main.py`'s orchestration logic. Splitting them accurately preserves the original behavior.
- The removal is at use-case level because it is an application-layer editorial business rule
  (not a domain invariant): the domain correctly models DIVULGACION as requiring Desarrollo;
  the application layer decides to forgive its absence.

**Rejected**: Exclude DEVELOPMENT from RequiredSectionsProvider — would diverge from legacy domain
model and lose the ability to detect Desarrollo when present (which IS reported in `present_sections`).

---

## ADR-4: Wiring is instance-based with `_get_*` per dependency

**Decision**: `ValidateStructureWiring` is an instantiable class. `create_use_case(self)` is an
instance method. Each dependency is created in its own `_get_*(self)` private method.

**Rationale**:
- Instance methods allow test subclasses to override individual `_get_*` methods to inject mocks/stubs
  without touching the composition logic. `@staticmethod` makes this impossible without monkey-patching.
- The `_get_*` pattern scales naturally: when future slices add adapters, databases, or config, each
  becomes a new `_get_*` method — the `create_use_case()` body stays clean and readable.
- All wirings in this project follow this same structure for consistency.

**Rejected**: `@staticmethod create_use_case()` — cannot be subclassed for testing; rejected after
implementation revealed the pattern from the hexagonal architecture template.

**Rejected**: Port interface `IStructureValidator` — no polymorphic behavior needed; the service is
concrete and pure.

---

## ADR-5: validate() returns tuple, not StructureValidationResult

**Decision**: `StructureValidator.validate()` returns `tuple[list[str], list[str]]` (present, missing), not a `StructureValidationResult`.

**Rationale**:
- `StructureValidationResult` is frozen=True. If the domain service built it, the use case could not apply the `has_references` post-processing (mutation impossible after construction).
- Returning a tuple separates concerns cleanly: the domain service produces data, the use case applies business rules, then constructs the immutable result exactly once.

**Rejected**: Domain service builds StructureValidationResult — frozen constraint prevents use-case post-processing; mutable intermediate DTO — adds unnecessary types with no benefit.

---

## Integration Points

**Inputs (already exist — Slice 0 / Slice 1)**:
- `DocumentContent` — `src/domain/dtos/document_content_dto.py`
- `ArticleType` — `src/domain/enums/article_type.py`
- `DocumentEmpty` — `src/domain/exceptions/document_errors.py`

**Output (already exists — Slice 0)**:
- `StructureValidationResult` — `src/domain/dtos/structure_validation_result_dto.py`

No new external dependencies. No port/adapter interfaces.

---

## Test Architecture

Convention: one `unittest.TestCase` class per file (established project pattern).

### Domain tests — `src/domain/tests/structure/`

| File | Coverage |
|------|----------|
| `test_structure_validator_cientifico.py` | CIENTIFICO requires 7 sections; all present → valid; each missing individually |
| `test_structure_validator_divulgacion.py` | DIVULGACION requires 4 sections; Desarrollo NOT flagged as missing |
| `test_structure_validator_opinion.py` | OPINION requires 3 sections |
| `test_structure_validator_unknown.py` | UNKNOWN → 0 required → always valid regardless of paragraphs |
| `test_structure_validator_aliases.py` | English aliases (abstract, introduction, methodology, results, discussion); inline "Resumen: ..." format; long paragraph (>= 100 chars) excluded |
| `test_required_sections_provider.py` | All ArticleType values; DESARROLLO never returned for any type |

### Application tests — `src/application/tests/`

| File | Coverage |
|------|----------|
| `test_validate_structure_use_case.py` | Empty doc raises DocumentEmpty; has_references=True removes Referencias from missing; has_references=False keeps Referencias; valid doc stays valid; missing sections correct per type |

---

## Files to Create

| Path | Type |
|------|------|
| `src/domain/structure/__init__.py` | Package marker |
| `src/domain/structure/structure_validator.py` | Domain service |
| `src/domain/structure/required_sections_provider.py` | Domain class |
| `src/domain/tests/structure/__init__.py` | Package marker |
| `src/domain/tests/structure/test_structure_validator_cientifico.py` | Domain test |
| `src/domain/tests/structure/test_structure_validator_divulgacion.py` | Domain test |
| `src/domain/tests/structure/test_structure_validator_opinion.py` | Domain test |
| `src/domain/tests/structure/test_structure_validator_unknown.py` | Domain test |
| `src/domain/tests/structure/test_structure_validator_aliases.py` | Domain test |
| `src/domain/tests/structure/test_required_sections_provider.py` | Domain test |
| `src/application/validate_structure_use_case.py` | Use case |
| `src/application/tests/__init__.py` | Package marker (NEW) |
| `src/application/tests/test_validate_structure_use_case.py` | Application test |
| `src/infrastructure/wirings/validate_structure_wiring.py` | Wiring factory |

## Files NOT Touched

| Path | Reason |
|------|--------|
| `business_logic/structure_validator.py` | Legacy — no deletion in this slice |
| `business_logic/structure_analyzer.py` | Legacy — no deletion in this slice |
| `main.py` | Not wired into new use case yet |
| `src/domain/exceptions/` | No new exception types needed |
| `src/domain/dtos/structure_validation_result_dto.py` | DTO already correct; no `present_sections` field added |

---

## Risks

1. **Alias dict must be ported exactly**: The `section_map` in `_extract_present_sections` is the behavioral contract for the 10 legacy tests. Any omission or typo in keyword lists causes test failures.

2. **StructureValidationResult has no `present_sections` field**: The DTO schema has `is_valid`, `missing_sections`, `section_details`, `timestamp` only. The design correctly uses only those fields. If `present_sections` needs surfacing in the future, that is an out-of-scope DTO change.

3. **DIVULGACION section set**: Legacy `validate_structure` includes "Referencias" in DIVULGACION but NOT "Desarrollo". Care required: do not accidentally add "Desarrollo" to the provider.

4. **`src/application/tests/` does not exist**: The `__init__.py` must be created. Test discovery will silently skip the directory without it.

5. **One-TestCase-per-file convention**: Must create multiple domain test files (one per ArticleType + aliases), not collapse into one file.
