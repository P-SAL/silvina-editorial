# Design: validate-apa (Slice 3)

## Technical Approach

Lift `apa_validator.py` (root) into the hexagonal domain layer by following the
Slice 2 (`validate-structure`) pattern exactly: enum → DTO × 2 → domain service →
use case → wiring. No ports or adapters are needed — the computation is pure. The
legacy file is left untouched; all new files are additive.

---

## Architecture Decisions

| # | Decision | Choice | Rejected | Rationale |
|---|----------|--------|----------|-----------|
| ADR-1 | `generate_report()` placement | Excluded from domain | Move to use case; keep in service | It formats Spanish Unicode strings with emoji — pure presentation. Domain layer must stay language/UI-agnostic. Report generation deferred to Slice 13 (formatter adapter). Use case returns structured data only. |
| ADR-2 | `violations` field type in `ApaValidationResult` | `list[ApaViolation]` with `field(default_factory=list)` | `tuple[ApaViolation, ...]` | `StructureValidationResult` uses `list[str]` + `default_factory=list` on a frozen DTO — established project pattern. `tuple` would require a non-default field before the default `is_valid`, breaking dataclass ordering rules. Consistency wins. |
| ADR-3 | Stateless `ApaValidator` | No `__init__` state; each call returns a new list | Keep `self.violations` accumulator | `self.violations` in legacy `APAValidator.__init__` is never read by any caller. All methods already return independent lists. Removing it makes the contract explicit and enables safe concurrent use. |
| ADR-4 | Empty citations input → `is_valid=True` | Return `ApaValidationResult(is_valid=True, violation_count=0, violations=[])` | Raise guard exception | Pure computation contract: no input → no violations. Matches `validate_all_citations([])` returning `[]`. Raising an exception would be incorrect — empty is a valid state (document has no author-year citations). |

---

## File Tree

```
src/
├── domain/
│   ├── enums/
│   │   └── apa_error_type.py            # NEW — ApaErrorType (8 members)
│   ├── dtos/
│   │   ├── apa_violation_dto.py         # NEW — ApaViolation frozen DTO
│   │   └── apa_validation_result_dto.py # NEW — ApaValidationResult frozen DTO
│   ├── citation/
│   │   ├── __init__.py                  # NEW — package init
│   │   └── apa_validator.py             # NEW — ApaValidator stateless service
│   └── tests/
│       └── citation/
│           ├── __init__.py              # NEW — package init
│           ├── test_apa_validator_parenthetical.py  # NEW — checks 1–6
│           ├── test_apa_validator_narrative.py      # NEW — checks 7–9
│           └── test_apa_validator_skip_patterns.py  # NEW — non-author skip
├── application/
│   ├── validate_apa_use_case.py         # NEW — ValidateApaUseCase
│   └── tests/
│       └── test_validate_apa_use_case.py # NEW — use case behavior
└── infrastructure/
    └── wirings/
        ├── validate_apa_wiring.py       # NEW — ValidateApaWiring
        └── tests/
            └── test_validate_apa_wiring.py  # NEW — wiring smoke test
```

---

## Component Interfaces

### `ApaErrorType` enum — `src/domain/enums/apa_error_type.py`

```python
from enum import Enum

class ApaErrorType(str, Enum):
    CONJUNCTION_ERROR    = "Conjunción incorrecta"
    COMMA_ERROR          = "Puntuación incorrecta"
    CAPITALIZATION_ERROR = "Mayúsculas/minúsculas incorrectas"
    ET_AL_FORMAT_ERROR   = "Formato 'et al.' incorrecto"
    PAGE_FORMAT_ERROR    = "Formato de página incorrecto"
    SPACING_ERROR        = "Espaciado incorrecto"
    YEAR_FORMAT_ERROR    = "Formato de año incorrecto"   # defined, currently unused
    PARENTHESES_ERROR    = "Paréntesis incorrectos"      # defined, currently unused
```

Inherits `str` (same pattern as `SectionName`) so enum members serialize directly
as their string value. String values are copied verbatim from legacy `APAErrorType`
to maintain wire compatibility.

### `ApaViolation` DTO — `src/domain/dtos/apa_violation_dto.py`

```python
from dataclasses import dataclass
from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.apa_error_type import ApaErrorType

@dataclass(frozen=True)
class ApaViolation(BaseDTO):
    citation_text:     str
    error_type:        ApaErrorType
    location:          int           # paragraph index
    explanation:       str
    correction:        str
    paragraph_preview: str = ""
```

All fields are scalars or enums — no `field(default_factory=...)` needed. `frozen=True`
is inherited from `BaseDTO`'s `@dataclass(frozen=True, eq=True)`.

### `ApaValidationResult` DTO — `src/domain/dtos/apa_validation_result_dto.py`

```python
from dataclasses import dataclass, field
from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.apa_violation_dto import ApaViolation

@dataclass(frozen=True)
class ApaValidationResult(BaseDTO):
    is_valid:        bool
    violation_count: int
    violations:      list[ApaViolation] = field(default_factory=list)
```

`list` + `default_factory` follows the `StructureValidationResult.missing_sections`
pattern (ADR-2). `violation_count` is explicit to avoid forcing callers to call
`len(violations)`.

### `ApaValidator` service — `src/domain/citation/apa_validator.py`

```python
class ApaValidator:
    _NON_AUTHOR_PATTERNS: list[str] = [
        r'^\([A-Z]{2,}\s+\d',
        r'^\(arXiv:',
        r'^\(doi:',
        r'^\(repositorio',
        r'^\(no hay',
        r'^\([a-záéíóúñ].*\d{4}.*\d{4}',
        r'^\(\w+\s+\w+.*\d{4}.*\d{4}',
    ]

    def validate_citation(
        self,
        text: str,
        paragraph_index: int,
        paragraph_text: str = "",
    ) -> list[ApaViolation]: ...

    def validate_all_citations(
        self,
        citations: list[tuple[str, int, str]],
    ) -> list[ApaViolation]: ...

    def _validate_parenthetical(
        self, citation: str, location: int, preview: str = ""
    ) -> list[ApaViolation]: ...

    def _validate_narrative(
        self, citation: str, location: int, preview: str = ""
    ) -> list[ApaViolation]: ...
```

`_NON_AUTHOR_PATTERNS` is a class-level constant (not instance state). No `__init__`
needed — ADR-3. All private helpers carry over verbatim from legacy to preserve
regex correctness.

### `ValidateApaUseCase` — `src/application/validate_apa_use_case.py`

```python
from src.domain.citation.apa_validator import ApaValidator
from src.domain.dtos.apa_validation_result_dto import ApaValidationResult

class ValidateApaUseCase:
    def __init__(self, validator: ApaValidator) -> None:
        self._validator = validator

    def execute(
        self,
        citations: list[tuple[str, int, str]],
    ) -> ApaValidationResult:
        violations = self._validator.validate_all_citations(citations)
        return ApaValidationResult(
            is_valid=len(violations) == 0,
            violation_count=len(violations),
            violations=violations,
        )
```

No guard exception for empty `citations` — ADR-4. Pure pass-through to domain
service; all validation logic stays in `ApaValidator`.

### `ValidateApaWiring` — `src/infrastructure/wirings/validate_apa_wiring.py`

```python
from src.application.validate_apa_use_case import ValidateApaUseCase
from src.domain.citation.apa_validator import ApaValidator

class ValidateApaWiring:
    def create_use_case(self) -> ValidateApaUseCase:
        return ValidateApaUseCase(validator=self._get_apa_validator())

    def _get_apa_validator(self) -> ApaValidator:
        return ApaValidator()
```

Identical instance-method pattern to `ValidateStructureWiring` (ADR from Slice 2).

---

## Data Flow

```
caller (main.py, future)
    │  list[tuple[str, int, str]]
    ▼
ValidateApaWiring.create_use_case()
    │  ValidateApaUseCase
    ▼
ValidateApaUseCase.execute(citations)
    │  list[tuple[str, int, str]]
    ▼
ApaValidator.validate_all_citations(citations)
    │  iterates → validate_citation() per entry
    │      → _validate_parenthetical() or _validate_narrative()
    │  returns list[ApaViolation]
    ▼
ValidateApaUseCase  (builds result DTO)
    │  ApaValidationResult(is_valid, violation_count, violations)
    ▼
caller
```

No I/O, no external ports — pure in-memory computation at every step.

---

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/domain/enums/apa_error_type.py` | Create | `ApaErrorType` enum, 8 members, `str` mixin |
| `src/domain/dtos/apa_violation_dto.py` | Create | `ApaViolation` frozen DTO, 6 fields |
| `src/domain/dtos/apa_validation_result_dto.py` | Create | `ApaValidationResult` frozen DTO |
| `src/domain/citation/__init__.py` | Create | Empty package init |
| `src/domain/citation/apa_validator.py` | Create | `ApaValidator` stateless service |
| `src/application/validate_apa_use_case.py` | Create | `ValidateApaUseCase` |
| `src/infrastructure/wirings/validate_apa_wiring.py` | Create | `ValidateApaWiring` |
| `src/domain/tests/citation/__init__.py` | Create | Empty package init |
| `src/domain/tests/citation/test_apa_validator_parenthetical.py` | Create | Checks 1–6 (parenthetical) |
| `src/domain/tests/citation/test_apa_validator_narrative.py` | Create | Checks 7–9 (narrative) |
| `src/domain/tests/citation/test_apa_validator_skip_patterns.py` | Create | Non-author skip pattern tests |
| `src/application/tests/test_validate_apa_use_case.py` | Create | Use case behavior (empty, violations, is_valid) |
| `src/infrastructure/tests/test_validate_apa_wiring.py` | Create | Wiring smoke test |
| `apa_validator.py` (root) | Unchanged | Legacy coexistence until Slice 14 |

---

## Test Architecture

| File | Class | Covers |
|------|-------|--------|
| `test_apa_validator_parenthetical.py` | `TestApaValidatorParenthetical` | Check 1 (ampersand), Check 2 (missing comma), Check 3 (lowercase), Check 4 (et al.), Check 5 (page format), Check 6 (spacing) |
| `test_apa_validator_narrative.py` | `TestApaValidatorNarrative` | Check 7 (ampersand), Check 8 (et al.), Check 9 (space before year) |
| `test_apa_validator_skip_patterns.py` | `TestApaValidatorSkipPatterns` | All 7 non-author patterns return no extra violations |
| `test_validate_apa_use_case.py` | `TestValidateApaUseCase` | empty → `is_valid=True`; 1 violation → `is_valid=False, violation_count=1`; violation list passthrough |
| `test_validate_apa_wiring.py` | `TestValidateApaWiring` | `create_use_case()` returns `ValidateApaUseCase`; `_get_apa_validator()` returns `ApaValidator` |

One `TestCase` per file. Domain tests exercise `ApaValidator` in isolation (no use
case). Application test exercises `ValidateApaUseCase` with a real `ApaValidator`
(no mock needed — pure computation). Infrastructure test is a smoke test only.

---

## Migration / Rollout

No migration required. All 13 new files are additive. `apa_validator.py` (root) and
`main.py` are untouched. Rollback: delete the new files.

---

## Open Questions

- None. All decisions are resolved.
