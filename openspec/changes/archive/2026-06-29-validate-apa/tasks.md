# SDD Tasks — validate-apa (Slice 3)

**Change**: validate-apa
**Phase**: tasks
**Date**: 2026-06-15
**Status**: READY
**TDD**: STRICT (RED → GREEN)
**Test runner**: `python -m pytest src/`

---

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 380–460 |
| 400-line budget risk | Medium |
| Chained PRs recommended | No |
| Suggested split | Single PR — pure domain + application, no ports/adapters |
| Delivery strategy | single-pr |
| Chain strategy | size-exception if diff exceeds 400 |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Medium

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | All 11 new files: enum, DTOs, domain service, use case, wiring + tests | PR 1 | feat(apa-validator): targets refactor/hexagonal-migration |

---

## Phase 1 — SCAFFOLD (sequential, no tests yet)

### [x] T-01 [SCAFFOLD] Create src/domain/citation/ package
- Create `src/domain/citation/__init__.py` (empty)
- Spec ref: §File Map

### [x] T-02 [SCAFFOLD] Create src/domain/tests/citation/ package
- Create `src/domain/tests/citation/__init__.py` (empty)
- Pre-condition: `src/domain/tests/__init__.py` already exists
- Spec ref: §File Map

### [x] T-03 [SCAFFOLD] Create src/application/tests/ package (if absent)
- Create `src/application/tests/__init__.py` (empty)
- Check first: may exist from Slice 2
- Spec ref: §File Map

---

## Phase 2 — ENUM + DTOs (sequential TDD loop)

### [x] T-04 [RED] Write failing tests for ApaErrorType enum
- File: `src/domain/tests/citation/test_apa_validator_parenthetical.py` (add enum assertions at top)
- Assert all 8 members exist with correct string values
- Assert YEAR_FORMAT_ERROR and PARENTHESES_ERROR exist (unused but preserved)
- Confirmed RED before implementation
- Spec ref: §1 ApaErrorType

### [x] T-05 [GREEN] Implement ApaErrorType enum
- File: `src/domain/enums/apa_error_type.py`
- `class ApaErrorType(str, Enum)` — 8 members, exact Spanish string values per spec §1
- All enum assertions GREEN

### [x] T-06 [RED] Write failing tests for ApaViolation DTO
- File: `src/domain/tests/citation/test_apa_validator_parenthetical.py` (add DTO assertions)
- Assert frozen: `FrozenInstanceError` on mutation (S-14b)
- Assert fields: citation_text, error_type, location, explanation, correction, paragraph_preview=""
- Spec ref: §2 ApaViolation, S-14b

### [x] T-07 [GREEN] Implement ApaViolation DTO
- File: `src/domain/dtos/apa_violation_dto.py`
- `@dataclass(frozen=True)` extending BaseDTO; 6 fields per spec §2
- All DTO assertions GREEN

### [x] T-08 [RED] Write failing tests for ApaValidationResult DTO
- File: `src/domain/tests/citation/test_apa_validator_parenthetical.py` (add result DTO assertions)
- Assert frozen (S-14); assert is_valid, violation_count, violations fields
- Assert invariant: is_valid == (violation_count == 0)
- Spec ref: §3 ApaValidationResult, S-14

### [x] T-09 [GREEN] Implement ApaValidationResult DTO
- File: `src/domain/dtos/apa_validation_result_dto.py`
- `@dataclass(frozen=True)` extending BaseDTO; violations uses `field(default_factory=list)`
- All result DTO assertions GREEN

---

## Phase 3 — DOMAIN SERVICE (sequential TDD loop, 3 test files)

### [x] T-10 [RED] Write failing tests — parenthetical citations
- File: `src/domain/tests/citation/test_apa_validator_parenthetical.py`
- Class: `TestApaValidatorParenthetical`
- Covers S-01 (valid), S-02 (CONJUNCTION_ERROR), S-03 (COMMA_ERROR), S-04 (CAPITALIZATION_ERROR), S-05 (ET_AL_FORMAT_ERROR extra period), S-05b (et al no trailing period), S-06 (PAGE_FORMAT_ERROR), S-07 (SPACING_ERROR)
- All tests confirmed RED before implementation

### [x] T-11 [RED] Write failing tests — narrative citations
- File: `src/domain/tests/citation/test_apa_validator_narrative.py`
- Class: `TestApaValidatorNarrative`
- Covers S-08 (ampersand in narrative → CONJUNCTION_ERROR), S-09 (et. al in narrative → ET_AL_FORMAT_ERROR), S-10 (missing space before year → SPACING_ERROR)
- All tests confirmed RED before implementation

### [x] T-12 [RED] Write failing tests — non-author skip patterns
- File: `src/domain/tests/citation/test_apa_validator_skip_patterns.py`
- Class: `TestApaValidatorSkipPatterns`
- Covers S-11 (UNESCO → no CAPITALIZATION_ERROR/COMMA_ERROR), S-11b (arXiv → no violations), S-11c (DOI → no violations); one test per pattern (7 patterns total)
- All tests confirmed RED before implementation

### [x] T-13 [GREEN] Implement ApaValidator domain service
- File: `src/domain/citation/apa_validator.py`
- Stateless class (no `__init__` state, no self.violations — ADR-3)
- `_NON_AUTHOR_PATTERNS`: 7 exact regexes from spec §4.1, applied with `re.search(..., re.IGNORECASE)`
- `_validate_parenthetical()`: 6 checks in order per spec §4.2; non-author skip after conjunction check
- `_validate_narrative()`: regex pattern + 3 checks per spec §4.3; silent skip if no match
- `validate_citation(text, paragraph_index, paragraph_text="") -> list[ApaViolation]`
- `validate_all_citations(citations: list[tuple[str, int, str]]) -> list[ApaViolation]`; empty → []
- All parenthetical + narrative + skip pattern tests GREEN (30+ tests)

---

## Phase 4 — USE CASE (sequential TDD loop)

### [x] T-14 [RED] Write failing tests for ValidateApaUseCase
- File: `src/application/tests/test_validate_apa_use_case.py`
- Class: `TestValidateApaUseCase`
- Covers S-12 (empty list → is_valid=True, violation_count=0, no exception — ADR-4), S-13 (violation_count and is_valid computed correctly from violations list)
- Confirmed RED before implementation

### [x] T-15 [GREEN] Implement ValidateApaUseCase
- File: `src/application/validate_apa_use_case.py`
- `execute(citations: list[tuple[str, int, str]]) -> ApaValidationResult`
- Empty → `ApaValidationResult(is_valid=True, violation_count=0, violations=[])` (ADR-4)
- Otherwise: delegate to `ApaValidator.validate_all_citations()`, compute `violation_count=len(violations)`, `is_valid=(violation_count==0)`
- All use case tests GREEN

---

## Phase 5 — WIRING (sequential TDD loop)

### [x] T-16 [RED] Write failing tests for ValidateApaWiring factory
- File: `src/infrastructure/tests/test_validate_apa_wiring.py`
- Class: `TestValidateApaWiring`
- Covers S-15 (wiring creates valid use case; `execute([])` returns `ApaValidationResult`)
- 2 tests: correct types returned from `create_use_case()` and smoke execute
- Confirmed RED before implementation

### [x] T-17 [GREEN] Implement ValidateApaWiring
- File: `src/infrastructure/wirings/validate_apa_wiring.py`
- `create_use_case(self) -> ValidateApaUseCase`; `_get_apa_validator(self) -> ApaValidator`
- Instance-based pattern (matches Slice 2 ValidateStructureWiring — no @classmethod; follow existing pattern)
- All 2 wiring tests GREEN

---

## Phase 6 — VERIFICATION (final, after all GREEN)

### [x] T-18 [VERIFY] Run full test suite, confirm zero regressions
- Command: `python -m pytest src/`
- Expected: all prior tests + all new tests pass; zero regressions
- Legacy `apa_validator.py` (root) untouched — verify it still passes its own tests if any
