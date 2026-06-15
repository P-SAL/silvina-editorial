# SDD Tasks — validate-structure (Slice 2)

**Change**: validate-structure
**Phase**: tasks
**Date**: 2026-06-15
**Status**: active
**TDD**: STRICT (RED → GREEN)
**Test runner**: `python -m pytest src/`

---

## Phase 1 — SCAFFOLD (parallel, no tests)

### T-01 [SCAFFOLD] Create `src/domain/tests/structure/` package
- Create `src/domain/tests/structure/__init__.py` (empty)
- Pre-condition: `src/domain/tests/__init__.py` already exists ✓
- Spec ref: §4 Package Structure

### T-02 [SCAFFOLD] Create `src/domain/structure/` package
- Create `src/domain/structure/__init__.py` (empty)
- Spec ref: §4 Package Structure

### T-03 [SCAFFOLD] Create `src/application/tests/` package
- Create `src/application/tests/__init__.py` (empty)
- Pre-condition: `src/application/__init__.py` already exists ✓
- Spec ref: §4 Package Structure

> T-01, T-02, T-03 are independent and can run in parallel.

---

## Phase 2 — REQUIRED SECTIONS PROVIDER (sequential TDD loop)

### T-04 [RED] Write failing tests for `RequiredSectionsProvider`
- File: `src/domain/tests/structure/test_required_sections_provider.py`
- One `TestCase` class: `TestRequiredSectionsProvider`
- Scenarios covered: S-16, S-17
- Test methods:
  - `test_cientifico_returns_7_sections` — exact list [Resumen, Introducción, Metodología, Resultados, Discusión, Conclusiones, Referencias]
  - `test_divulgacion_returns_5_sections` — [Resumen, Introducción, Desarrollo, Conclusiones, Referencias] (Desarrollo IS required at domain level)
  - `test_opinion_returns_3_sections` — [Introducción, Argumentación, Conclusiones]
  - `test_unknown_returns_empty_list`
  - `test_desarrollo_not_in_cientifico` — S-16
  - `test_desarrollo_not_in_opinion` — S-17
- Depends on: T-01, T-02

### T-05 [GREEN] Implement `RequiredSectionsProvider`
- File: `src/domain/structure/required_sections_provider.py`
- `@staticmethod get(article_type: ArticleType) -> list[str]`
- DIVULGACION INCLUDES "Desarrollo" (faithful port of legacy `validate_structure()`)
- UNKNOWN returns `[]`
- Run: `python -m pytest src/domain/tests/structure/test_required_sections_provider.py` → all green
- Depends on: T-04

---

## Phase 3 — STRUCTURE VALIDATOR (parallel RED tasks, then single GREEN)

### T-06 [RED] Write failing tests — CIENTIFICO validation
- File: `src/domain/tests/structure/test_structure_validator_cientifico.py`
- One `TestCase` class: `TestStructureValidatorCientifico`
- Scenarios: S-01, S-02, S-03, S-07
- Test methods:
  - `test_all_7_sections_present_is_valid` — S-01
  - `test_inline_colon_format_headers_detected` — S-02 (e.g., "Resumen: texto aquí")
  - `test_missing_resumen_is_invalid` — S-03
  - `test_returns_tuple_of_lists` — S-07 (validate() returns tuple[list, list])
  - `test_missing_sections_listed_correctly`
- Depends on: T-02, T-05

### T-07 [RED] Write failing tests — DIVULGACION validation
- File: `src/domain/tests/structure/test_structure_validator_divulgacion.py`
- One `TestCase` class: `TestStructureValidatorDivulgacion`
- Scenarios: S-04, S-05
- Test methods:
  - `test_all_divulgacion_sections_present_is_valid` — S-04 (4 sections)
  - `test_missing_desarrollo_is_invalid` — S-05 (Desarrollo IS required for DIVULGACION)
- Depends on: T-02, T-05

### T-08 [RED] Write failing tests — OPINION validation
- File: `src/domain/tests/structure/test_structure_validator_opinion.py`
- One `TestCase` class: `TestStructureValidatorOpinion`
- Scenarios: S-06
- Test methods:
  - `test_all_opinion_sections_present_is_valid` — S-06
  - `test_missing_argumentacion_is_invalid`
- Depends on: T-02, T-05

### T-09 [RED] Write failing tests — UNKNOWN validation
- File: `src/domain/tests/structure/test_structure_validator_unknown.py`
- One `TestCase` class: `TestStructureValidatorUnknown`
- Scenarios: S-15
- Test methods:
  - `test_unknown_type_always_valid` — S-15
  - `test_unknown_missing_sections_is_empty`
- Depends on: T-02, T-05

### T-10 [RED] Write failing tests — alias and header detection
- File: `src/domain/tests/structure/test_structure_validator_aliases.py`
- One `TestCase` class: `TestStructureValidatorAliases`
- Scenarios: S-08, S-09, S-10, S-11, S-18
- Test methods:
  - `test_english_alias_abstract_maps_to_resumen` — S-08
  - `test_multiple_aliases_detected` (metodologia, methodology, discussion, results) — S-09
  - `test_fuentes_bibliograficas_maps_to_referencias` — S-18
  - `test_long_body_text_not_detected_as_header` (>= 100 chars, no inline pattern) — S-10
  - `test_short_header_under_100_chars_detected` — S-11
  - `test_inline_colon_keyword_detected_regardless_of_length`
- Depends on: T-02, T-05

> T-06, T-07, T-08, T-09, T-10 are independent (test files only) and can run in parallel after T-05.

### T-11 [GREEN] Implement `StructureValidator` domain service
- File: `src/domain/structure/structure_validator.py`
- Class methods:
  - `__init__(self) -> None`
  - `validate(self, document_content: DocumentContent, article_type: ArticleType) -> tuple[list[str], list[str]]` — returns (present, missing)
  - `_extract_present_sections(self, paragraphs: list[str]) -> list[str]` — 100-char threshold + inline-header rule
  - `_get_required_sections(self, article_type: ArticleType) -> list[str]` — delegates to `RequiredSectionsProvider.get()`
- `section_map` alias dict ported verbatim from `business_logic/structure_validator.py` (9 entries, all aliases preserved)
- ADR-1: 100-char threshold, NOT 5-word filter
- ADR-5: returns tuple, not StructureValidationResult (DTO construction belongs in use case)
- Run: `python -m pytest src/domain/tests/structure/` → all green (T-04 through T-10 suites)
- Depends on: T-06, T-07, T-08, T-09, T-10

---

## Phase 4 — USE CASE (sequential TDD loop)

### T-12 [RED] Write failing tests for `ValidateStructureUseCase`
- File: `src/application/tests/test_validate_structure_use_case.py`
- One `TestCase` class: `TestValidateStructureUseCase`
- Scenarios: S-12, S-13, S-14, S-20
- Test methods:
  - `test_empty_paragraphs_raises_document_empty` — S-12
  - `test_desarrollo_always_removed_from_missing_sections` — S-13
  - `test_has_references_true_removes_referencias_from_missing` — S-14
  - `test_has_references_false_preserves_referencias_in_missing` — S-15
  - `test_result_is_valid_when_all_sections_present`
  - `test_result_is_frozen` — S-20 (FrozenInstanceError on attribute set)
  - `test_missing_sections_returned_correctly`
- Depends on: T-03, T-11

### T-13 [GREEN] Implement `ValidateStructureUseCase`
- File: `src/application/validate_structure_use_case.py`
- Exact execution sequence per design:
  1. Guard: `if not document_content.paragraphs` → raise `DocumentEmpty`
  2. `present, missing = self._validator.validate(document_content, article_type)`
  3. Always: `missing = [s for s in missing if s != "Desarrollo"]` (port of `main.py:230`)
  4. Conditional: `if has_references: missing = [s for s in missing if s != "Referencias"]`
  5. `return StructureValidationResult(is_valid=len(missing) == 0, missing_sections=missing)`
- Imports: `DocumentEmpty` from `src/domain/exceptions/document_errors.py`, `StructureValidationResult` from `src/domain/dtos/structure_validation_result_dto.py`
- `has_references` is NOT forwarded to `StructureValidator` (spec §2.5)
- Run: `python -m pytest src/application/tests/` → all green
- Depends on: T-12

---

## Phase 5 — WIRING (sequential TDD loop)

### T-14 [RED] Write failing test for `ValidateStructureWiring`
- File: `src/infrastructure/tests/test_validate_structure_wiring.py`
- One `TestCase` class: `TestValidateStructureWiring`
- Scenarios: S-19
- Test methods:
  - `test_create_use_case_returns_validate_structure_use_case_instance` — S-19
  - `test_create_use_case_returns_new_instance_each_call`
- Note: `src/infrastructure/tests/__init__.py` ALREADY EXISTS — do NOT recreate
- Depends on: T-13

### T-15 [GREEN] Implement `ValidateStructureWiring`
- File: `src/infrastructure/wirings/validate_structure_wiring.py`
- `@staticmethod create_use_case() -> ValidateStructureUseCase`
- No ports, no adapters, no config (ADR-4)
- Note: `src/infrastructure/wirings/__init__.py` ALREADY EXISTS — do NOT recreate
- Run: `python -m pytest src/infrastructure/tests/test_validate_structure_wiring.py` → green
- Depends on: T-14

---

## Phase 6 — VERIFICATION

### T-16 [VERIFY] Run full test suite — zero regressions
- Command: `python -m pytest src/`
- Assertions:
  - All pre-existing tests pass (Slice 0 DTOs, Slice 1 exceptions)
  - 20 acceptance scenarios covered across new test files
  - No import errors from `src/domain/structure/` or `src/application/tests/`
  - `business_logic/structure_validator.py` untouched (legacy not modified)
  - No new exception types created (`DocumentEmpty` from Slice 1 reused)
- Depends on: T-15

---

## Dependency Graph

```
T-01 ─────────────────────────────────────────────────────────┐
T-02 ──► T-04 ──► T-05 ──► T-06 ──┐                          │
                        ──► T-07 ──┤                          │
                        ──► T-08 ──┤                          │
                        ──► T-09 ──┼──► T-11 ──► T-12 ──► T-13 ──► T-14 ──► T-15 ──► T-16
                        ──► T-10 ──┘          ▲
T-03 ─────────────────────────────────────────┘
```

**Parallel groups:**
- Group A (scaffold): T-01, T-02, T-03
- Group B (test files): T-06, T-07, T-08, T-09, T-10 (after T-05)

---

## Files Summary

### New files to create (15)

| Path | Phase |
|------|-------|
| `src/domain/tests/structure/__init__.py` | T-01 |
| `src/domain/structure/__init__.py` | T-02 |
| `src/application/tests/__init__.py` | T-03 |
| `src/domain/tests/structure/test_required_sections_provider.py` | T-04 |
| `src/domain/structure/required_sections_provider.py` | T-05 |
| `src/domain/tests/structure/test_structure_validator_cientifico.py` | T-06 |
| `src/domain/tests/structure/test_structure_validator_divulgacion.py` | T-07 |
| `src/domain/tests/structure/test_structure_validator_opinion.py` | T-08 |
| `src/domain/tests/structure/test_structure_validator_unknown.py` | T-09 |
| `src/domain/tests/structure/test_structure_validator_aliases.py` | T-10 |
| `src/domain/structure/structure_validator.py` | T-11 |
| `src/application/tests/test_validate_structure_use_case.py` | T-12 |
| `src/application/validate_structure_use_case.py` | T-13 |
| `src/infrastructure/tests/test_validate_structure_wiring.py` | T-14 |
| `src/infrastructure/wirings/validate_structure_wiring.py` | T-15 |

### Files that already exist — DO NOT recreate

| Path | Verified |
|------|---------|
| `src/application/__init__.py` | ✓ exists |
| `src/infrastructure/wirings/__init__.py` | ✓ exists |
| `src/infrastructure/tests/__init__.py` | ✓ exists |
| `src/domain/tests/__init__.py` | ✓ exists |

---

## Review Workload Forecast

| Category | Files | Est. Lines |
|----------|-------|------------|
| Scaffold `__init__.py` | 3 | ~3 |
| `test_required_sections_provider.py` | 1 | ~55 |
| `required_sections_provider.py` | 1 | ~25 |
| Domain validator test files (×5) | 5 | ~200 |
| `structure_validator.py` | 1 | ~80 |
| `test_validate_structure_use_case.py` | 1 | ~80 |
| `validate_structure_use_case.py` | 1 | ~35 |
| `test_validate_structure_wiring.py` | 1 | ~30 |
| `validate_structure_wiring.py` | 1 | ~20 |
| **Total** | **15** | **~528 lines** |

**Chained PRs recommended: Yes**
**400-line budget risk: High (~528 estimated lines)**
**Decision needed before apply: Yes**

Suggested PR split:
- **PR-A** (domain layer): T-01 + T-02 + T-04 + T-05 + T-06–T-10 + T-11 (~360 lines)
- **PR-B** (application + wiring): T-03 + T-12 + T-13 + T-14 + T-15 + T-16 (~168 lines)
