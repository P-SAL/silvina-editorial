# SDD Proposal — validate-structure

> **Status**: proposed
> **Change**: validate-structure
> **Date**: 2026-06-15
> **Artifact store**: hybrid (engram + openspec/)

---

## Intent

Migrate the legacy `StructureValidator` and `StructureAnalyzer` from `business_logic/` into the hexagonal architecture as the first complete use-case slice: a domain service (`StructureValidator` in `src/domain/structure/`), an application use case (`ValidateStructureUseCase` in `src/application/`), a production wiring, and a full test suite.

This slice is designated "pure" — it depends only on domain-layer primitives (DTOs and enums already migrated in Slice 0) and no infrastructure ports. Its primary purpose is to prove the service + use case + wiring + test pattern before slices that require adapters.

**Why now**: Slice 0 (DTOs/enums) and Slice 1 (domain exceptions) are complete. The domain layer now has everything `StructureValidator` needs as inputs and outputs. Blocking further slice work on this would stall the entire migration roadmap.

**Success**: The 10 existing behavioral tests in `tests/test_structure_validator.py` pass against the new domain service. The use case correctly encapsulates the two post-processing mutations currently leaking into `main.py`. No legacy file is deleted or modified.

---

## Scope

### In scope

1. **`src/domain/structure/`** — create the package:
   - `__init__.py`
   - `structure_validator.py` — merged `StructureValidator` domain service (absorbs `StructureAnalyzer` functionality as an internal method)
   - `required_sections_provider.py` — pure domain class that maps `ArticleType` → `list[str]` required section names; replaces the hardcoded `REQUIRED_SECTIONS` config import

2. **`src/application/validate_structure_use_case.py`** — `ValidateStructureUseCase` class:
   - `execute(document_content: DocumentContent, article_type: ArticleType, has_references: bool = False) -> StructureValidationResult`
   - Encapsulates the "auto-remove Referencias if `has_references`" business rule (currently lines 228–232 of `main.py`)
   - "Desarrollo" exclusion is encoded in `RequiredSectionsProvider` — `DESARROLLO` is never listed as required for any `ArticleType` (no explicit use-case logic needed)

3. **`src/infrastructure/wirings/validate_structure_wiring.py`** — production wiring factory: instantiates `StructureValidator`, constructs `ValidateStructureUseCase`; no adapters needed

4. **`src/domain/tests/structure/`** — create the package:
   - `__init__.py`
   - `test_structure_validator.py` — `unittest.TestCase` port of `tests/test_structure_validator.py` (10 tests minimum, covering all `ArticleType` values, alias detection, long-paragraph exclusion, English aliases)

5. **`src/application/tests/`** — scaffold the package (does not currently exist):
   - `__init__.py`
   - `test_validate_structure_use_case.py` — tests for the use case covering `has_references` branching, empty-document guard, and behavioral parity with post-processed results

### Explicitly out of scope

- Deleting or modifying any file under `business_logic/` — legacy code stays untouched
- Modifying `main.py` — the new use case is NOT wired into the main application caller in this slice
- Creating `src/domain/exceptions/structure_errors.py` — `DocumentEmpty` from `document_errors.py` covers the empty-input guard
- Implementing infrastructure adapters or ports — this slice has zero infrastructure dependencies
- Using `DocumentContent.sections: dict[str, str]` instead of `paragraphs: list[str]` — the validator continues to scan raw paragraphs (exact legacy behavior)
- Migrating `StructureAnalyzer.analyze()` as a standalone public method — it is absorbed as an internal helper; its public API disappears (POO-total convention)
- The `_get_required_sections(category: ClassificationCategory)` dead-code path — not ported

---

## Behavioral Contracts

### Header detection algorithm

Use the **100-character threshold** from legacy `StructureValidator._extract_present_sections()`:

- A paragraph is treated as a section header if `len(paragraph) < 100` OR it matches the `keyword:` inline pattern.
- The 5-word filter from `StructureAnalyzer` (1–5 words) is NOT used for the main validation path.
- `StructureAnalyzer`'s IMRyD signal logic (short paragraphs only) may be retained as an internal helper for the merged service if needed for future use cases, but does not affect `validate_structure` output.

### Required sections per ArticleType

`RequiredSectionsProvider` encodes this as pure domain knowledge (no config import):

| ArticleType | Required sections |
|-------------|-------------------|
| CIENTIFICO  | Standard IMRyD + Referencias |
| DIVULGACION | Introduction-type, body, conclusion-type |
| OPINION     | Minimal set |
| UNKNOWN     | Empty list (no sections required) |

**Domain invariant**: `DESARROLLO` is **never** in any required list for any `ArticleType`.

### has_references rule (use-case responsibility)

`ValidateStructureUseCase.execute(…, has_references: bool = False)`:

- If `has_references is True`: remove "Referencias" from `missing_sections` before constructing the `StructureValidationResult`
- `StructureValidator` domain service does NOT receive or know about `has_references`

### Empty document guard

If `document_content.paragraphs` is empty, the use case raises `DocumentEmpty` (from `src/domain/exceptions/document_errors.py`) before delegating to the domain service.

### Frozen DTO constraint

`StructureValidationResult` is `frozen=True`. The use case builds the final result (with post-processing applied) and constructs the DTO once — it never mutates it after construction.

---

## Approach and Rationale

### Why merge StructureAnalyzer into StructureValidator

The plan mandates this (§4.4). Both classes scan `DocumentContent.paragraphs` for section signals. Keeping them separate would require the use case to call two services and merge results. Absorbing `StructureAnalyzer` as an internal method of `StructureValidator` preserves encapsulation without exposing IMRyD signals through the primary use case output (which is typed as `StructureValidationResult`).

### Why RequiredSectionsProvider is a separate class (not a dict in the service)

The migration plan (§4.2) explicitly calls this out as a named domain service. It makes the required-section sets testable in isolation and removes the `config.py` coupling. For this slice, it can be a simple class with a class method (no instantiation needed); future slices could inject it as a dependency if variant behavior is needed.

### Why wiring is included despite being trivial

The plan's Definition of Done requires production wiring. Even though there are no adapters, having the wiring factory ensures the pattern is established for slices that DO have adapters. It also prevents test code from constructing use cases directly.

### Why src/application/tests/ is scaffolded here

This is the first use case in `src/application/`. The directory does not exist. Scaffolding it as part of this slice avoids test-discovery gaps and establishes the convention for all subsequent application-layer tests.

---

## Files Affected

### Created

| Path | Description |
|------|-------------|
| `src/domain/structure/__init__.py` | Package marker |
| `src/domain/structure/structure_validator.py` | Domain service (merged validator + analyzer) |
| `src/domain/structure/required_sections_provider.py` | Pure domain class: ArticleType → required sections |
| `src/application/validate_structure_use_case.py` | Use case with has_references rule |
| `src/infrastructure/wirings/__init__.py` | Package marker (if missing) |
| `src/infrastructure/wirings/validate_structure_wiring.py` | Production wiring factory |
| `src/domain/tests/structure/__init__.py` | Package marker |
| `src/domain/tests/structure/test_structure_validator.py` | Domain service tests (10+ cases) |
| `src/application/tests/__init__.py` | Package marker |
| `src/application/tests/test_validate_structure_use_case.py` | Use case tests |

### Not touched

| Path | Reason |
|------|--------|
| `business_logic/structure_validator.py` | Legacy — no deletion in this slice |
| `business_logic/structure_analyzer.py` | Legacy — no deletion in this slice |
| `main.py` | Not wired into the new use case yet |
| `src/domain/exceptions/` | No new exception file needed |

---

## Dependencies

- **Slice 0 DTOs**: `DocumentContent`, `ArticleType`, `StructureValidationResult` — all in `src/domain/dtos/` and `src/domain/enums/`
- **Slice 1 exceptions**: `DocumentEmpty` from `src/domain/exceptions/document_errors.py`
- **No new external dependencies**

---

## Non-Goals

- Replacing the `main.py` caller of `StructureValidator`
- Deleting legacy code
- Implementing `StructureAnalyzer` IMRyD signals as a public use case output
- Creating new domain exception types for structure validation failures
- Introducing infrastructure ports or adapters

---

## Risks

1. **section_map alias coverage**: The hardcoded alias dict in `_extract_present_sections` must be ported exactly or the 10 legacy tests will fail on alias detection cases. The spec must enumerate all aliases explicitly.
2. **Frozen DTO construction order**: The use case must apply all post-processing (has_references removal) before constructing `StructureValidationResult`. If the order is wrong, tests that check `missing_sections` content will fail.
3. **StructureAnalyzer merger shape**: The plan states "merge" but does not specify whether the IMRyD signals appear in the use case output. This proposal resolves it as "internal helper only" — the output type remains `StructureValidationResult`. If the spec phase finds a business need to surface IMRyD signals, that becomes a separate use case.
