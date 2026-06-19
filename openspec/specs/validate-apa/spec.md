# Spec — validate-apa (Slice 3)

**Status**: approved
**Change**: validate-apa
**Project**: silvina-editorial
**Date**: 2026-06-15

---

## 1. ApaErrorType Enum

**File**: `src/domain/enums/apa_error_type.py`
**Class name**: `ApaErrorType` (PascalCase; legacy name was `APAErrorType`)

### 1.1 Members (all 8 must be present)

| Member name            | String value                        | Used in validation? |
|------------------------|-------------------------------------|---------------------|
| `CONJUNCTION_ERROR`    | `"Conjunción incorrecta"`           | Yes                 |
| `COMMA_ERROR`          | `"Puntuación incorrecta"`           | Yes                 |
| `CAPITALIZATION_ERROR` | `"Mayúsculas/minúsculas incorrectas"` | Yes               |
| `ET_AL_FORMAT_ERROR`   | `"Formato 'et al.' incorrecto"`     | Yes                 |
| `PAGE_FORMAT_ERROR`    | `"Formato de página incorrecto"`    | Yes                 |
| `SPACING_ERROR`        | `"Espaciado incorrecto"`            | Yes                 |
| `YEAR_FORMAT_ERROR`    | `"Formato de año incorrecto"`       | **Unused — preserve** |
| `PARENTHESES_ERROR`    | `"Paréntesis incorrectos"`          | **Unused — preserve** |

**Constraint**: `YEAR_FORMAT_ERROR` and `PARENTHESES_ERROR` MUST be declared in the enum even though no check currently produces them. Dropping them is a breaking change for downstream consumers.

---

## 2. ApaViolation DTO

**File**: `src/domain/dtos/apa_violation_dto.py`
**Class name**: `ApaViolation`

### 2.1 Fields

| Field              | Type           | Default | Notes                              |
|--------------------|----------------|---------|------------------------------------|
| `citation_text`    | `str`          | —       | The raw citation string            |
| `error_type`       | `ApaErrorType` | —       | Enum member                        |
| `location`         | `int`          | —       | Paragraph index (0-based)          |
| `explanation`      | `str`          | —       | Human-readable problem description |
| `correction`       | `str`          | —       | Suggested corrected citation       |
| `paragraph_preview`| `str`          | `""`    | First 30 chars of paragraph + `"..."` |

### 2.2 Constraints

- Decorated with `@dataclass(frozen=True)`.
- All fields are immutable value types (str / int / enum); no mutable defaults.
- No methods beyond those provided by `@dataclass`.
- `generate_report()` is NOT present on this DTO (out of scope for this slice).

---

## 3. ApaValidationResult DTO

**File**: `src/domain/dtos/apa_validation_result_dto.py`
**Class name**: `ApaValidationResult`

### 3.1 Fields

| Field             | Type                  | Notes                                    |
|-------------------|-----------------------|------------------------------------------|
| `is_valid`        | `bool`                | `True` iff `violation_count == 0`        |
| `violation_count` | `int`                 | Total number of violations               |
| `violations`      | `list[ApaViolation]`  | Ordered list (same order as input)       |

### 3.2 Constraints

- Decorated with `@dataclass(frozen=True)`.
- `is_valid` MUST equal `(violation_count == 0)`. Callers must not construct an instance where these are inconsistent.
- `violations` is a plain `list` (not tuple) but the DTO itself is frozen (the list reference is immutable, not the list contents; this is the same pattern used by `StructureValidationResult`).
- No `generate_report()` method.

---

## 4. ApaValidator Domain Service — Behavioral Contract

**File**: `src/domain/citation/apa_validator.py`
**Class name**: `ApaValidator`
**Nature**: stateless — no instance state between calls.

### 4.0 Constructor

```
ApaValidator.__init__(self) -> None
```

- MUST NOT declare `self.violations` or any other mutable accumulator field.
- MAY declare `self.rules` dict (read-only constants):
  `{'conjunction': 'y', 'et_al': 'et al.', 'page_single': 'p.', 'page_multiple': 'pp.'}`

### 4.1 Non-Author Skip Patterns

Applied inside `_validate_parenthetical()` AFTER Check 1 (conjunction) and BEFORE Check 2
(comma). When any pattern matches, the method returns immediately with only the violations
already collected (i.e., conjunction errors may still be reported for non-author citations).

The following 7 regex patterns MUST be preserved verbatim:

```python
_NON_AUTHOR_PATTERNS = [
    r'^\([A-Z]{2,}\s+\d',           # (PLANCAMIL 2023), (UNESCO 2020)
    r'^\(arXiv:',                    # (arXiv:2404.19573)
    r'^\(doi:',                      # (doi:10.1234)
    r'^\(repositorio',               # (repositorio trazable...)
    r'^\(no hay',                    # (no hay dataset...)
    r'^\([a-záéíóúñ].*\d{4}.*\d{4}',# (años 2024, 2025...) date ranges
    r'^\(\w+\s+\w+.*\d{4}.*\d{4}',  # multi-word + two years
]
```

Matching uses `re.search(pattern, citation, re.IGNORECASE)`.

### 4.2 Parenthetical Citation Rules — `_validate_parenthetical()`

Input form: `(Author, Year)` — starts with `(` and ends with `)`.

Checks execute in order; each is independent unless noted:

| # | Check | Trigger condition | Error type | Correction |
|---|-------|-------------------|------------|------------|
| 1 | Ampersand conjunction | `' & '` in inner text (between outer parens) | `CONJUNCTION_ERROR` | Replace ` & ` with ` y ` |
| — | Non-author skip | Any `_NON_AUTHOR_PATTERNS` matches full citation string | — | Return early; skip checks 2–6 |
| 2 | Missing comma before year | `re.match(r'\(([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)\s+(\d{4}[a-z]?)\)', citation)` AND `','` not in citation | `COMMA_ERROR` | `(Author, Year)` |
| 3 | Lowercase author | `re.search(r'\(([a-záéíóúñ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)', citation)` matches | `CAPITALIZATION_ERROR` | `citation.capitalize()` |
| 4a | Et al. with extra period | `'et. al'` in inner text (and `r'\bet\.?\s+al\b'` also matches) | `ET_AL_FORMAT_ERROR` | `re.sub(r'et\.\s+al', 'et al', citation)` |
| 4b | Et al. missing period after "al" | `r'\bet\.?\s+al\b'` matches AND `'et. al'` NOT in inner AND `re.search(r'et al[,\)]', inner)` matches | `ET_AL_FORMAT_ERROR` | `re.sub(r'et al\b(?!\.)','et al.', citation)` |
| 5 | Spanish page abbreviation | `'pág'` or `'página'` in `inner.lower()` | `PAGE_FORMAT_ERROR` | `.replace('pág.','p.').replace('págs.','pp.')` |
| 6 | Double space | `'  '` (two spaces) in full citation string | `SPACING_ERROR` | `' '.join(citation.split())` |

**Note on checks 4a/4b**: they are mutually exclusive. 4a fires when `et. al` is present (period after "et"). 4b fires when `et al` appears without a trailing period (e.g., `et al,` or `et al)`).

### 4.3 Narrative Citation Rules — `_validate_narrative()`

Input form: `Author (Year)` — does NOT start with `(`.

First, attempt to match:
```python
pattern = r'([A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-\s]+(?:et al\.)?)[\s]*\((\d{4}[a-z]?)\)'
```

If the pattern does NOT match → return `[]` (cannot validate; silently skip).

When pattern matches, `author_part = match.group(1).strip()`, `year_part = match.group(2)`:

| # | Check | Trigger condition | Error type | Correction |
|---|-------|-------------------|------------|------------|
| 1 | Ampersand in author part | `' & '` in `author_part` | `CONJUNCTION_ERROR` | Replace ` & ` with ` y ` |
| 2 | Et al. with extra period | `'et. al'` in `author_part` (AND `'et al'` in `author_part.lower()`) | `ET_AL_FORMAT_ERROR` | `.replace('et. al', 'et al')` |
| 3 | Missing space before year | `re.search(r'\s\(\d{4}[a-z]?\)', citation)` does NOT match | `SPACING_ERROR` | `re.sub(r'([A-Za-z])\(', r'\1 (', citation)` |

**Note**: `García & Pérez (2020)` — the ampersand causes the narrative regex to NOT match (the pattern expects only word-chars and spaces in the author part), so NO violation is reported. This is documented legacy behavior; the spec preserves it. (See S-08 for a matched narrative ampersand case.)

### 4.4 `validate_citation()` Behavior

```python
def validate_citation(
    self,
    citation_text: str,
    paragraph_index: int,
    paragraph_text: str = ""
) -> list[ApaViolation]:
```

- Determines citation type: `is_parenthetical = citation_text.startswith('(') and citation_text.endswith(')')`.
- Computes `preview = paragraph_text[:30] + "..." if len(paragraph_text) > 30 else paragraph_text`.
- Delegates to `_validate_parenthetical()` or `_validate_narrative()`.
- Returns the list from the delegate directly (no accumulation into `self`).
- Always returns a `list` (never `None`).
- An empty list means the citation is compliant (or was skipped as non-author).

### 4.5 `validate_all_citations()` Behavior

```python
def validate_all_citations(
    self,
    citations: list[tuple[str, int, str]]
) -> list[ApaViolation]:
```

- Accepts a list of `(citation_text, paragraph_index, paragraph_text)` 3-tuples.
- For each tuple, calls `self.validate_citation(citation_text, paragraph_index, paragraph_text)`.
- Extends an accumulator list with results (order preserved; matches input order).
- Returns the accumulated list.
- Empty input list → returns `[]`.
- No instance state is modified.

### 4.6 Absent from Domain Service

- `generate_report()` MUST NOT be present (presentation concern; deferred to Slice 13).
- `self.violations` accumulator MUST NOT be present.
- Module-level `validate_apa_citations()` function MUST NOT be present in the new module.

---

## 5. ValidateApaUseCase

**File**: `src/application/validate_apa_use_case.py`
**Class name**: `ValidateApaUseCase`

### 5.1 Constructor

```python
def __init__(self, validator: ApaValidator) -> None:
```

Stores `validator` as `self._validator` (or equivalent private attribute).

### 5.2 `execute()` Method

```python
def execute(
    self,
    citations: list[tuple[str, int, str]]
) -> ApaValidationResult:
```

**Behavior**:

1. If `citations` is empty → return `ApaValidationResult(is_valid=True, violation_count=0, violations=[])` immediately. MUST NOT raise an exception (ADR-4: domain services return empty results for empty inputs).
2. Otherwise, call `self._validator.validate_all_citations(citations)` → `violations: list[ApaViolation]`.
3. Compute `violation_count = len(violations)`.
4. Compute `is_valid = (violation_count == 0)`.
5. Return `ApaValidationResult(is_valid=is_valid, violation_count=violation_count, violations=violations)`.

**Constraints**:
- Never raises an exception for valid (possibly empty) input.
- Is a pure pass-through orchestrator; no APA logic lives here.

---

## 6. ValidateApaWiring

**File**: `src/infrastructure/wirings/validate_apa_wiring.py`
**Class name**: `ValidateApaWiring`

### 6.1 `create_use_case()` Method

```python
@classmethod
def create_use_case(cls) -> ValidateApaUseCase:
```

- Internally calls `cls._get_validator()` to obtain an `ApaValidator` instance.
- Returns `ValidateApaUseCase(validator=cls._get_validator())`.

### 6.2 `_get_validator()` Private Method

```python
@classmethod
def _get_validator(cls) -> ApaValidator:
```

- Returns a fresh `ApaValidator()` instance.
- Follows the `_get_*` naming convention established by `ValidateStructureWiring`.

---

## 7. File Locations (all new)

| Artifact                | Path                                                       |
|-------------------------|------------------------------------------------------------|
| `ApaErrorType` enum     | `src/domain/enums/apa_error_type.py`                       |
| `ApaViolation` DTO      | `src/domain/dtos/apa_violation_dto.py`                     |
| `ApaValidationResult`   | `src/domain/dtos/apa_validation_result_dto.py`             |
| Citation package init   | `src/domain/citation/__init__.py`                          |
| `ApaValidator` service  | `src/domain/citation/apa_validator.py`                     |
| `ValidateApaUseCase`    | `src/application/validate_apa_use_case.py`                 |
| `ValidateApaWiring`     | `src/infrastructure/wirings/validate_apa_wiring.py`        |
| Domain tests            | `src/domain/tests/citation/test_apa_validator.py`          |

---

## 8. Acceptance Scenarios

### S-01: Valid parenthetical citation produces no violations

```
Given the citation "(García, 2020)"
When validate_citation is called with paragraph_index=0
Then the result is an empty list
```

### S-02: Ampersand instead of "y" in parenthetical

```
Given the citation "(García & Pérez, 2020)"
When validate_citation is called
Then exactly one violation is returned
And its error_type is CONJUNCTION_ERROR
And violation.correction contains ' y '
```

### S-03: Missing comma before year in parenthetical

```
Given the citation "(García 2020)"
When validate_citation is called
Then violations contain an entry with error_type COMMA_ERROR
And violation.correction equals "(García, 2020)"
```

### S-04: Lowercase author in parenthetical

```
Given the citation "(garcía, 2020)"
When validate_citation is called
Then violations contain an entry with error_type CAPITALIZATION_ERROR
```

### S-05: Malformed et al. — extra period on "et"

```
Given the citation "(García et. al., 2020)"
When validate_citation is called
Then violations contain an entry with error_type ET_AL_FORMAT_ERROR
And violation.correction does not contain "et. al"
```

### S-05b: Malformed et al. — missing period after "al"

```
Given the citation "(García et al, 2020)"
When validate_citation is called
Then violations contain an entry with error_type ET_AL_FORMAT_ERROR
And violation.correction contains "et al."
```

### S-06: Spanish page abbreviation

```
Given the citation "(García, 2020, pág. 5)"
When validate_citation is called
Then violations contain an entry with error_type PAGE_FORMAT_ERROR
And violation.correction contains "p." and not "pág."
```

### S-07: Double space before year in parenthetical

```
Given the citation "(García,  2020)" [two spaces]
When validate_citation is called
Then violations contain an entry with error_type SPACING_ERROR
And violation.correction has single spaces only
```

### S-08: Ampersand in parseable narrative citation

```
Given a narrative citation where ' & ' appears in the author part
  AND the full string matches the narrative pattern
When validate_citation is called
Then violations contain an entry with error_type CONJUNCTION_ERROR
```

### S-09: Malformed et al. in narrative citation

```
Given the citation "García et. al. (2020)"
When validate_citation is called
Then violations contain an entry with error_type ET_AL_FORMAT_ERROR
```

### S-10: Narrative citation missing space before year

```
Given the citation "García(2020)" [no space before parenthesis]
When validate_citation is called
Then violations contain an entry with error_type SPACING_ERROR
And violation.correction matches "García (2020)"
```

### S-11: Non-author pattern is skipped — institutional acronym

```
Given the citation "(UNESCO 2020)"
When validate_citation is called
Then no CAPITALIZATION_ERROR is in the result
And no COMMA_ERROR is in the result
```

### S-11b: Non-author pattern is skipped — arXiv identifier

```
Given the citation "(arXiv:2404.19573)"
When validate_citation is called
Then the result contains no violations beyond any conjunction check
```

### S-11c: Non-author pattern is skipped — DOI

```
Given the citation "(doi:10.1234/example)"
When validate_citation is called
Then the result contains no violations
```

### S-12: Empty citation list returns is_valid=True

```
Given an empty list of citations
When ValidateApaUseCase.execute([]) is called
Then the result is ApaValidationResult(is_valid=True, violation_count=0, violations=[])
And no exception is raised
```

### S-13: Use case computes is_valid and violation_count correctly

```
Given a list containing one citation with a CONJUNCTION_ERROR
When ValidateApaUseCase.execute(citations) is called
Then result.violation_count == 1
And result.is_valid == False
And len(result.violations) == 1

Given a list containing only valid citations
When ValidateApaUseCase.execute(citations) is called
Then result.violation_count == 0
And result.is_valid == True
```

### S-14: ApaValidationResult is frozen

```
Given a result = ApaValidationResult(is_valid=True, violation_count=0, violations=[])
When caller attempts result.is_valid = False
Then FrozenInstanceError is raised
```

### S-14b: ApaViolation is frozen

```
Given v = ApaViolation(citation_text="...", error_type=..., location=0, explanation="...", correction="...")
When caller attempts v.location = 99
Then FrozenInstanceError is raised
```

### S-15: Wiring creates a valid use case

```
Given ValidateApaWiring
When create_use_case() is called
Then the returned object is an instance of ValidateApaUseCase
And calling execute([]) on the result returns ApaValidationResult(is_valid=True, ...)
```

---

## 9. Out of Scope

The following are explicitly excluded from this slice:

- `generate_report()` — presentation concern; deferred to Slice 13 (report formatter adapter).
- Wiring `ValidateApaUseCase` into `main.py` — deferred to Slice 14 (caller switchover).
- Deleting `apa_validator.py` (root) — legacy file stays alive during coexistence; `main.py` continues importing from root.
- Application-layer tests for `ValidateApaUseCase` — the use case is a pure pass-through; domain service tests are sufficient.
- `YEAR_FORMAT_ERROR` and `PARENTHESES_ERROR` validation logic — enum values must exist but no check must produce them.

---

## 10. Invariants and ADR References

| Invariant | Source |
|-----------|--------|
| Empty input never raises; returns empty result | ADR-4 |
| Domain service is stateless | Hexagonal migration plan §4.2 |
| All DTOs are frozen | Slice 0 DTO convention |
| `generate_report()` is a presentation concern | Hexagonal migration plan §10.5 |
| Legacy root file is unchanged | Coexistence policy (Slice 14) |
| Non-author skip patterns are preserved verbatim | Business rule; risk flagged in proposal |
