# Delta Spec — validate-structure

**Change**: validate-structure
**Phase**: spec
**Date**: 2026-06-15
**Status**: active

---

## Scope

What must be true after this change is applied. This spec covers the behavioral contracts for
the domain service, use case, wiring factory, and test suites introduced by the validate-structure
slice. It does NOT describe implementation details — only observable behavior.

---

## 1. Domain Service — StructureValidator

### 1.1 Section Alias Map (behavioral-critical)

`StructureValidator` MUST use the following `_SECTION_ALIASES` dict verbatim for header detection.
Keys are `SectionName` enum members; values are lists of lowercase detection strings.
Every key-value pair is required; adding or removing aliases breaks legacy test parity.

```python
_SECTION_ALIASES: dict[SectionName, list[str]] = {
    SectionName.SUMMARY:       ['resumen', 'abstract'],
    SectionName.INTRODUCTION:  ['introducción', 'introduccion', 'introduction'],
    SectionName.METHODOLOGY:   ['metodología', 'metodologia', 'methodology'],
    SectionName.RESULTS:       ['resultados', 'results'],
    SectionName.DISCUSSION:    ['discusión', 'discusion', 'discussion'],
    SectionName.ARGUMENTATION: ['argumentación', 'argumentacion', 'argumentation'],
    SectionName.DEVELOPMENT:   ['desarrollo', 'development'],
    SectionName.CONCLUSIONS:   ['conclusiones', 'conclusión', 'conclusion'],
    SectionName.REFERENCES:    ['referencias', 'bibliografía', 'bibliografia', 'fuentes bibliográficas'],
}
```

### 1.2 Header Detection Rule

A paragraph from `DocumentContent.paragraphs` qualifies as a section header when EITHER of the
following conditions is true:

1. `len(paragraph) < max_header_length` (short-header rule), OR
2. The paragraph (lowercased, stripped) starts with `{keyword}:` or `{keyword} :`
   for any keyword in any alias list in `section_map` (inline-header rule).

`StructureValidator` MUST accept an optional `max_header_length: int` in its constructor (defaulting to 100).
Paragraphs of length >= `max_header_length` that do NOT match the inline-header pattern are NEVER treated as
section headers, even if they contain section keywords in body text.

When a paragraph qualifies as a header, the service searches it (case-insensitively) for
any keyword from each `section_map` entry. The first matching entry's key is recorded as the
detected canonical section name (capitalized first letter).

### 1.3 Required Sections per ArticleType

`RequiredSectionsProvider` encodes the following as authoritative, pure domain knowledge.
Returns `list[SectionName]` — enum members, not raw strings.

| ArticleType        | Required sections (ordered)                                                              |
|--------------------|------------------------------------------------------------------------------------------|
| `CIENTIFICO`       | `[SUMMARY, INTRODUCTION, METHODOLOGY, RESULTS, DISCUSSION, CONCLUSIONS, REFERENCES]` (7) |
| `DIVULGACION`      | `[SUMMARY, INTRODUCTION, DEVELOPMENT, CONCLUSIONS, REFERENCES]` (5)                      |
| `OPINION`          | `[INTRODUCTION, ARGUMENTATION, CONCLUSIONS]` (3)                                         |
| `UNKNOWN`          | `[]` (empty — no sections required)                                                       |

**Domain invariant**: `SectionName.DEVELOPMENT` MUST NOT appear in the required sections for
`CIENTIFICO` or `OPINION`. It IS included for `DIVULGACION` (faithful port of legacy).
The use case removes it unconditionally from `missing_sections` (port of `main.py:230`).

### 1.4 Missing Sections Calculation

`validate(document_content: DocumentContent, article_type: ArticleType) -> tuple[list[SectionName], list[SectionName]]`

- Calls `_extract_present_sections(paragraphs)` → `list[SectionName]` of detected canonical sections.
- Compares against required sections (case-insensitive via `.lower()`): a section is missing if
  its lowercased value does not appear in the lowercased present list.
- Returns `(present_sections, missing_sections)` as a raw tuple.
- Does NOT apply `has_references` filtering — that is use-case responsibility.
- Does NOT build `StructureValidationResult` — that is also use-case responsibility.

### 1.5 Output

Returns `tuple[list[SectionName], list[SectionName]]`:
- `present_sections` — enum members detected in the document (unused by current use case: `_`)
- `missing_sections` — enum members required but not detected

---

## 2. Orchestration — AnalyzeDocumentUseCase._validate_structure()

> **Superseded (2026-07-04, `refactor_analyze_document_wiring`)**: `ValidateStructureUseCase`
> and `ValidateStructureWiring` were eliminated as redundant pass-through layers. This
> orchestration now lives in `AnalyzeDocumentUseCase._validate_structure()`, and the wiring
> in `AnalyzeDocumentUseCaseWiring._get_structure_validator()` — see
> `openspec/specs/analyze-document/spec.md`.

### 2.1 Signature

```python
def _validate_structure(
    self,
    document_content: DocumentContentDTO,
    article_type: ArticleType,
    has_references: bool,
) -> StructureValidationResultDTO
```

### 2.2 Empty Document Guard

If `document_content.paragraphs` is empty (i.e., `len(document_content.paragraphs) == 0`),
`_validate_structure` MUST raise `DocumentEmpty` (from `src.domain.exceptions.document_errors`)
BEFORE delegating to the domain service.

### 2.3 Post-Processing Rules (applied before DTO construction)

After `StructureValidator.validate()` returns `(_, missing)`, `_validate_structure` applies
these rules in order before building the frozen DTO:

1. **Development removal (unconditional)**: Always remove `SectionName.DEVELOPMENT` from `missing`.
   Source: `main.py` line 230 — legacy removes it unconditionally after every call.
2. **References removal (conditional)**: If `has_references is True`, also remove `SectionName.REFERENCES`
   from `missing`.

The `present_sections` return value from `validate()` is discarded (`_`) — the orchestrator
only operates on `missing_sections`.

Post-processing is applied BEFORE the final DTO is constructed. The frozen DTO is built once
and never mutated after construction.

### 2.4 Resulting is_valid Recomputation

After post-processing:
- `is_valid = len(filtered_missing_sections) == 0`

`is_valid` MUST reflect the filtered list, not the raw domain service result.

### 2.5 Isolation

`_validate_structure` MUST NOT accept or forward `has_references` to `StructureValidator.validate()`.
The domain service has no knowledge of the `has_references` business rule.

---

## 3. Wiring — AnalyzeDocumentUseCaseWiring._get_structure_validator()

### 3.1 Factory Behavior

`AnalyzeDocumentUseCaseWiring` MUST:
- Create the `StructureValidator` dependency in its own private `_get_structure_validator(self) -> StructureValidator` method.
- Inject it into `AnalyzeDocumentUseCase`'s constructor (not constructed inline inside `execute()`).
- Load the maximum header length from the environment variable `STRUCTURE_MAX_HEADER_LENGTH` (defaulting to 100) and pass it to `StructureValidator`.

Pattern:
```python
class AnalyzeDocumentUseCaseWiring:
    def create_use_case(self) -> AnalyzeDocumentUseCase:
        return AnalyzeDocumentUseCase(
            ...,
            structure_validator=self._get_structure_validator(),
            ...,
        )

    def _get_structure_validator(self) -> StructureValidator:
        max_header_length = int(os.environ.get("STRUCTURE_MAX_HEADER_LENGTH", 100))
        return StructureValidator(max_header_length=max_header_length)
```

### 3.2 No Infrastructure Dependencies

`_get_structure_validator()` MUST NOT import from `src.infrastructure.adapters`, any port
interface, or any external library. `StructureValidator` is a pure domain-layer class.

---

## 4. Package Structure

The following packages and files MUST exist after the change is applied:

```
src/
  domain/
    structure/
      __init__.py
      structure_validator.py
      required_sections_provider.py
    tests/
      structure/
        __init__.py
        test_structure_validator.py
  application/
    analyze_document_use_case.py    # _validate_structure() lives here
  infrastructure/
    wirings/
      __init__.py
      analyze_document_use_case_wiring.py    # _get_structure_validator() lives here
```

---

## 5. Acceptance Scenarios

### S-01: Scientific article with all sections — valid

```
Given a DocumentContent with paragraphs:
  ["Resumen: El resumen.", "Introducción: Intro.", "Metodología: Método.",
   "Resultados: Los resultados.", "Discusión: Discusión.", "Conclusiones: Conclusión.",
   "Referencias: Refs."]
And article_type = ArticleType.CIENTIFICO
When validate is called on the domain service
Then result.is_valid is True
And result.missing_sections == []
```

### S-02: Scientific article with all sections (inline colon format) — valid

```
Given a DocumentContent with paragraphs using "Section: content" format
  where all 7 CIENTIFICO sections appear as short inline-header paragraphs
And article_type = ArticleType.CIENTIFICO
When validate is called
Then result.is_valid is True
And result.missing_sections == []
```

### S-03: Scientific article missing Resumen — invalid

```
Given a DocumentContent without any "Resumen" or "Abstract" paragraph
And article_type = ArticleType.CIENTIFICO
When validate is called
Then result.is_valid is False
And "Resumen" is in result.missing_sections
```

### S-04: Divulgacion article with all required sections — valid

```
Given a DocumentContent with paragraphs:
  ["Resumen: ...", "Introducción: ...", "Desarrollo: ...", "Conclusiones: ...", "Referencias: ..."]
And article_type = ArticleType.DIVULGACION
When validate is called
Then result.is_valid is True
And result.missing_sections == []
```

### S-05: Divulgacion article missing Desarrollo — invalid

```
Given a DocumentContent without any "Desarrollo" or "Development" paragraph
And article_type = ArticleType.DIVULGACION
When validate is called
Then result.is_valid is False
And "Desarrollo" is in result.missing_sections
```

### S-06: Opinion article with all required sections — valid

```
Given a DocumentContent with paragraphs:
  ["Introducción: ...", "Argumentación: ...", "Conclusiones: ..."]
And article_type = ArticleType.OPINION
When validate is called
Then result.is_valid is True
And result.missing_sections == []
```

### S-07: validate_structure returns StructureValidationResult object

```
Given a DocumentContent with any paragraphs
And article_type = ArticleType.DIVULGACION
When validate is called on the domain service
Then the return value has attribute is_valid (bool)
And the return value has attribute missing_sections (list)
```

### S-08: English alias "abstract" maps to canonical "Resumen"

```
Given a DocumentContent with paragraph "Abstract: This is the abstract."
When _extract_present_sections is called
Then "resumen" appears in the lowercased list of detected section names
```

### S-09: Multiple aliases resolved correctly (metodologia, methodology, discussion, results)

```
Given a DocumentContent with paragraphs:
  ["metodologia: Methods without accent.",
   "methodology: English version of methods.",
   "discussion: English version of discussion.",
   "results: English results."]
When _extract_present_sections is called
Then "metodología" appears in the lowercased detected sections
And "discusión" appears in the lowercased detected sections
And "resultados" appears in the lowercased detected sections
```

### S-10: Long body paragraph (>= 100 chars) is NOT detected as section header

```
Given a paragraph: "La introducción de nuevas metodologías en el campo de la investigación
  académica requiere un análisis." (len >= 100)
And it does NOT start with any keyword followed by ":"
When _extract_present_sections is called with this paragraph
Then result is []
```

### S-11: Short paragraph (< 100 chars) containing section keyword IS detected

```
Given a paragraph: "Introducción" (len < 100)
When _extract_present_sections is called
Then "Introducción" is in the returned list
```

### S-11b: Custom max_header_length threshold

```
Given a StructureValidator constructed with max_header_length = 50
And a paragraph of length 60: "Introducción de nuevas metodologías en el campo de..."
When _extract_present_sections is called with this paragraph
Then result is []
```

### S-12: Orchestrator raises DocumentEmpty on empty paragraphs list

```
Given a DocumentContentDTO where paragraphs == []
And article_type = ArticleType.CIENTIFICO
When AnalyzeDocumentUseCase._validate_structure() is called
Then DocumentEmpty is raised
And StructureValidator.validate is NOT called
```

### S-13: Orchestrator always removes "Desarrollo" from missing_sections

```
Given a DocumentContentDTO with paragraphs that contain all DIVULGACION sections
  EXCEPT "Desarrollo"
And article_type = ArticleType.DIVULGACION
And has_references = False
When AnalyzeDocumentUseCase._validate_structure() is called
Then "Desarrollo" is NOT in result.missing_sections
And result.is_valid is True
```

### S-14: Orchestrator with has_references=True removes "Referencias" from missing_sections

```
Given a DocumentContentDTO with paragraphs that contain all CIENTIFICO sections
  EXCEPT "Referencias"
And article_type = ArticleType.CIENTIFICO
And has_references = True
When AnalyzeDocumentUseCase._validate_structure() is called
Then "Referencias" is NOT in result.missing_sections
And result.is_valid is True
```

### S-15: Orchestrator with has_references=False does NOT remove "Referencias"

```
Given a DocumentContentDTO with paragraphs that contain all CIENTIFICO sections
  EXCEPT "Referencias"
And article_type = ArticleType.CIENTIFICO
And has_references = False
When AnalyzeDocumentUseCase._validate_structure() is called
Then "Referencias" IS in result.missing_sections
And result.is_valid is False
```

### S-15b: Domain service DOES report "Desarrollo" as missing (before orchestrator post-processing)

```
Given a DocumentContentDTO without any "Desarrollo" or "Development" paragraph
And article_type = ArticleType.DIVULGACION
When validate is called DIRECTLY on the domain service
Then "Desarrollo" IS in result.missing_sections
```

### S-16: UNKNOWN article type — no required sections, always valid

```
Given a DocumentContentDTO with any paragraphs (non-empty)
And article_type = ArticleType.UNKNOWN
When AnalyzeDocumentUseCase._validate_structure() is called
Then result.is_valid is True
And result.missing_sections == []
```

### S-16: DESARROLLO not required for CIENTIFICO

```
Given RequiredSectionsProvider.get_required(ArticleType.CIENTIFICO)
Then "Desarrollo" is NOT in the returned list
```

### S-17: DESARROLLO not required for OPINION

```
Given RequiredSectionsProvider.get_required(ArticleType.OPINION)
Then "Desarrollo" is NOT in the returned list
```

### S-18: "fuentes bibliográficas" alias maps to "Referencias"

```
Given a DocumentContent with paragraph "Fuentes bibliográficas: Autor, 2023."
When _extract_present_sections is called
Then "referencias" appears in the lowercased detected sections
```

### S-19: Wiring creates the domain service without errors

```
Given AnalyzeDocumentUseCaseWiring()._get_structure_validator() is called
Then the return value is an instance of StructureValidator
And the instance is configured with the value from environment variable STRUCTURE_MAX_HEADER_LENGTH (defaulting to 100)
And no external adapters, ports, or infrastructure classes are imported
```

### S-20: StructureValidationResult is frozen (immutable after construction)

```
Given a StructureValidationResult returned by validate_structure
When an attempt is made to set any attribute on the result
Then a FrozenInstanceError (or equivalent dataclass frozen error) is raised
```

---

## 6. Legacy Test Coverage Mapping

The 10 tests in `tests/test_structure_validator.py` map to the following acceptance scenarios:

| Legacy test method                                  | Scenario(s) |
|-----------------------------------------------------|-------------|
| `test_scientific_article_detects_all_imryd_sections` | S-01, S-02  |
| `test_scientific_article_complete_is_valid`          | S-01        |
| `test_scientific_article_missing_resumen`            | S-03        |
| `test_divulgacion_article_compliant`                 | S-04        |
| `test_divulgacion_missing_desarrollo`                | S-05        |
| `test_opinion_article_complete_is_valid`             | S-06        |
| `test_validate_structure_returns_result_object`      | S-07        |
| `test_english_abstract_detected`                     | S-08        |
| `test_section_aliases_detected`                      | S-09        |
| `test_long_body_text_not_detected_as_section`        | S-10        |
| `test_short_section_header_is_detected`              | S-11        |

All 10 legacy tests MUST pass against the new domain service.
(The table lists 11 rows because S-01 covers two legacy tests.)

---

## 7. Constraints and Invariants

1. `_extract_present_sections` is a non-public method but MUST remain testable directly
   (legacy tests call it directly; it MUST NOT be renamed or made truly private via name mangling).
2. Comparison between detected sections and required sections is case-insensitive.
3. The domain service is stateless — no instance state is mutated across calls.
4. `StructureValidationResult` MUST be frozen (`frozen=True` dataclass). Use-case builds it once.
5. No new domain exception types are introduced. `DocumentEmpty` from `document_errors.py` is sufficient.
6. `StructureAnalyzer` is absorbed as internal logic. Its public `analyze()` method MUST NOT appear
   on the new domain service's public interface.
7. Legacy files (`business_logic/structure_validator.py`, `business_logic/structure_analyzer.py`,
   `main.py`) are NOT modified or deleted in this slice.
8. The 5-word filter from `StructureAnalyzer` is NOT used in the `validate_structure` path.
   Only the 100-character threshold applies.

---

## 8. Out of Scope

- Wiring the new use case into `main.py`
- Deleting any file under `business_logic/`
- Surfacing IMRyD signals (`has_introduction`, `has_methods`, etc.) in `StructureValidationResult`
- Creating `structure_errors.py` or any new exception type
- Infrastructure ports or adapters
- Using `DocumentContent.sections` dict instead of `paragraphs` list
- The `_get_required_sections(category: ClassificationCategory)` dead-code path in legacy

---

## 9. Input/Output Contracts

### Domain service inputs
- `document_content: DocumentContent` — scans `.paragraphs: list[str]`
- `article_type: ArticleType` — selects required section list

### Domain service output
- `StructureValidationResult(is_valid, missing_sections, section_details={}, timestamp=<auto>)`

### Orchestrator (`_validate_structure`) inputs
- `document_content: DocumentContentDTO`
- `article_type: ArticleType`
- `has_references: bool`

### Orchestrator (`_validate_structure`) output
- `StructureValidationResultDTO` with post-processing applied

### Exception
- `DocumentEmpty` raised when `document_content.paragraphs == []`
