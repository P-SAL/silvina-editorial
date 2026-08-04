# Delta Spec — validate-structure

**Change**: resolve-magic-values-debt

## MODIFIED Requirements

### 1.2 Header Detection Rule

A paragraph from `DocumentContent.paragraphs` qualifies as a section header when EITHER of the following conditions is true:

1. `len(paragraph) < max_header_length` (short-header rule), OR
2. The paragraph (lowercased, stripped) starts with `{keyword}:` or `{keyword} :` for any keyword in any alias list in `section_map` (inline-header rule).

`StructureValidator` MUST accept an optional `max_header_length: int` in its constructor (defaulting to 100). Paragraphs of length >= `max_header_length` that do NOT match the inline-header pattern are NEVER treated as section headers, even if they contain section keywords in body text.

When a paragraph qualifies as a header, the service searches it (case-insensitively) for any keyword from each `section_map` entry. The first matching entry's key is recorded as the detected canonical section name (capitalized first letter).

(Previously: The short-header threshold was hardcoded to 100 characters.)

#### Scenario: S-10: Long body paragraph is NOT detected as section header
- GIVEN a paragraph of length >= `max_header_length` (default: 100)
- AND it does NOT start with any keyword followed by ":"
- WHEN `_extract_present_sections` is called with this paragraph
- THEN result is []

#### Scenario: S-11: Short paragraph containing section keyword IS detected
- GIVEN a paragraph of length < `max_header_length` (default: 100)
- WHEN `_extract_present_sections` is called
- THEN "Introducción" is in the returned list

#### Scenario: Custom max_header_length threshold
- GIVEN a `StructureValidator` constructed with `max_header_length` = 50
- AND a paragraph of length 60: "Introducción de nuevas metodologías en el campo de..."
- WHEN `_extract_present_sections` is called with this paragraph
- THEN result is []

---

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

(Previously: `_get_structure_validator()` instantiated `StructureValidator()` with no parameters.)

#### Scenario: S-19: Wiring creates the domain service without errors
- GIVEN `AnalyzeDocumentUseCaseWiring()._get_structure_validator()` is called
- THEN the return value is an instance of `StructureValidator`
- AND the instance is configured with the value from environment variable `STRUCTURE_MAX_HEADER_LENGTH` (defaulting to 100)
- AND no external adapters, ports, or infrastructure classes are imported
