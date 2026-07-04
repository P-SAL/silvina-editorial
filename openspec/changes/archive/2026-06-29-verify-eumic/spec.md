# SDD Spec — verify-eumic (Slice 11)

**Change name**: verify-eumic
**Slice**: 11
**Branch**: refactor/hexagonal-migration
**Status**: Spec
**Date**: 2026-06-29

---

## 1. What This Spec Covers

This spec defines the delta state that MUST be true after Slice 11 is applied.
It describes WHAT each artifact must be, not HOW to implement it.

Slice 11 migrates the EUMIC format-compliance logic from `eumic_verifier.py` (root-level legacy
file) into the hexagonal architecture. The result is five new source files, six new test files,
and zero modifications to any existing file.

---

## 2. Requirements

### REQ-1: EumicViolationDTO

**File**: `src/domain/dtos/eumic_violation_dto.py`
**Class**: `EumicViolationDTO`

- R1.1 — `EumicViolationDTO` MUST be a frozen dataclass that inherits from `BaseDTO`
  (`src/domain/dtos/base_dto.py`).
- R1.2 — It MUST declare the following fields in order:
  - `category: str` — which EUMIC check area raised the violation
  - `message: str` — human-readable description of the violation
  - `severity: SeverityLevel` — uses existing `src/domain/enums/severity_level.py`; no new enum
  - `details: str = ""` — optional extra context; default is empty string
- R1.3 — The class MUST be immutable: any attempt to mutate a field after construction MUST raise
  `dataclasses.FrozenInstanceError`.
- R1.4 — The class MUST NOT import from `src/application/` or `src/infrastructure/` (domain
  import invariant).
- R1.5 — The `severity` field MUST accept any `SeverityLevel` enum value
  (`INFO`, `WARNING`, `ERROR`, `CRITICAL`).
- R1.6 — The file MUST contain exactly one class. File ends with exactly one blank line.

### REQ-2: DocumentFormatInspectionPort

**File**: `src/domain/document/document_format_inspection_port.py`
**Class**: `DocumentFormatInspectionPort(ABC)`

- R2.1 — `DocumentFormatInspectionPort` MUST be an abstract base class (`from abc import ABC,
  abstractmethod`).
- R2.2 — It MUST declare exactly one abstract method:

  ```
  def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]
  ```

- R2.3 — Attempting to instantiate `DocumentFormatInspectionPort()` directly MUST raise
  `TypeError`.
- R2.4 — The `inspect` signature MUST use the exact parameter names `docx_path` and
  `word_count`. Both annotations (`str`, `int`) and return annotation
  (`list[EumicViolationDTO]`) MUST be present.
- R2.5 — A class that subclasses `DocumentFormatInspectionPort` and provides a concrete
  `inspect` implementation MUST be instantiable without errors.
- R2.6 — The method MUST have a docstring following PEP 257 single-line form.
- R2.7 — The file MUST NOT import from `src/application/` or `src/infrastructure/`.

### REQ-3: VerifyEumicUseCase

**File**: `src/application/verify_eumic_use_case.py`
**Class**: `VerifyEumicUseCase`

- R3.1 — `VerifyEumicUseCase.__init__` MUST accept exactly one parameter:
  `format_inspection_port: DocumentFormatInspectionPort`.
- R3.2 — The port MUST be stored as `self._format_inspection_port` (private attribute).
- R3.3 — The class MUST expose a method `execute(self, docx_path: str, word_count: int) ->
  list[EumicViolationDTO]`.
- R3.4 — `execute` MUST be decorated with `@generic_error_handler` from
  `src/domain/exceptions/decorators/generic_error_handler.py`.
- R3.5 — `execute` MUST delegate entirely to
  `self._format_inspection_port.inspect(docx_path=docx_path, word_count=word_count)`.
- R3.6 — `execute` MUST return the port's return value unchanged (no transformation, no
  filtering).
- R3.7 — `execute` MUST NOT contain any domain computation or business logic.
- R3.8 — When the port raises a `BaseSrcError` subclass, `@generic_error_handler` MUST allow it
  to propagate as-is.
- R3.9 — When the port raises an unexpected `Exception`, `@generic_error_handler` MUST wrap it
  in `SrcGenericError` before re-raising.
- R3.10 — The file MUST NOT import from `src/infrastructure/`.

### REQ-4: DocxEumicAdapter

**File**: `src/infrastructure/adapters/document/docx_eumic_adapter.py`
**Class**: `DocxEumicAdapter(DocumentFormatInspectionPort)`

#### R4-General

- R4.1 — `DocxEumicAdapter` MUST implement `DocumentFormatInspectionPort` (concrete subclass).
- R4.2 — All imports MUST appear at the module top level. No imports inside methods or functions.
  The following imports (which are local in `eumic_verifier.py`) MUST be moved to module level:
  `from docx.shared import Pt, Cm` and `from docx.enum.text import WD_ALIGN_PARAGRAPH`.
- R4.3 — `DocxEumicAdapter` MUST be stateless: no instance-level list of violations. Each check
  method MUST return `list[EumicViolationDTO]`; `inspect()` aggregates by extending.
- R4.4 — `inspect` MUST be decorated with `@generic_error_handler`.
- R4.5 — `inspect` MUST open the docx file internally via `Document(docx_path)`. If this call
  raises any exception, the adapter MUST raise `DocumentUnreadable` (from
  `src/domain/exceptions/document_errors.py`) before `@generic_error_handler` processes it.
  `DocumentUnreadable` is a `BaseSrcError` subclass and propagates as-is through
  `@generic_error_handler`.
- R4.6 — `inspect` MUST call all five private check methods in order and return the concatenated
  list of `EumicViolationDTO` objects.
- R4.7 — When no violations are found, `inspect` MUST return an empty list `[]`.
- R4.8 — `EumicViolation` (legacy dataclass) MUST NOT appear in the adapter. Use `EumicViolationDTO`.
- R4.9 — `EumicSeverity` (legacy enum) MUST NOT appear in the adapter. Use `SeverityLevel`.

#### R4-Constants

All of the following module-level constants MUST be defined before the class declaration in
`docx_eumic_adapter.py`. No magic number or string literal encoding an EUMIC editorial rule
may appear inside any method body.

| Constant | Type | Value | Method(s) |
|---|---|---|---|
| `REQUIRED_MARGIN_CM` | `float` | `2.5` | `_verify_format` |
| `MARGIN_TOLERANCE_CM` | `float` | `0.3` | `_verify_format` |
| `ALLOWED_FONTS` | `frozenset[str]` | `frozenset({"Times New Roman", "Arial", "Calibri"})` | `_verify_format` |
| `REQUIRED_FONT_SIZE_PT` | `int` | `12` | `_verify_format` |
| `FONT_SIZE_TOLERANCE_PT` | `float` | `1.0` | `_verify_format` |
| `MAX_NON_JUSTIFIED_RATIO` | `float` | `0.3` | `_verify_format` |
| `FIGURE_CAPTION_PREFIXES` | `tuple[str, ...]` | `("figura", "fig.", "figure")` | `_verify_figures` |
| `FIGURE_NUMBERING_PATTERN` | `str` | `r"figura\s+(\d+)"` | `_verify_figures` |
| `TABLE_CAPTION_PREFIXES` | `tuple[str, ...]` | `("tabla", "table", "cuadro")` | `_verify_tables` |
| `TABLE_NUMBERING_PATTERN` | `str` | `r"tabla\s+(\d+)"` | `_verify_tables` |
| `OMATH_XML_MARKER` | `str` | `"<m:oMath"` | `_verify_formulas` |
| `WORD_EQUATION_XML_MARKER` | `str` | `"<w:equation"` | `_verify_formulas` |
| `ABSTRACT_SECTION_KEYWORDS` | `tuple[str, ...]` | `("resumen", "abstract", "síntesis", "sumario")` | `_verify_abstract_keywords` |
| `ABSTRACT_PARAGRAPH_LOOKAHEAD` | `int` | `5` | `_verify_abstract_keywords` |
| `MIN_WORDS_FOR_ABSTRACT_CHECK` | `int` | `1000` | `_verify_abstract_keywords` |
| `ABSTRACT_MIN_WORDS` | `int` | `100` | `_verify_abstract_keywords` |
| `ABSTRACT_MAX_WORDS` | `int` | `300` | `_verify_abstract_keywords` |
| `KEYWORD_SECTION_MARKERS` | `tuple[str, ...]` | `("palabras clave", "keywords", "key words", "descriptores")` | `_verify_abstract_keywords` |
| `MIN_KEYWORDS` | `int` | `3` | `_verify_abstract_keywords` |
| `MAX_KEYWORDS` | `int` | `5` | `_verify_abstract_keywords` |

Total: 20 named constants. The implementation MUST define all 20.

#### R4-CheckMethods

Each of the five private check methods MUST return `list[EumicViolationDTO]` (not append to an
instance variable). Signatures:

```
_verify_format(self, document) -> list[EumicViolationDTO]
_verify_figures(self, document) -> list[EumicViolationDTO]
_verify_tables(self, document) -> list[EumicViolationDTO]
_verify_formulas(self, document) -> list[EumicViolationDTO]
_verify_abstract_keywords(self, document, word_count: int) -> list[EumicViolationDTO]
```

#### R4-ExceptionSpecificity

- R4.10 — The bare `except: pass` in `_verify_figures` (around `doc.part.rels` access) MUST be
  replaced with `except (AttributeError, KeyError): pass`. Silent swallow is preserved; no
  logging is added.
- R4.11 — The bare `except: continue` occurrences in `_verify_formulas` (around
  `run._element.xml` access) MUST be replaced with `except AttributeError: continue`. Silent
  swallow is preserved; no logging is added.

#### R4-PreservedBehavior

- R4.12 — The `run._element.xml` access with `isinstance(xml_str, bytes)` decode guard MUST be
  preserved verbatim for behavioral parity with the legacy code. This internal python-docx
  attribute access is a known dependency and MUST be documented in the class docstring.
- R4.13 — The `doc.part.rels` access pattern in `_verify_figures` MUST be preserved with the
  same silent exception handling (now `except (AttributeError, KeyError): pass`).

#### R4-PythonicQuality

- R4.14 — All imports at module top (no local imports — SKILL §3).
- R4.15 — Guard clauses and early returns used where they improve readability (e.g., return empty
  list early when `image_count == 0` or `len(tables) == 0`).
- R4.16 — No `print()` statements. No inline `#` comments in production code (SKILL §7).
- R4.17 — Public class and methods have docstrings (PEP 257, single-line preferred).
- R4.18 — No `Optional[T]` — use `T | None` syntax throughout (SKILL §0).

### REQ-5: VerifyEumicUseCaseWiring

**File**: `src/infrastructure/wirings/verify_eumic_use_case_wiring.py`
**Class**: `VerifyEumicUseCaseWiring`

- R5.1 — `VerifyEumicUseCaseWiring` MUST expose exactly one public method:
  `get_verify_eumic_use_case(self) -> VerifyEumicUseCase` (per SKILL §8 naming convention
  `get_<use_case_snake_case>()`).
- R5.2 — `get_verify_eumic_use_case` MUST return a fully assembled `VerifyEumicUseCase`
  with `DocxEumicAdapter` injected as `format_inspection_port`.
- R5.3 — The wiring MUST define exactly one private method:
  `_get_document_format_inspection_port(self) -> DocumentFormatInspectionPort`.
  This private method returns the port type (interface), not the concrete adapter.
- R5.4 — The wiring class MUST NOT contain business logic — only object creation and wiring.
- R5.5 — No `VerifyEumicUseCaseWiringForTest` is required in Slice 11. `DocxEumicAdapter.__init__`
  has no infrastructure side-effects (it reads files only on `inspect()` call), so the
  production wiring can be tested directly by checking instance types without calling `inspect()`.

### REQ-6: Coexistence Guarantee

- R6.1 — No existing file in the repository MUST be modified by Slice 11.
- R6.2 — `eumic_verifier.py` at the project root MUST remain unchanged and continue to function.
- R6.3 — `main.py` MUST continue calling `verify_eumic_compliance()` from `eumic_verifier.py`
  without modification.
- R6.4 — All 459 tests that pass before Slice 11 MUST continue to pass after Slice 11.

### REQ-7: No New __init__.py Files

- R7.1 — All directories required by Slice 11 already have `__init__.py` files from prior slices:
  - `src/infrastructure/tests/adapters/__init__.py` — exists
  - `src/infrastructure/tests/adapters/document/__init__.py` — exists
  - `src/domain/tests/document/__init__.py` — exists
  - `src/domain/tests/dtos/__init__.py` — exists
  - `src/domain/document/__init__.py` — exists
  - `src/infrastructure/adapters/document/__init__.py` — exists
  - `src/application/tests/__init__.py` — exists
- R7.2 — Slice 11 MUST NOT create any `__init__.py` file.

---

## 3. File Inventory

### Source files (5 new)

| File | Class | Layer |
|---|---|---|
| `src/domain/dtos/eumic_violation_dto.py` | `EumicViolationDTO` | domain |
| `src/domain/document/document_format_inspection_port.py` | `DocumentFormatInspectionPort` | domain |
| `src/application/verify_eumic_use_case.py` | `VerifyEumicUseCase` | application |
| `src/infrastructure/adapters/document/docx_eumic_adapter.py` | `DocxEumicAdapter` | infrastructure |
| `src/infrastructure/wirings/verify_eumic_use_case_wiring.py` | `VerifyEumicUseCaseWiring` | infrastructure |

### Test files (6 new)

| File | What it tests | Layer |
|---|---|---|
| `src/domain/tests/dtos/test_eumic_violation_dto.py` | `EumicViolationDTO` contract | domain |
| `src/domain/tests/document/test_document_format_inspection_port.py` | `DocumentFormatInspectionPort` abstract contract | domain |
| `src/domain/tests/document/fake_document_format_inspection_port.py` | configurable test double | domain |
| `src/application/tests/test_verify_eumic_use_case.py` | `VerifyEumicUseCase` orchestration | application |
| `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py` | wiring assembly | infrastructure |
| `src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py` | adapter check methods | infrastructure |

### Existing files modified: NONE

---

## 4. Named Constants — Exhaustive Specification

All 20 constants are defined at module level in `src/infrastructure/adapters/document/docx_eumic_adapter.py`, before the class definition.

### From `_verify_format`

```python
REQUIRED_MARGIN_CM: float = 2.5
MARGIN_TOLERANCE_CM: float = 0.3
ALLOWED_FONTS: frozenset[str] = frozenset({"Times New Roman", "Arial", "Calibri"})
REQUIRED_FONT_SIZE_PT: int = 12
FONT_SIZE_TOLERANCE_PT: float = 1.0
MAX_NON_JUSTIFIED_RATIO: float = 0.3
```

**Usage mapping**:
- `REQUIRED_MARGIN_CM` + `MARGIN_TOLERANCE_CM`: margin range check
  (`abs(margin.twips - Cm(REQUIRED_MARGIN_CM).twips) > Cm(MARGIN_TOLERANCE_CM).twips`)
- `ALLOWED_FONTS`: non-standard font detection (`fonts_used - ALLOWED_FONTS`)
- `REQUIRED_FONT_SIZE_PT` + `FONT_SIZE_TOLERANCE_PT`: non-standard size detection
  (`abs(size.pt - REQUIRED_FONT_SIZE_PT) > FONT_SIZE_TOLERANCE_PT`)
- `MAX_NON_JUSTIFIED_RATIO`: justification threshold
  (`non_justified / total_paragraphs > MAX_NON_JUSTIFIED_RATIO`)

### From `_verify_figures`

```python
FIGURE_CAPTION_PREFIXES: tuple[str, ...] = ("figura", "fig.", "figure")
FIGURE_NUMBERING_PATTERN: str = r"figura\s+(\d+)"
```

**Usage mapping**:
- `FIGURE_CAPTION_PREFIXES`: detect figure captions
  (`para.text.strip().lower().startswith(FIGURE_CAPTION_PREFIXES)`)
- `FIGURE_NUMBERING_PATTERN`: sequential numbering check
  (`re.search(FIGURE_NUMBERING_PATTERN, caption.lower())`)

### From `_verify_tables`

```python
TABLE_CAPTION_PREFIXES: tuple[str, ...] = ("tabla", "table", "cuadro")
TABLE_NUMBERING_PATTERN: str = r"tabla\s+(\d+)"
```

**Usage mapping**:
- `TABLE_CAPTION_PREFIXES`: detect table title paragraphs
  (`para.text.strip().lower().startswith(TABLE_CAPTION_PREFIXES)`)
- `TABLE_NUMBERING_PATTERN`: sequential numbering check
  (`re.search(TABLE_NUMBERING_PATTERN, title.lower())`)

### From `_verify_formulas`

```python
OMATH_XML_MARKER: str = "<m:oMath"
WORD_EQUATION_XML_MARKER: str = "<w:equation"
```

**Usage mapping**:
- Both markers: OMath/equation detection in run XML
  (`OMATH_XML_MARKER in xml_str or WORD_EQUATION_XML_MARKER in xml_str`)

### From `_verify_abstract_keywords`

```python
ABSTRACT_SECTION_KEYWORDS: tuple[str, ...] = ("resumen", "abstract", "síntesis", "sumario")
ABSTRACT_PARAGRAPH_LOOKAHEAD: int = 5
MIN_WORDS_FOR_ABSTRACT_CHECK: int = 1000
ABSTRACT_MIN_WORDS: int = 100
ABSTRACT_MAX_WORDS: int = 300
KEYWORD_SECTION_MARKERS: tuple[str, ...] = ("palabras clave", "keywords", "key words", "descriptores")
MIN_KEYWORDS: int = 3
MAX_KEYWORDS: int = 5
```

**Usage mapping**:
- `ABSTRACT_SECTION_KEYWORDS`: detect abstract heading paragraph
- `ABSTRACT_PARAGRAPH_LOOKAHEAD`: window size for collecting abstract text after heading
- `MIN_WORDS_FOR_ABSTRACT_CHECK`: gate for triggering abstract/keyword checks
  (`word_count > MIN_WORDS_FOR_ABSTRACT_CHECK`)
- `ABSTRACT_MIN_WORDS` / `ABSTRACT_MAX_WORDS`: validate abstract length range
- `KEYWORD_SECTION_MARKERS`: detect keyword section paragraph
- `MIN_KEYWORDS` / `MAX_KEYWORDS`: validate keyword count range

---

## 5. Test Scenarios

All test files MUST use `unittest.TestCase`. Framework: stdlib `unittest`. No pytest-specific
assertions. Tests written BEFORE the corresponding implementation file (Strict TDD).

### TDD Order (mandatory)

```
Step 1: test_eumic_violation_dto.py         → eumic_violation_dto.py
Step 2: test_document_format_inspection_port.py
        + fake_document_format_inspection_port.py  → document_format_inspection_port.py
Step 3: test_verify_eumic_use_case.py       → verify_eumic_use_case.py
Step 4: test_docx_eumic_adapter.py          → docx_eumic_adapter.py
Step 5: test_verify_eumic_use_case_wiring.py → verify_eumic_use_case_wiring.py
```

### 5.1 `test_eumic_violation_dto.py`

**File**: `src/domain/tests/dtos/test_eumic_violation_dto.py`

| Scenario | What it asserts |
|---|---|
| `test_creates_with_required_fields` | Constructs `EumicViolationDTO(category="format", message="msg", severity=SeverityLevel.WARNING)` without error |
| `test_details_defaults_to_empty_string` | `dto.details == ""` when not provided |
| `test_is_frozen_raises_on_mutation` | Mutating `dto.message` raises `FrozenInstanceError` |
| `test_is_subclass_of_base_dto` | `issubclass(EumicViolationDTO, BaseDTO)` is `True` |
| `test_severity_field_accepts_severity_level_enum` | `dto.severity` is a `SeverityLevel` instance |
| `test_severity_field_rejects_string` | Assigning a plain string to `severity` does not match the type (structural check: field annotation is `SeverityLevel`) |

Note: `test_severity_field_rejects_string` validates the declared annotation via `dataclasses.fields()`, not runtime enforcement (Python dataclasses do not enforce types at runtime).

### 5.2 `test_document_format_inspection_port.py`

**File**: `src/domain/tests/document/test_document_format_inspection_port.py`

| Scenario | What it asserts |
|---|---|
| `test_cannot_instantiate_directly` | `DocumentFormatInspectionPort()` raises `TypeError` |
| `test_inspect_method_has_docx_path_str_parameter` | `signature(DocumentFormatInspectionPort.inspect).parameters["docx_path"].annotation == str` |
| `test_inspect_method_has_word_count_int_parameter` | `signature(DocumentFormatInspectionPort.inspect).parameters["word_count"].annotation == int` |
| `test_inspect_method_returns_list_of_eumic_violation_dto` | `signature(DocumentFormatInspectionPort.inspect).return_annotation == list[EumicViolationDTO]` |
| `test_fake_port_satisfies_interface` | `FakeDocumentFormatInspectionPort` instantiates and `inspect()` returns a `list` |
| `test_fake_port_returns_configured_violations` | Fake constructed with a violation list returns that list from `inspect()` |
| `test_fake_port_raises_configured_exception` | Fake constructed with `error=ValueError("x")` raises `ValueError` from `inspect()` |

### 5.3 `fake_document_format_inspection_port.py`

**File**: `src/domain/tests/document/fake_document_format_inspection_port.py`
**Class**: `FakeDocumentFormatInspectionPort(DocumentFormatInspectionPort)`

Contract:
- Constructor: `__init__(self, violations: list[EumicViolationDTO] | None = None, error: Exception | None = None)`
- `_violations` defaults to `[]` when `violations` is `None`
- `inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]`:
  - If `self._error is not None`, raises `self._error`
  - Otherwise returns `self._violations`
- No file I/O. Pure in-memory test double.
- Pattern matches `FakeCitationExtractionPort` (`src/domain/tests/document/fake_citation_extraction_port.py`).

### 5.4 `test_verify_eumic_use_case.py`

**File**: `src/application/tests/test_verify_eumic_use_case.py`

| Scenario | What it asserts |
|---|---|
| `test_returns_empty_list_when_port_returns_no_violations` | Fake returns `[]`; `use_case.execute("doc.docx", 500)` returns `[]` |
| `test_returns_violations_from_port_unchanged` | Fake returns a list of 2 violations; `execute()` returns the identical list |
| `test_propagates_src_base_error_subclass_from_port` | Fake raises a `BaseSrcError` subclass; `execute()` re-raises it (not wrapped) |
| `test_wraps_unexpected_exception_in_src_generic_error` | Fake raises a plain `RuntimeError`; `execute()` re-raises as `SrcGenericError` |

Notes:
- Use `FakeDocumentFormatInspectionPort` as the port double.
- The `@generic_error_handler` behavior for `SrcBaseWarning`, `SrcBaseNotFound`,
  `SrcBaseNotAuthorized` re-raises them as-is. The test for `BaseSrcError` subclass propagation
  can use any concrete subclass of `BaseSrcError`.

### 5.5 `test_docx_eumic_adapter.py`

**File**: `src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py`

Uses `unittest.mock.MagicMock` to mock python-docx `Document` objects.
Follows the mock pattern from `tests/test_eumic_verifier.py` (`_make_mock_doc` helper style):
a module-level helper function `_make_mock_document(...)` returns a `MagicMock` with the
minimum attributes set up.

#### 5.5.1 `inspect()` — top-level scenarios

| Scenario | What it asserts |
|---|---|
| `test_raises_document_unreadable_when_docx_path_is_invalid` | Patch `Document` to raise `Exception`; `inspect()` raises `DocumentUnreadable` |
| `test_returns_empty_list_for_fully_compliant_document` | All margins correct, no non-standard fonts, correct size, all paragraphs justified, no images, no tables, no formulas, short word_count; `inspect()` returns `[]` |
| `test_returns_all_violations_from_all_check_methods` | Document that violates multiple checks; `inspect()` returns a non-empty list including violations from at least two different categories |

#### 5.5.2 `_verify_format` scenarios

| Scenario | What it asserts |
|---|---|
| `test_format_no_violations_when_all_margins_are_compliant` | All 4 margins set to `Cm(2.5).twips`; `inspect()` violations do not contain any margin violation |
| `test_format_returns_warning_when_margin_exceeds_tolerance` | One margin set to `Cm(4.0).twips`; violation with `severity == SeverityLevel.WARNING` is in the result |
| `test_format_returns_warning_when_non_standard_font_detected` | Run with `font.name = "Comic Sans MS"`; violation about non-standard font is in the result |
| `test_format_returns_no_font_violation_when_font_is_in_allowed_fonts` | Run with `font.name = "Times New Roman"`; no font violation |
| `test_format_returns_info_when_font_size_outside_tolerance` | Run with `font.size.pt = 20`; violation with `severity == SeverityLevel.INFO` is in the result |
| `test_format_returns_warning_when_majority_of_paragraphs_not_justified` | >30% of non-empty paragraphs have `alignment != WD_ALIGN_PARAGRAPH.JUSTIFY`; warning about justification is in the result |

#### 5.5.3 `_verify_figures` scenarios

| Scenario | What it asserts |
|---|---|
| `test_figures_returns_no_violations_when_no_images_in_document` | `doc.part.rels = {}`; `inspect()` contains no figures-related violation |
| `test_figures_returns_warning_when_captions_fewer_than_image_count` | 2 image rels, 1 caption paragraph; violation about missing formal figure titles |
| `test_figures_returns_warning_when_figure_numbering_is_inconsistent` | 2 captions "Figura 1" and "Figura 3" (skips 2); numbering inconsistency violation |
| `test_figures_returns_no_violations_when_captions_match_images_and_numbered_correctly` | 2 image rels, 2 captions "Figura 1" and "Figura 2"; no figures violation |
| `test_figures_handles_missing_rels_attribute_without_error` | `doc.part.rels` raises `AttributeError`; `inspect()` completes normally, no crash |

#### 5.5.4 `_verify_tables` scenarios

| Scenario | What it asserts |
|---|---|
| `test_tables_returns_no_violations_when_no_tables` | `doc.tables = []`; no table violation |
| `test_tables_returns_warning_when_titles_fewer_than_tables` | 2 tables, 1 title paragraph starting with "Tabla"; violation about missing table title |
| `test_tables_returns_warning_when_table_numbering_is_inconsistent` | 2 titles "Tabla 1" and "Tabla 3"; numbering inconsistency violation |
| `test_tables_returns_no_violations_when_all_tables_have_correct_titles` | 2 tables, titles "Tabla 1" and "Tabla 2"; no table violation |

#### 5.5.5 `_verify_formulas` scenarios

| Scenario | What it asserts |
|---|---|
| `test_formulas_returns_no_violations_when_no_omath_xml_in_runs` | Runs have XML without `<m:oMath` or `<w:equation`; no formula violation |
| `test_formulas_returns_info_when_formula_paragraph_not_centered` | Run XML contains `<m:oMath`; paragraph alignment is not CENTER; `SeverityLevel.INFO` violation |
| `test_formulas_returns_no_violations_when_formula_paragraph_is_centered` | Run XML contains `<m:oMath`; paragraph alignment is `WD_ALIGN_PARAGRAPH.CENTER`; no violation |
| `test_formulas_handles_xml_attribute_error_silently` | `run._element.xml` raises `AttributeError`; `inspect()` completes without error |

#### 5.5.6 `_verify_abstract_keywords` scenarios

| Scenario | What it asserts |
|---|---|
| `test_abstract_keywords_no_violations_when_word_count_below_threshold` | `word_count = 500` (< `MIN_WORDS_FOR_ABSTRACT_CHECK`); no abstract/keyword violation regardless of content |
| `test_abstract_keywords_returns_critical_when_abstract_missing_in_long_document` | No paragraph matching abstract keywords; `word_count = 2000`; `SeverityLevel.CRITICAL` violation about missing abstract |
| `test_abstract_keywords_returns_critical_when_keywords_missing_in_long_document` | Abstract paragraph present; no paragraph matching keyword markers; `word_count = 2000`; `SeverityLevel.CRITICAL` violation about missing keywords |
| `test_abstract_keywords_returns_warning_when_abstract_word_count_out_of_range` | Abstract found but only 30 words in look-ahead window; `SeverityLevel.WARNING` violation about abstract length |
| `test_abstract_keywords_returns_warning_when_keyword_count_out_of_range` | Keyword section found but 10 keywords (> `MAX_KEYWORDS`); `SeverityLevel.WARNING` violation about keyword count |
| `test_abstract_keywords_no_violations_when_abstract_and_keywords_valid` | Abstract present (120 words), 4 keywords; `word_count = 2000`; no abstract/keyword violation |

### 5.6 `test_verify_eumic_use_case_wiring.py`

**File**: `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py`

| Scenario | What it asserts |
|---|---|
| `test_get_verify_eumic_use_case_returns_verify_eumic_use_case_instance` | `VerifyEumicUseCaseWiring().get_verify_eumic_use_case()` returns a `VerifyEumicUseCase` instance |
| `test_get_verify_eumic_use_case_injects_docx_eumic_adapter_as_format_inspection_port` | The returned use case's `._format_inspection_port` is a `DocxEumicAdapter` instance |

Notes:
- No `WiringForTest` is used. The production wiring is tested directly.
- `DocxEumicAdapter.__init__` has no side-effects so instantiation in test is safe.

---

## 6. Out of Scope for Slice 11

The following items MUST NOT be implemented in Slice 11:

| Item | Reason / Deferred To |
|---|---|
| `format_violations_report()` migration | Presentation concern (emoji formatting, grouping). Deferred to Slice 14 (CLI controller) |
| `main.py` replacement of `verify_eumic_compliance()` call | Coexistence strategy. Slice 14 |
| Deletion of `eumic_verifier.py` | Kept as-is until Slice 14 deletes it |
| New domain exception class for EUMIC | Not needed. `DocumentUnreadable` (existing) covers docx open failure |
| New `EumicSeverity` enum | Not needed. `SeverityLevel` covers INFO / WARNING / CRITICAL |
| Additional EUMIC check methods beyond the original five | Post-Slice 11 feature work |
| `VerifyEumicUseCaseWiringForTest` | Not required. No infrastructure dependency needing test-double swap |
| Gradio app integration | `gradio_app.py` does not use `eumic_verifier.py` |
| New `__init__.py` files | All required directories already initialized from prior slices |
| Database or external service calls | Pure docx file inspection; no persistence |

---

## 7. Acceptance Criteria

Slice 11 is complete when ALL of the following are true:

1. **All new tests pass**: `python -m pytest src/ -q` shows 0 failures across all 11 new test
   files.
2. **Total test count**: Total passing tests is 459 + N where N is the number of new test
   methods added (expected approximately 40–50 new test methods).
3. **No regressions**: The 459 tests passing before Slice 11 all continue to pass.
4. **TDD order respected**: Each test file was created and run (red) before its corresponding
   implementation file was created. The Strict TDD constraint is not verifiable post-hoc, but
   the apply phase MUST follow TDD order from Section 5.
5. **Zero modifications to existing files**: `git diff --name-only` for modified files shows only
   new files (all new additions, no edits to pre-existing files).
6. **Named constants**: All 20 constants from Section 4 are present at module level in
   `docx_eumic_adapter.py`. No magic literal encoding an EUMIC editorial rule appears inside any
   method body.
7. **No bare except**: `grep -n "except:" src/infrastructure/adapters/document/docx_eumic_adapter.py`
   returns zero results.
8. **Import invariant**: No file in `src/domain/` imports from `src/application/` or
   `src/infrastructure/`. No file in `src/application/` imports from `src/infrastructure/`.
9. **Coexistence**: `python main.py` (or whatever invokes the legacy path) continues to work
   without error.

---

## 8. Risks Noted

| Risk | Decision in Spec |
|---|---|
| `margin_value.cm` attribute mocking for `_verify_format` tests — `docx.shared.Length` objects need careful mock setup | Use `MagicMock`; set `.twips` directly (int value from `Cm(x).twips`); use `PropertyMock` on `.cm` for the details string. See `tests/test_eumic_verifier.py` for reference pattern. |
| `_verify_formulas` uses `run._element.xml` (internal python-docx attribute) | Mock `run._element.xml` as a regular attribute returning a string; test the bytes path by returning a `bytes` object. |
| `doc.part.rels` can be `AttributeError` — need to verify exception handling in test | Covered by `test_figures_handles_missing_rels_attribute_without_error` scenario. |
| Wiring method name inconsistency — some existing wirings use `create_use_case()`, SKILL §8 mandates `get_<use_case_name>()` | Spec mandates `get_verify_eumic_use_case()` per SKILL §8. The inconsistency in earlier slices is noted but not retrofitted here. |
| Abstract word count counting method — `_verify_abstract_keywords` counts words in a window of paragraphs after the abstract heading | This behavior is preserved from legacy. Test must mock the `doc.paragraphs` list with enough entries in the look-ahead window. |
