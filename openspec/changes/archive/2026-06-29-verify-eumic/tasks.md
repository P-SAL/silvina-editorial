# SDD Tasks — verify-eumic (Slice 11)

**Change**: verify-eumic | **Branch**: refactor/hexagonal-migration
**Strategy**: Strict TDD — every test file precedes its implementation file.
**Test runner**: `.venv\Scripts\python -m pytest src/ -q` (from repo root)
**Baseline**: 459 passing tests, zero regressions allowed.

---

## Dependency Order

```
T1.1 → T1.2 → T2.1 → T2.2 → T3.1 → T3.2 → T4.1 → T4.2 → T5.1 → T5.2 → T6.1
```

All tasks are sequential (each depends on the previous). T2.1 produces both test
and fake in one step because the fake is a test artifact, not a production file.

---

## Constants Reconciliation (exhaustive audit of eumic_verifier.py)

The spec listed 20 constants, the design listed 13. After full audit of `eumic_verifier.py`,
the exhaustive set is **21 named constants** at module level in `docx_eumic_adapter.py`.
Every magic literal in a method body must become a named constant.

```python
# Margins
REQUIRED_MARGIN_CM: float = 2.5
MARGIN_TOLERANCE_CM: float = 0.3

# Fonts
ALLOWED_FONTS: frozenset[str] = frozenset({"Times New Roman", "Arial", "Calibri"})
REQUIRED_FONT_SIZE_PT: int = 12
FONT_SIZE_TOLERANCE_PT: float = 1.0
MAX_DISPLAYED_NON_STANDARD_SIZES: int = 3

# Alignment
MAX_UNJUSTIFIED_PARAGRAPH_RATIO: float = 0.3

# Figures
FIGURE_CAPTION_PREFIXES: tuple[str, ...] = ("figura", "fig.", "figure")
FIGURE_NUMBERING_PATTERN: str = r"figura\s+(\d+)"

# Tables
TABLE_CAPTION_PREFIXES: tuple[str, ...] = ("tabla", "table", "cuadro")
TABLE_NUMBERING_PATTERN: str = r"tabla\s+(\d+)"

# Formulas
OMATH_XML_MARKER: str = "<m:oMath"
WORD_EQUATION_XML_MARKER: str = "<w:equation"

# Abstract / Keywords
ABSTRACT_SECTION_KEYWORDS: tuple[str, ...] = ("resumen", "abstract", "síntesis", "sumario")
ABSTRACT_PARAGRAPH_LOOKAHEAD: int = 5
MIN_WORDS_FOR_ABSTRACT_CHECK: int = 1000
ABSTRACT_MIN_WORD_COUNT: int = 100
ABSTRACT_MAX_WORD_COUNT: int = 300
KEYWORD_SECTION_MARKERS: tuple[str, ...] = ("palabras clave", "keywords", "key words", "descriptores")
MIN_KEYWORD_COUNT: int = 3
MAX_KEYWORD_COUNT: int = 5
```

Name authority: design wins over spec for naming conflicts
(`MAX_UNJUSTIFIED_PARAGRAPH_RATIO`, `ABSTRACT_MIN_WORD_COUNT`, `ABSTRACT_MAX_WORD_COUNT`,
`MIN_KEYWORD_COUNT`, `MAX_KEYWORD_COUNT`).

---

## Group 1 — DTO (domain layer)

### T1.1 — Write test: EumicViolationDTO
**File**: `src/domain/tests/dtos/test_eumic_violation_dto.py`
**Depends on**: nothing (first task)
**Spec ref**: "EumicViolationDTO(BaseDTO) frozen dataclass"

Write a `TestEumicViolationDTO(TestCase)` covering:
- `test_fields_present` — instantiate with all required fields, assert each attribute value
- `test_details_defaults_to_empty_string` — omit `details`, assert `""`
- `test_frozen_raises_on_mutation` — assert `FrozenInstanceError` on field assignment
- `test_severity_accepts_severity_level_enum` — assign `SeverityLevel.WARNING`, assert type

Required imports: `FrozenInstanceError` from `dataclasses`, `SeverityLevel` from
`src.domain.enums.severity_level`, `EumicViolationDTO` from `src.domain.dtos.eumic_violation_dto`.

**Verify after T1.1**: test collection only (file doesn't exist yet, collection will fail — this
is expected and confirms TDD is being followed; run after T1.2).

---

### T1.2 — Implement: EumicViolationDTO
**File**: `src/domain/dtos/eumic_violation_dto.py`
**Depends on**: T1.1
**Spec ref**: "EumicViolationDTO(BaseDTO) frozen dataclass: category: str, message: str,
severity: SeverityLevel, details: str = ''"

```python
from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO
from src.domain.enums.severity_level import SeverityLevel


@dataclass(frozen=True)
class EumicViolationDTO(BaseDTO):
    """Data transfer object representing a single EUMIC format violation."""

    category: str
    message: str
    severity: SeverityLevel
    details: str = ""
```

**Verify after T1.2**: `.venv\Scripts\python -m pytest src/domain/tests/dtos/test_eumic_violation_dto.py -q`
— must pass; no regressions in `src/domain/tests/dtos/`.

---

## Group 2 — Port + Fake (domain layer)

### T2.1 — Write test + fake: DocumentFormatInspectionPort
**Files** (two files created in one step — fake is a test artifact):
- `src/domain/tests/document/test_document_format_inspection_port.py`
- `src/domain/tests/document/fake_document_format_inspection_port.py`

**Depends on**: T1.2 (EumicViolationDTO must exist for fake return type)
**Spec ref**: "DocumentFormatInspectionPort(ABC) with inspect(docx_path: str, word_count: int) -> list[EumicViolationDTO]"

**test_document_format_inspection_port.py** — `TestDocumentFormatInspectionPort(TestCase)`:
- `test_is_abstract_base_class` — `assertRaises(TypeError)` when instantiating directly
- `test_declares_exactly_one_abstract_method_inspect` — `__abstractmethods__ == frozenset({"inspect"})`
- `test_module_has_no_docx_or_infrastructure_imports` — use `getsource` + `assertNotIn`
  (pattern identical to `test_document_text_port.py`)

**fake_document_format_inspection_port.py** — `FakeDocumentFormatInspectionPort(DocumentFormatInspectionPort)`:
```python
class FakeDocumentFormatInspectionPort(DocumentFormatInspectionPort):
    def __init__(
        self,
        violations: list[EumicViolationDTO] | None = None,
        error: Exception | None = None,
    ) -> None:
        self._violations = violations if violations is not None else []
        self._error = error

    def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        if self._error is not None:
            raise self._error
        return self._violations
```

**Verify after T2.1**: run after T2.2.

---

### T2.2 — Implement: DocumentFormatInspectionPort
**File**: `src/domain/document/document_format_inspection_port.py`
**Depends on**: T2.1
**Spec ref**: "DocumentFormatInspectionPort(ABC) with abstract inspect()"

```python
from abc import ABC, abstractmethod

from src.domain.dtos.eumic_violation_dto import EumicViolationDTO


class DocumentFormatInspectionPort(ABC):
    """Port defining EUMIC format inspection for a document."""

    @abstractmethod
    def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        """Inspect the document and return all EUMIC format violations found."""
```

**Verify after T2.2**: `.venv\Scripts\python -m pytest src/domain/tests/document/test_document_format_inspection_port.py -q`
— must pass.

---

## Group 3 — Use Case (application layer)

### T3.1 — Write test: VerifyEumicUseCase
**File**: `src/application/tests/test_verify_eumic_use_case.py`
**Depends on**: T2.2 (port and fake must exist)
**Spec ref**: "VerifyEumicUseCase with @generic_error_handler on execute(); pure delegation to port"

`TestVerifyEumicUseCase(TestCase)` covering:
- `test_execute_returns_empty_list_when_no_violations` — fake with `violations=[]`, call
  `execute(docx_path="any.docx", word_count=500)`, assert `[]`
- `test_execute_returns_violations_from_port` — fake with 2 violations, assert same list returned
- `test_execute_passes_docx_path_and_word_count_to_port` — extend fake to capture last call args,
  assert the passed values match; OR use `unittest.mock.patch` on the fake's `inspect` method
- `test_execute_propagates_exception_from_port` — fake with `error=RuntimeError("fail")`,
  assert `RuntimeError` is raised by `execute()`

Required imports: `VerifyEumicUseCase` from `src.application.verify_eumic_use_case`,
`FakeDocumentFormatInspectionPort` from `src.domain.tests.document.fake_document_format_inspection_port`,
`EumicViolationDTO` from `src.domain.dtos.eumic_violation_dto`,
`SeverityLevel` from `src.domain.enums.severity_level`.

Helper to build a violation:
```python
def _make_violation() -> EumicViolationDTO:
    return EumicViolationDTO(
        category="Test", message="test", severity=SeverityLevel.WARNING
    )
```

---

### T3.2 — Implement: VerifyEumicUseCase
**File**: `src/application/verify_eumic_use_case.py`
**Depends on**: T3.1
**Spec ref**: "VerifyEumicUseCase with @generic_error_handler on execute(); pure delegation"

```python
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO
from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler


class VerifyEumicUseCase:
    """Use case that delegates EUMIC format inspection to the injected port."""

    def __init__(self, document_format_inspection_port: DocumentFormatInspectionPort) -> None:
        self._document_format_inspection_port = document_format_inspection_port

    @generic_error_handler
    def execute(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        """Execute EUMIC format inspection and return all violations found."""
        return self._document_format_inspection_port.inspect(
            docx_path=docx_path, word_count=word_count
        )
```

**Verify after T3.2**: `.venv\Scripts\python -m pytest src/application/tests/test_verify_eumic_use_case.py -q`
— must pass.

---

## Group 4 — Adapter (infrastructure layer, most complex)

### T4.1 — Write test: DocxEumicAdapter
**File**: `src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py`
**Depends on**: T3.2 (EumicViolationDTO and port must exist)
**Spec ref**: "DocxEumicAdapter(DocumentFormatInspectionPort) — 21 constants, 5 check groups,
16 private methods, functional return per group"

**NOTE**: `src/infrastructure/tests/adapters/document/__init__.py` already exists — do NOT create it.

**Test structure**: `TestDocxEumicAdapter(TestCase)`.

Patch target for all tests: `src.infrastructure.adapters.document.docx_eumic_adapter.Document`.

Module-level mock helper (add before the test class):
```python
def _mock_length(cm_value: float) -> int:
    """Convert cm to EMU twips for margin mock. Formula: cm * 914400 / 2.54 / 100."""
    return int(round(cm_value * 914400 / 2.54 / 100))
```

**Test scenarios required** (group by check area):

**Format checks** — patch `Document`, configure `mock_doc.sections`, `mock_doc.paragraphs`:
- `test_inspect_returns_no_violations_when_margins_are_within_tolerance` — margins set to
  `_mock_length(2.5)` on all sides; `top_margin`, `bottom_margin`, `left_margin`, `right_margin`
  each returning a `Mock(twips=_mock_length(2.5))`. Assert violations list is empty.
- `test_inspect_returns_margin_violation_when_margin_exceeds_tolerance` — one margin set to
  `_mock_length(1.0)` (far below 2.5). Assert one violation with `category=="Formato General"`.
- `test_inspect_returns_font_violation_when_non_standard_font_used` — paragraph runs with
  `run.font.name = "Comic Sans"`. Assert violation with `category=="Formato General"`.
- `test_inspect_returns_no_font_violation_when_all_fonts_are_standard` — runs with
  `font.name` in `{"Times New Roman", "Arial", "Calibri"}`. Assert no font violation.
- `test_inspect_returns_font_size_violation_when_non_standard_size_used` — run with
  `font.size` mocked as `Mock(pt=18)` (> 12 + 1.0). Assert violation.
- `test_inspect_returns_alignment_violation_when_too_many_non_justified_paragraphs` —
  configure > 30% of non-empty paragraphs with alignment != JUSTIFY. Assert violation.

**Figure checks**:
- `test_inspect_returns_no_figure_violations_when_no_images` — `doc.part.rels.values()` returns
  empty; assert no figure violations.
- `test_inspect_returns_figure_caption_violation_when_captions_fewer_than_images` — 2 images
  in rels, only 1 paragraph starting with "figura". Assert violation.
- `test_inspect_returns_numbering_violation_when_figure_numbers_not_sequential` — 2 captions:
  "Figura 1. ..." and "Figura 3. ...". Assert numbering violation.
- `test_inspect_suppresses_exception_from_rels_iteration` — `doc.part.rels.values()` raises
  `KeyError`. Assert no exception propagates; `image_count` treated as 0.

**Table checks**:
- `test_inspect_returns_no_table_violations_when_no_tables` — `doc.tables = []`. Assert no
  table violations.
- `test_inspect_returns_table_title_violation_when_titles_fewer_than_tables` — 2 tables,
  1 title paragraph starting with "tabla". Assert violation.
- `test_inspect_returns_table_numbering_violation_when_non_sequential` — 2 titles: "Tabla 1."
  and "Tabla 3.". Assert violation.

**Formula checks**:
- `test_inspect_returns_no_formula_violations_when_no_formulas` — no runs with omath/equation
  XML. Assert no formula violations.
- `test_inspect_returns_formula_alignment_violation_when_formula_paragraphs_not_centered` —
  run with `_element.xml` containing `"<m:oMath"`, paragraph alignment != CENTER. Assert violation.
- `test_inspect_suppresses_attribute_error_in_formula_run_parsing` — run's `_element.xml`
  raises `AttributeError`. Assert no exception propagates.

**Abstract / keywords checks**:
- `test_inspect_skips_abstract_check_when_word_count_below_threshold` — `word_count=500`
  (below 1000). Assert no abstract/keyword violations regardless of content.
- `test_inspect_returns_abstract_missing_violation_when_word_count_above_threshold` —
  `word_count=1500`, no paragraph with abstract keywords. Assert CRITICAL violation.
- `test_inspect_returns_abstract_length_violation_when_abstract_too_short` — `word_count=1500`,
  paragraph with "Resumen" found, but next 5 paragraphs contain < 100 words total. Assert WARNING.
- `test_inspect_returns_abstract_length_violation_when_abstract_too_long` — abstract content
  yields > 300 words. Assert WARNING.
- `test_inspect_returns_keyword_missing_violation_when_no_keyword_section` — `word_count=1500`,
  abstract present but no keyword marker paragraph. Assert CRITICAL keyword violation.
- `test_inspect_returns_keyword_count_violation_when_too_few_keywords` — keyword line has
  only 2 comma-separated entries. Assert WARNING.
- `test_inspect_returns_keyword_count_violation_when_too_many_keywords` — 6 entries. Assert WARNING.
- `test_inspect_returns_empty_list_when_fully_compliant_document` — all checks pass. Assert `[]`.

**Verify after T4.1**: run after T4.2.

---

### T4.2 — Implement: DocxEumicAdapter
**File**: `src/infrastructure/adapters/document/docx_eumic_adapter.py`
**Depends on**: T4.1
**Spec ref**: "DocxEumicAdapter(DocumentFormatInspectionPort) — senior-level refactor,
21 constants, 5 check group methods, 16 private helper methods, functional style"

**Structure overview**:

```
DocxEumicAdapter(DocumentFormatInspectionPort)
├── inspect(docx_path, word_count) -> list[EumicViolationDTO]   [public, @generic_error_handler]
│
├── _verify_format(document) -> list[EumicViolationDTO]
│   ├── _check_margins(document) -> list[EumicViolationDTO]
│   ├── _check_fonts(document) -> list[EumicViolationDTO]
│   └── _check_text_alignment(document) -> list[EumicViolationDTO]
│
├── _verify_figures(document) -> list[EumicViolationDTO]
│   ├── _count_image_relationships(document) -> int
│   ├── _collect_paragraphs_starting_with(document, prefixes) -> list[str]  [shared]
│   ├── _check_figure_caption_count(image_count, captions) -> list[EumicViolationDTO]
│   └── _check_sequential_numbering_violations(titles, pattern) -> list[EumicViolationDTO]  [shared]
│
├── _verify_tables(document) -> list[EumicViolationDTO]
│   ├── _collect_paragraphs_starting_with(document, prefixes) -> list[str]  [shared, reused]
│   ├── _check_table_title_count(table_count, titles) -> list[EumicViolationDTO]
│   └── _check_sequential_numbering_violations(titles, pattern) -> list[EumicViolationDTO]  [shared, reused]
│
├── _verify_formulas(document) -> list[EumicViolationDTO]
│   ├── _collect_formula_paragraphs(document) -> list[paragraph]
│   ├── _paragraph_contains_formula(paragraph) -> bool
│   ├── _run_contains_omath(run) -> bool
│   └── _check_formula_alignment(document, formula_count) -> list[EumicViolationDTO]
│
└── _verify_abstract_keywords(document, word_count) -> list[EumicViolationDTO]
    ├── _find_abstract_word_count(document) -> tuple[bool, int]
    ├── _check_abstract(has_abstract, word_count_in_abstract, total_word_count) -> list[EumicViolationDTO]
    ├── _find_keyword_count(document) -> tuple[bool, int]
    └── _check_keywords(has_keywords, keyword_count, total_word_count) -> list[EumicViolationDTO]
```

**Key implementation constraints**:

1. `inspect()` opens `Document(docx_path)` (imported from `docx`), then calls all 5 `_verify_*`
   methods and returns the concatenated list. Must have `@generic_error_handler`.

2. `_verify_format`: margin comparison uses `abs(margin_value - _mock_length_equivalent) > tolerance`.
   In production, convert using the same EMU formula: `int(round(cm * 914400 / 2.54 / 100))`.
   The adapter must NOT import `Cm` from `docx.shared` for comparison — compute directly:
   ```python
   _required_margin_emu: int = int(round(REQUIRED_MARGIN_CM * 914400 / 2.54 / 100))
   _margin_tolerance_emu: int = int(round(MARGIN_TOLERANCE_CM * 914400 / 2.54 / 100))
   ```
   (These can be module-level computed constants, not additional named constants.)

3. `_count_image_relationships`: wraps `doc.part.rels.values()` iteration in
   `except (KeyError, AttributeError): pass` — returns 0 on failure (silent).

4. `_run_contains_omath`: handles `AttributeError` via `except AttributeError: continue` when
   accessing `run._element.xml`. Checks for `OMATH_XML_MARKER` and `WORD_EQUATION_XML_MARKER`.

5. `_verify_formulas`: uses `except AttributeError: continue` per design.

6. `_check_sequential_numbering_violations`: shared utility used by both figures and tables.
   Takes `titles: list[str]` and `pattern: str` (the regex). Returns violations only when
   `len(titles) > 1`.

7. `_collect_paragraphs_starting_with`: shared utility — `para.text.strip().lower().startswith(prefixes)`.

8. `_verify_abstract_keywords` only runs when `word_count >= MIN_WORDS_FOR_ABSTRACT_CHECK`.
   Returns `[]` immediately otherwise.

9. No `self.violations` accumulation anywhere. Each method builds and returns its own list.

10. All 21 constants at module level (before the class definition). No magic literals inside method bodies.

**Verify after T4.2**: `.venv\Scripts\python -m pytest src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py -q`
— must pass. Also run full suite: `.venv\Scripts\python -m pytest src/ -q` — no regressions.

---

## Group 5 — Wiring (infrastructure layer)

### T5.1 — Write test: VerifyEumicUseCaseWiring
**File**: `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py`
**Depends on**: T4.2 (adapter must exist)
**Spec ref**: "VerifyEumicUseCaseWiring with create_use_case()"

**Confirmed decision**: wiring method is `create_use_case()` — NOT `get_verify_eumic_use_case()`.
Design is authoritative. Pattern matches `CheckGrammarUseCaseWiring`.

No `skipIf` needed — `python-docx` is always installed (core project dependency, no runtime side-effects
in `DocxEumicAdapter.__init__`).

`TestVerifyEumicUseCaseWiring(TestCase)`:
- `test_create_use_case_returns_verify_eumic_use_case_instance` — assert `isinstance(result, VerifyEumicUseCase)`
- `test_create_use_case_wires_docx_eumic_adapter_as_inspection_port` — assert
  `isinstance(result._document_format_inspection_port, DocxEumicAdapter)`

Required imports: `VerifyEumicUseCase` from `src.application.verify_eumic_use_case`,
`DocxEumicAdapter` from `src.infrastructure.adapters.document.docx_eumic_adapter`,
`VerifyEumicUseCaseWiring` from `src.infrastructure.wirings.verify_eumic_use_case_wiring`.

---

### T5.2 — Implement: VerifyEumicUseCaseWiring
**File**: `src/infrastructure/wirings/verify_eumic_use_case_wiring.py`
**Depends on**: T5.1
**Spec ref**: "VerifyEumicUseCaseWiring with create_use_case()"

```python
from src.application.verify_eumic_use_case import VerifyEumicUseCase
from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.infrastructure.adapters.document.docx_eumic_adapter import DocxEumicAdapter


class VerifyEumicUseCaseWiring:
    """Factory for building a ready-to-use VerifyEumicUseCase."""

    def create_use_case(self) -> VerifyEumicUseCase:
        """Return a fully assembled VerifyEumicUseCase."""
        return VerifyEumicUseCase(
            document_format_inspection_port=self._get_document_format_inspection_port()
        )

    def _get_document_format_inspection_port(self) -> DocumentFormatInspectionPort:
        """Return the DocxEumicAdapter as the document format inspection port."""
        return DocxEumicAdapter()
```

**Verify after T5.2**: `.venv\Scripts\python -m pytest src/infrastructure/tests/test_verify_eumic_use_case_wiring.py -q`
— must pass.

---

## Group 6 — Full Test Suite Verification

### T6.1 — Run full test suite
**Depends on**: T5.2
**Command**: `.venv\Scripts\python -m pytest src/ -q`

**Pass criteria**:
- All 459 original tests still pass (zero regressions)
- All N new test methods pass (N ≈ 40–50, confirmed by counting test methods written in T1.1–T5.1)
- Zero bare `except:` or `except Exception:` in `docx_eumic_adapter.py`
- All 21 constants at module level in `docx_eumic_adapter.py`, no magic literals in method bodies
- Zero existing files modified

---

## File Manifest

### Production files (5)
| Task | File |
|------|------|
| T1.2 | `src/domain/dtos/eumic_violation_dto.py` |
| T2.2 | `src/domain/document/document_format_inspection_port.py` |
| T3.2 | `src/application/verify_eumic_use_case.py` |
| T4.2 | `src/infrastructure/adapters/document/docx_eumic_adapter.py` |
| T5.2 | `src/infrastructure/wirings/verify_eumic_use_case_wiring.py` |

### Test files (6)
| Task | File |
|------|------|
| T1.1 | `src/domain/tests/dtos/test_eumic_violation_dto.py` |
| T2.1 | `src/domain/tests/document/test_document_format_inspection_port.py` |
| T2.1 | `src/domain/tests/document/fake_document_format_inspection_port.py` |
| T3.1 | `src/application/tests/test_verify_eumic_use_case.py` |
| T4.1 | `src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py` |
| T5.1 | `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py` |

**Total**: 11 new files, 0 existing files modified.

---

## Risks

1. **Margin EMU formula**: The mock helper `_mock_length` and the production adapter must use
   the same formula (`int(round(cm * 914400 / 2.54 / 100))`). If they drift, margin tests will
   give false negatives. The formula must appear verbatim in both the test module (as helper) and
   the adapter (as computed constants `_required_margin_emu`, `_margin_tolerance_emu`).

2. **_collect_paragraphs_starting_with sharing**: Used by both `_verify_figures` and `_verify_tables`.
   Tests must cover both call sites to confirm the shared method works for both prefix tuples.

3. **Constants count**: 21 constants is the authoritative count. Any future refactor that
   introduces a new literal without a constant is a regression against this spec. Implementer
   must audit method bodies after writing before calling T4.2 done.

4. **No WiringForTest**: Pattern not established in this project. Wiring test uses production
   wiring directly — if `DocxEumicAdapter.__init__` ever gains side effects, this task requires a revisit.
