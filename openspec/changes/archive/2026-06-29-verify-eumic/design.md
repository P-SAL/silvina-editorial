# SDD Design — verify-eumic (Slice 11)

**Change name**: verify-eumic
**Slice**: 11
**Branch**: refactor/hexagonal-migration
**Author**: SDD design phase
**Date**: 2026-06-29
**Status**: Approved

---

## 1. Architecture Decisions (ADRs)

### ADR-1: `word_count: int` in port signature (not full DTO)

**Decision**: `DocumentFormatInspectionPort.inspect(docx_path: str, word_count: int)` takes a primitive `int`.

**Rationale**: The only data the adapter reads from the legacy `document_content` object is `document_content.word_count` (in `_verify_abstract_keywords`). Coupling the port to a full `DocumentContentDTO` would violate Interface Segregation — the port declares a dependency on a type it only uses for one integer field. Passing the primitive directly makes the interface minimal and honest about what it actually needs.

**Rejected alternative**: `inspect(docx_path: str, document_content: DocumentContentDTO)` — introduces a DTO dependency across the port boundary for a single integer. Any future adapter (even a stub) would need to construct or mock the full DTO.

---

### ADR-2: Named constants live in the adapter (not domain)

**Decision**: All editorial threshold constants are placed at module level in `docx_eumic_adapter.py`.

**Rationale**: These constants answer "how does the DocxEumicAdapter decide compliance?" — they are implementation details of a specific docx inspection approach, not pure domain invariants. Domain must not import adapter constants (import invariant: domain → domain only). No second EUMIC adapter is planned (YAGNI). If a second adapter appears, extracting to a domain constants module is a 10-minute refactor; premature extraction now creates false abstraction.

**Rejected alternative**: `src/domain/constants/eumic_editorial_standards.py` — domain would encode adapter-specific thresholds; violates the principle that domain is independent of infrastructure implementation details.

---

### ADR-3: `SeverityLevel` reuse, no new enum

**Decision**: Reuse `src/domain/enums/severity_level.py` with INFO / WARNING / CRITICAL.

**Rationale**: The original `EumicSeverity` enum's values (`"🔴 CRÍTICO"`, `"🟡 ADVERTENCIA"`, `"ℹ️ INFO"`) are presentation strings — they encode formatting and emoji for a report output. These belong in the controller (Slice 14), not the domain. The domain enum `SeverityLevel` has the correct abstraction: named severity without presentation markup.

**Rejected alternative**: Migrate `EumicSeverity` into the domain — would persist a presentation concern (`"🔴 CRÍTICO"` as a domain enum value) into the clean architecture core.

---

### ADR-4: `DocumentUnreadable` reuse, no new exception

**Decision**: Raise `DocumentUnreadable` (existing, from `src/domain/exceptions/document_errors.py`) when the docx file cannot be opened.

**Rationale**: All document adapters use this exception for docx open failures (`DocxTextAdapter` uses the same pattern). No EUMIC-specific failure mode exists at infrastructure that would require a distinct exception class. Callers catch `DocumentUnreadable` uniformly for any unreadable document.

**Rejected alternative**: `EumicDocumentUnreadable(DocumentUnreadable)` — no caller needs to distinguish which adapter failed to open the file; the added subclass has no semantic value.

---

### ADR-5: `@generic_error_handler` on `inspect()` only

**Decision**: Decorate only the public `inspect()` method.

**Rationale**: Consistent with all established adapters (`DocxTextAdapter.read_paragraphs`, `DocxCitationAdapter.extract_citations`). Internal private methods' exceptions bubble up to the decorated public boundary. Decorating every private helper would add noise without additional safety — the handler fires once at the outermost level.

**Rejected alternative**: Decorate each private check method — redundant nesting; `SrcGenericError` would wrap an already-wrapped exception.

---

### ADR-6: Functional style — each check method returns violations

**Decision**: All five check methods return `list[EumicViolationDTO]`. `inspect()` aggregates with `violations.extend(...)`. No mutable class state.

**Rationale**: The legacy `EumicVerifier` accumulates violations in `self.violations`, which is reset on each call. This is thread-unsafe and breaks isolation in tests — running two checks from the same instance mixes results. Returning lists makes each check method independently testable, composable, and free of side effects.

**Rejected alternative**: Preserve `self.violations` mutable list on the adapter — produces subtle bugs when the adapter is reused; `inspect()` must always reset state before use.

---

### ADR-7: Wiring public method named `create_use_case()`

**Decision**: `VerifyEumicUseCaseWiring.create_use_case()` — not `get_verify_eumic_use_case()`.

**Rationale**: Every production wiring in the actual codebase uses `create_use_case()` as the public method name (verified: `CheckGrammarUseCaseWiring`, `ExtractCitationsUseCaseWiring`, `ExtractCitationsUseCaseWiring`, etc.). SKILL §8 prescribes `get_<use_case_snake_case>()`, but the project has established `create_use_case()` as the practical convention across all slices. Consistency within the codebase outweighs strict SKILL compliance on internal naming.

**Rejected alternative**: `get_verify_eumic_use_case()` — matches SKILL §8 literally but inconsistent with every other wiring in the project; forces callers to memorize per-class method names.

---

### ADR-8: No `VerifyEumicUseCaseWiringForTest`

**Decision**: Do not create a WiringForTest class.

**Rationale**: The `src/infrastructure/tests/test_doubles/` directory exists but is empty — this pattern has not been implemented in any slice. The application-level test (`test_verify_eumic_use_case.py`) uses `FakeDocumentFormatInspectionPort` directly, not through a wiring. The wiring integration test exercises the real `DocxEumicAdapter`. No integration test scenario requires swapping the adapter through a test wiring.

**Rejected alternative**: Create `VerifyEumicUseCaseWiringForTest(VerifyEumicUseCaseWiring)` per SKILL §8 — introduces a pattern with no concrete test scenario to validate it; the pattern remains untested infrastructure.

---

### ADR-9: Sub-helpers extracted from each check method

**Decision**: Each of the five check methods delegates to named private sub-helpers. See Section 2 for the complete breakdown.

**Rationale**: Confirmed decision #3 (senior-level refactor, not 1:1 migration). The legacy methods mix counting, collecting, and checking in single nested loops. Extracting named helpers makes each unit independently describable, testable, and readable. Each helper has one responsibility and a name that expresses it.

---

## 2. Adapter Internal Design

This section is the primary design artifact. The `DocxEumicAdapter` is the most complex file and requires the most architectural specification.

### Module-level constants

All thresholds and literal values that encode EUMIC editorial policy are extracted here. No literal numbers or strings may appear in method bodies.

```
REQUIRED_MARGIN_CM: float = 2.5
MARGIN_TOLERANCE_CM: float = 0.3
REQUIRED_FONT_SIZE_PT: int = 12
FONT_SIZE_TOLERANCE_PT: float = 1.0
MAX_DISPLAYED_NON_STANDARD_SIZES: int = 3
MAX_UNJUSTIFIED_PARAGRAPH_RATIO: float = 0.3
ALLOWED_FONTS: frozenset[str] = frozenset({"Times New Roman", "Arial", "Calibri"})
MIN_WORDS_FOR_ABSTRACT_CHECK: int = 1000
ABSTRACT_MIN_WORD_COUNT: int = 100
ABSTRACT_MAX_WORD_COUNT: int = 300
ABSTRACT_PARAGRAPH_LOOKAHEAD: int = 5
MIN_KEYWORD_COUNT: int = 3
MAX_KEYWORD_COUNT: int = 5
```

**Rationale for each new constant beyond the four from the proposal**:
- `MARGIN_TOLERANCE_CM`: `0.3` literal in `_verify_format` loop — must be named.
- `FONT_SIZE_TOLERANCE_PT`: `1` literal in font size comparison — must be named.
- `MAX_DISPLAYED_NON_STANDARD_SIZES`: `3` literal in slicing non-standard sizes for display — must be named.
- `MAX_UNJUSTIFIED_PARAGRAPH_RATIO`: `0.3` literal (30% threshold) — must be named.
- `ABSTRACT_MIN_WORD_COUNT` / `ABSTRACT_MAX_WORD_COUNT`: `100` and `300` in abstract length check — must be named.
- `ABSTRACT_PARAGRAPH_LOOKAHEAD`: `5` in abstract text accumulation loop — must be named.
- `MIN_KEYWORD_COUNT` / `MAX_KEYWORD_COUNT`: `3` and `5` in keyword count check — must be named.

### Top-level imports (all module-level, per SKILL §3)

```python
from re import compile, search, split as re_split

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.exceptions import PackageNotFoundError
from docx.shared import Cm, Pt

from src.domain.document.document_format_inspection_port import DocumentFormatInspectionPort
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO
from src.domain.enums.severity_level import SeverityLevel
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.exceptions.document_errors import DocumentUnreadable
```

The legacy code had `from docx.shared import Pt, Cm` and `from docx.enum.text import WD_ALIGN_PARAGRAPH` inside method bodies. Both are moved to module top.

### `inspect(docx_path, word_count) -> list[EumicViolationDTO]`

Public method, decorated with `@generic_error_handler`. Opens the docx file with a specific exception set (`PackageNotFoundError`, `ValueError`, `OSError`) mapped to `DocumentUnreadable`. Delegates to each check method and aggregates results with `violations.extend(...)`.

```
@generic_error_handler
def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
    try:
        document = Document(docx_path)
    except (PackageNotFoundError, ValueError, OSError) as exc:
        raise DocumentUnreadable() from exc
    violations: list[EumicViolationDTO] = []
    violations.extend(self._verify_format(document))
    violations.extend(self._verify_figures(document))
    violations.extend(self._verify_tables(document))
    violations.extend(self._verify_formulas(document))
    violations.extend(self._verify_abstract_keywords(document, word_count))
    return violations
```

---

### `_verify_format(document) -> list[EumicViolationDTO]`

Coordinator for the three format sub-checks. Delegates and aggregates.

```
def _verify_format(self, document):
    violations = []
    violations.extend(self._check_margins(document))
    violations.extend(self._check_fonts(document))
    violations.extend(self._check_text_alignment(document))
    return violations
```

#### `_check_margins(document) -> list[EumicViolationDTO]`

- Guard: early return `[]` if `not document.sections`.
- Takes `document.sections[0]`.
- Computes `required_twips = Cm(REQUIRED_MARGIN_CM).twips` and `tolerance_twips = Cm(MARGIN_TOLERANCE_CM).twips`.
- Iterates the four margins: `("superior", section.top_margin)`, `("inferior", section.bottom_margin)`, `("izquierdo", section.left_margin)`, `("derecho", section.right_margin)`.
- For each margin where `abs(margin_value.twips - required_twips) > tolerance_twips`, appends a `EumicViolationDTO(category="Formato General", severity=SeverityLevel.WARNING, ...)`.

#### `_check_fonts(document) -> list[EumicViolationDTO]`

- Collects `fonts_used: set[str]` and `sizes_used: set` from all paragraphs and runs.
- Computes `non_standard_fonts = fonts_used - ALLOWED_FONTS`. If non-empty → WARNING violation.
- Computes `non_standard_sizes = [size for size in sizes_used if size and abs(size.pt - REQUIRED_FONT_SIZE_PT) > FONT_SIZE_TOLERANCE_PT]`. If non-empty → INFO violation. Uses `MAX_DISPLAYED_NON_STANDARD_SIZES` to limit the displayed sizes in the detail message.

#### `_check_text_alignment(document) -> list[EumicViolationDTO]`

- Iterates paragraphs, counts `total_paragraphs` (non-empty text only) and `non_justified_count` (paragraphs where `para.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY`).
- Guard: early return `[]` if `total_paragraphs == 0`.
- If `non_justified_count / total_paragraphs > MAX_UNJUSTIFIED_PARAGRAPH_RATIO` → WARNING violation.

---

### `_verify_figures(document) -> list[EumicViolationDTO]`

```
def _verify_figures(self, document):
    image_count = self._count_image_relationships(document)
    if image_count == 0:
        return []
    figure_captions = self._collect_paragraphs_starting_with(
        document, ("figura", "fig.", "figure")
    )
    violations = []
    violations.extend(self._check_figure_caption_count(image_count, figure_captions))
    violations.extend(self._check_sequential_numbering_violations(
        figure_captions,
        pattern=r'figura\s+(\d+)',
        category="Figuras",
        message="Numeración de figuras inconsistente",
        details="Las figuras deben numerarse consecutivamente (Figura 1, Figura 2, ...)",
    ))
    return violations
```

#### `_count_image_relationships(document) -> int`

- Iterates `document.part.rels.values()`.
- **Bare except replacement** (confirmed decision #1): `except Exception: pass` — silent, no logging, preserves exact behavior.
- Counts rels where `"image" in rel.target_ref`.
- Returns count as `int`.

#### `_collect_paragraphs_starting_with(document, prefixes) -> list[str]`

- Shared helper reused by both `_verify_figures` and `_verify_tables`.
- Returns `[para.text.strip() for para in document.paragraphs if para.text.strip().lower().startswith(prefixes)]`.

#### `_check_figure_caption_count(image_count, figure_captions) -> list[EumicViolationDTO]`

- If `len(figure_captions) < image_count` → returns list with one WARNING violation.

#### `_check_sequential_numbering_violations(captions, pattern, category, message, details) -> list[EumicViolationDTO]`

- Shared helper for both figure and table sequential numbering.
- Guard: `if len(captions) <= 1: return []` (only one caption → no sequence to validate).
- Iterates captions, extracts number with `re.search(pattern, caption.lower())`.
- If any number deviates from the expected consecutive sequence → returns list with one WARNING violation using `category`, `message`, `details` arguments.
- Returns `[]` if numbering is correct.

---

### `_verify_tables(document) -> list[EumicViolationDTO]`

```
def _verify_tables(self, document):
    if not document.tables:
        return []
    table_titles = self._collect_paragraphs_starting_with(
        document, ("tabla", "table", "cuadro")
    )
    violations = []
    violations.extend(self._check_table_title_count(len(document.tables), table_titles))
    violations.extend(self._check_sequential_numbering_violations(
        table_titles,
        pattern=r'tabla\s+(\d+)',
        category="Tablas",
        message="Numeración de tablas inconsistente",
        details="Las tablas deben numerarse consecutivamente (Tabla 1, Tabla 2, ...)",
    ))
    return violations
```

#### `_check_table_title_count(table_count, table_titles) -> list[EumicViolationDTO]`

- If `len(table_titles) < table_count` → returns list with one WARNING violation.

---

### `_verify_formulas(document) -> list[EumicViolationDTO]`

```
def _verify_formulas(self, document):
    formula_paragraphs = self._collect_formula_paragraphs(document)
    if not formula_paragraphs:
        return []
    return self._check_formula_alignment(formula_paragraphs)
```

#### `_collect_formula_paragraphs(document) -> list[paragraph]`

- Returns all paragraphs where `self._paragraph_contains_formula(para)` is True.

#### `_paragraph_contains_formula(paragraph) -> bool`

- Iterates `paragraph.runs`.
- Returns `True` on the first run where `self._run_contains_omath(run)` is True.
- Returns `False` if no such run found.

#### `_run_contains_omath(run) -> bool`

- **Core of the bytes/str guard** (preserved verbatim from legacy):
  ```
  try:
      xml_string = run._element.xml
      if isinstance(xml_string, bytes):
          xml_string = xml_string.decode('utf-8')
      return '<m:oMath' in xml_string or '<w:equation' in xml_string
  except Exception:
      return False
  ```
- The `except Exception` matches the legacy `except: continue` behavior — if XML access fails, the run is treated as not containing a formula.

#### `_check_formula_alignment(formula_paragraphs) -> list[EumicViolationDTO]`

- Counts `unaligned_count = sum(1 for para in formula_paragraphs if para.alignment != WD_ALIGN_PARAGRAPH.CENTER)`.
- If `unaligned_count > 0` → returns list with one INFO violation.

---

### `_verify_abstract_keywords(document, word_count) -> list[EumicViolationDTO]`

```
def _verify_abstract_keywords(self, document, word_count):
    violations = []
    violations.extend(self._check_abstract(document, word_count))
    violations.extend(self._check_keywords(document, word_count))
    return violations
```

#### `_find_abstract_word_count(document) -> int | None`

- Iterates `document.paragraphs` with `enumerate`.
- Checks each paragraph text for any of `['resumen', 'abstract', 'síntesis', 'sumario']` (case-insensitive).
- On first match: accumulates text from `para_index` to `min(para_index + ABSTRACT_PARAGRAPH_LOOKAHEAD, len(document.paragraphs))`.
- Returns `len(accumulated_text.split())` as the word count.
- Returns `None` if no abstract section is found.

#### `_check_abstract(document, word_count) -> list[EumicViolationDTO]`

- Calls `abstract_word_count = self._find_abstract_word_count(document)`.
- If `abstract_word_count is None` (not found) AND `word_count > MIN_WORDS_FOR_ABSTRACT_CHECK` → CRITICAL violation.
- Elif abstract found AND `(abstract_word_count < ABSTRACT_MIN_WORD_COUNT or abstract_word_count > ABSTRACT_MAX_WORD_COUNT)` → WARNING violation.
- Guard structure: these two branches are mutually exclusive by design.

#### `_find_keyword_count(document) -> int | None`

- Iterates paragraphs looking for `['palabras clave', 'keywords', 'key words', 'descriptores']` (case-insensitive).
- On first match: extracts keyword text after `:` if present, splits by `[,;]`, counts non-empty items.
- Returns `int` count, or `None` if no keyword section found.

#### `_check_keywords(document, word_count) -> list[EumicViolationDTO]`

- Calls `keyword_count = self._find_keyword_count(document)`.
- If `keyword_count is None` AND `word_count > MIN_WORDS_FOR_ABSTRACT_CHECK` → CRITICAL violation.
- Elif keyword section found AND `(keyword_count < MIN_KEYWORD_COUNT or keyword_count > MAX_KEYWORD_COUNT)` → WARNING violation.

---

### Complete private method inventory

| Method | Returns | Purpose |
|---|---|---|
| `_check_margins(document)` | `list[EumicViolationDTO]` | Margin threshold check |
| `_check_fonts(document)` | `list[EumicViolationDTO]` | Font name and size check |
| `_check_text_alignment(document)` | `list[EumicViolationDTO]` | Paragraph justification check |
| `_count_image_relationships(document)` | `int` | Count image rels via `doc.part.rels` |
| `_collect_paragraphs_starting_with(document, prefixes)` | `list[str]` | Shared: collect captions by prefix (figures and tables) |
| `_check_figure_caption_count(image_count, captions)` | `list[EumicViolationDTO]` | Captions vs image count |
| `_check_sequential_numbering_violations(captions, pattern, category, message, details)` | `list[EumicViolationDTO]` | Shared: figure and table sequential numbering |
| `_check_table_title_count(table_count, titles)` | `list[EumicViolationDTO]` | Titles vs table count |
| `_collect_formula_paragraphs(document)` | `list[paragraph]` | Paragraphs containing OMath runs |
| `_paragraph_contains_formula(paragraph)` | `bool` | Does paragraph have any OMath run? |
| `_run_contains_omath(run)` | `bool` | bytes/str guard + XML check |
| `_check_formula_alignment(formula_paragraphs)` | `list[EumicViolationDTO]` | Formula paragraph center alignment |
| `_find_abstract_word_count(document)` | `int \| None` | Locate abstract section, count words |
| `_check_abstract(document, word_count)` | `list[EumicViolationDTO]` | Abstract presence and length |
| `_find_keyword_count(document)` | `int \| None` | Locate keyword section, count items |
| `_check_keywords(document, word_count)` | `list[EumicViolationDTO]` | Keyword presence and count |

Total: 16 private methods. The `inspect()` public method plus these 16 private methods constitute the full `DocxEumicAdapter` interface.

---

## 3. File Structure (All 11 New Files)

All directories already exist with the necessary `__init__.py` files. No new directories are needed.

| # | Full Path | Description |
|---|---|---|
| 1 | `src/domain/dtos/eumic_violation_dto.py` | `EumicViolationDTO(BaseDTO)` — frozen dataclass with category, message, severity, details |
| 2 | `src/domain/document/document_format_inspection_port.py` | `DocumentFormatInspectionPort(ABC)` — single `inspect(docx_path, word_count)` abstract method |
| 3 | `src/application/verify_eumic_use_case.py` | `VerifyEumicUseCase` — thin orchestration, delegates to port, `@generic_error_handler` on `execute()` |
| 4 | `src/infrastructure/adapters/document/docx_eumic_adapter.py` | `DocxEumicAdapter(DocumentFormatInspectionPort)` — 13 named constants + 1 public + 16 private methods |
| 5 | `src/infrastructure/wirings/verify_eumic_use_case_wiring.py` | `VerifyEumicUseCaseWiring` — `create_use_case()` public, `_get_document_format_inspection_port()` private |
| 6 | `src/domain/tests/dtos/test_eumic_violation_dto.py` | Unit tests for `EumicViolationDTO` (5 test methods) |
| 7 | `src/domain/tests/document/test_document_format_inspection_port.py` | ABC instantiation tests + fake port interface check |
| 8 | `src/domain/tests/document/fake_document_format_inspection_port.py` | Configurable test double: accepts return list or exception to raise |
| 9 | `src/application/tests/test_verify_eumic_use_case.py` | Use case tests via fake port (3 test methods) |
| 10 | `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py` | Integration tests for wiring (2 test methods) |
| 11 | `src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py` | Adapter tests with mocked python-docx (12+ test methods) |

---

## 4. Dependency Graph

```
src/domain/enums/severity_level.py (existing)
    ↑
src/domain/dtos/eumic_violation_dto.py (new)
    ↑
src/domain/document/document_format_inspection_port.py (new)
    ↑                                    ↑
src/application/verify_eumic_use_case.py   src/infrastructure/adapters/document/docx_eumic_adapter.py
    ↑                                          ↑
src/infrastructure/wirings/verify_eumic_use_case_wiring.py (new)

─── test dependencies ───

fake_document_format_inspection_port.py → implements DocumentFormatInspectionPort
test_verify_eumic_use_case.py → uses FakeDocumentFormatInspectionPort
test_docx_eumic_adapter.py → patches docx.Document, uses EumicViolationDTO assertions
test_verify_eumic_use_case_wiring.py → instantiates real VerifyEumicUseCaseWiring

─── external ───

docx_eumic_adapter.py
  ← python-docx (Document, Cm, Pt, WD_ALIGN_PARAGRAPH, PackageNotFoundError)
  ← src/domain/exceptions/document_errors.py (existing: DocumentUnreadable)
  ← src/domain/exceptions/decorators/generic_error_handler.py (existing)
```

Import invariants:
- `src/domain/` files import only from `src/domain/` and stdlib.
- `src/application/` imports from `src/domain/` only.
- `src/infrastructure/` imports from all layers and external libraries.

---

## 5. Test Design

### `test_eumic_violation_dto.py`

Pure unit tests. No mocking required.

| Test method | What it verifies |
|---|---|
| `test_creates_with_required_fields` | Fields accept correct types |
| `test_details_defaults_to_empty_string` | `details` has default `""` |
| `test_is_frozen` | `FrozenInstanceError` raised on mutation attempt |
| `test_is_subclass_of_base_dto` | Inheritance chain is correct |
| `test_severity_accepts_severity_level_enum` | `SeverityLevel.WARNING` is accepted |

### `test_document_format_inspection_port.py`

| Test method | What it verifies |
|---|---|
| `test_cannot_instantiate_directly` | `TypeError` on `DocumentFormatInspectionPort()` |
| `test_fake_port_satisfies_interface` | `FakeDocumentFormatInspectionPort` instantiates and returns `list[EumicViolationDTO]` |

### `fake_document_format_inspection_port.py`

```python
class FakeDocumentFormatInspectionPort(DocumentFormatInspectionPort):
    def __init__(
        self,
        violations: list[EumicViolationDTO] | None = None,
        exception_to_raise: Exception | None = None,
    ) -> None:
        self._violations = violations or []
        self._exception_to_raise = exception_to_raise

    def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        if self._exception_to_raise is not None:
            raise self._exception_to_raise
        return self._violations
```

### `test_verify_eumic_use_case.py`

Uses `FakeDocumentFormatInspectionPort`. No mocking of external libraries.

| Test method | What it verifies |
|---|---|
| `test_returns_empty_list_when_no_violations` | Port returns `[]`; use case returns `[]` |
| `test_returns_violations_from_port` | Port returns violations; use case propagates them |
| `test_propagates_document_unreadable_as_src_generic_error` | Port raises `DocumentUnreadable`; `@generic_error_handler` re-raises as-is (it is a `BaseSrcError` subclass) |

Note: `DocumentUnreadable` is a `BaseSrcError` subclass. `@generic_error_handler` re-raises `BaseSrcError` subclasses as-is (not wrapped in `SrcGenericError`). The test must assert `DocumentUnreadable` is raised, not `SrcGenericError`.

### `test_docx_eumic_adapter.py`

**Mock strategy**: Patch `src.infrastructure.adapters.document.docx_eumic_adapter.Document` at the module level so `Document(docx_path)` inside `inspect()` returns a controlled mock. Each test method sets up the mock document object.

**Mock objects needed per check method**:

| Check | python-docx objects to mock |
|---|---|
| `_verify_format` | `document.sections[0]` with `.top_margin`, `.bottom_margin`, `.left_margin`, `.right_margin` (each having `.twips` and `.cm`); `document.paragraphs` list of mocks with `.runs` (each run having `.font.name`, `.font.size`); `.alignment` on paragraphs |
| `_verify_figures` | `document.part.rels.values()` list of mocks with `.target_ref`; paragraphs with `.text.strip()` returning figure caption text |
| `_verify_tables` | `document.tables` list of mock tables; paragraphs with table title text |
| `_verify_formulas` | Paragraphs with runs having `._element.xml` returning `'<m:oMath ...'`; paragraph `.alignment` |
| `_verify_abstract_keywords` | Paragraphs with text containing abstract/keyword section headers |

**Mock margin value helper**:

```python
def _mock_length(cm_value: float) -> MagicMock:
    length = MagicMock()
    length.cm = cm_value
    length.twips = int(cm_value * 567)   # docx twips per cm
    return length
```

**Test method inventory**:

| Test method | Scenario | Expected result |
|---|---|---|
| `test_returns_empty_list_for_compliant_document` | All checks pass | `[]` |
| `test_raises_document_unreadable_for_invalid_path` | `PackageNotFoundError` from `Document()` | `DocumentUnreadable` |
| `test_margin_violation_returned_when_margin_below_threshold` | `top_margin.twips` below threshold | 1 WARNING violation, category "Formato General" |
| `test_font_violation_returned_for_non_standard_font` | Run with font "Comic Sans" | 1 WARNING violation |
| `test_font_size_violation_returned_for_non_standard_size` | Run with 8pt font | 1 INFO violation |
| `test_figure_count_mismatch_returns_violation` | 2 images, 1 caption | 1 WARNING violation, category "Figuras" |
| `test_figure_numbering_inconsistency_returns_violation` | Captions "Figura 1", "Figura 3" | 1 WARNING violation |
| `test_table_count_mismatch_returns_violation` | 2 tables, 1 title | 1 WARNING violation, category "Tablas" |
| `test_uncentered_formula_returns_violation` | Formula run, non-CENTER alignment | 1 INFO violation, category "Fórmulas" |
| `test_abstract_missing_on_long_document_returns_critical` | No abstract, `word_count=2000` | 1 CRITICAL violation |
| `test_abstract_check_skipped_on_short_document` | No abstract, `word_count=500` | No violation for abstract |
| `test_keywords_missing_on_long_document_returns_critical` | No keywords section, `word_count=2000` | 1 CRITICAL violation |
| `test_wrong_keyword_count_returns_warning` | Keywords section with 1 keyword | 1 WARNING violation |

**Constants test approach**: Constants are not tested directly. They are verified implicitly through the behavior tests — e.g., a margin of exactly `2.5 cm` passes, `2.9 cm` triggers a warning (because `0.4 cm > MARGIN_TOLERANCE_CM`). This is sufficient for coverage without brittle constant-value snapshot tests.

### `test_verify_eumic_use_case_wiring.py`

| Test method | What it verifies |
|---|---|
| `test_create_use_case_returns_verify_eumic_use_case_instance` | Returns `VerifyEumicUseCase` instance |
| `test_create_use_case_injects_docx_eumic_adapter_as_format_inspection_port` | `._format_inspection_port` is a `DocxEumicAdapter` |

---

## 6. Coexistence Strategy

`eumic_verifier.py` (root) and `main.py` are **not modified in Slice 11**. The legacy call path remains fully functional:

```
main.py  →  eumic_verifier.py.verify_eumic_compliance()  (unchanged)
```

The new hexagonal path exists in parallel:

```
VerifyEumicUseCaseWiring → VerifyEumicUseCase → DocxEumicAdapter
```

**What changes in Slice 14** (outside this slice's scope):
- `main.py` replaces its `verify_eumic_compliance(doc, document_content)` call with `VerifyEumicUseCaseWiring().create_use_case().execute(docx_path, word_count)`.
- `format_violations_report()` logic moves to a controller/presenter layer.
- `eumic_verifier.py` is deleted or archived.

**Slice 11 guarantee**: running `python -m pytest src/ -q` after implementation must show exactly 11 new test files and the existing 459 tests still passing. The pass count increases; no test changes or deletions.

---

## 7. TDD Order (Strict TDD — tests before implementation)

```
Step 1: test_eumic_violation_dto.py         → eumic_violation_dto.py
Step 2: test_document_format_inspection_port.py
        + fake_document_format_inspection_port.py  → document_format_inspection_port.py
Step 3: test_verify_eumic_use_case.py       → verify_eumic_use_case.py
Step 4: test_docx_eumic_adapter.py          → docx_eumic_adapter.py
Step 5: test_verify_eumic_use_case_wiring.py → verify_eumic_use_case_wiring.py
```

Each step: write the test file first (all tests must fail), then write the implementation (all tests must pass), then run the full suite before proceeding.

---

## 8. Risks and Mitigations

| Risk | Severity | Concrete Mitigation in Design |
|---|---|---|
| **Local imports in source** | Medium | All `docx.shared`, `docx.enum.text` imports specified at module top of `DocxEumicAdapter`. Tasks phase lists each import line. |
| **`run._element.xml` internal attribute** | Low | Isolated in `_run_contains_omath()`. The bytes/str guard is preserved verbatim. The method's docstring documents the python-docx internal access as a known dependency. |
| **Bare `except: pass` in `_verify_figures`** | Low | Replaced with `except Exception: pass` in `_count_image_relationships()`. Silent behavior preserved. No logging. |
| **`margin_value.twips` mock complexity** | Medium | `_mock_length(cm_value)` helper in test class computes `.twips = int(cm_value * 567)` to match docx's conversion. The mock-level contract is documented in the test design above. |
| **Regression on 459 existing tests** | High | Coexistence: zero modifications to existing files. The 11 new files are additive only. |
| **`format_violations_report` has no hexagonal home** | Accepted | Deferred to Slice 14. The legacy call path is untouched. Not a Slice 11 risk. |
| **`_check_sequential_numbering_violations` shared by figures and tables** | Low | The helper takes `pattern`, `category`, `message`, `details` as arguments. Both callers pass distinct values. The guard `if len(captions) <= 1: return []` prevents false positives when there is only one caption to check. |

---

## Appendix: Resolved Questions from Proposal

| Open question (proposal §12) | Resolution in this design |
|---|---|
| Bare except replacement | `except Exception: pass` in `_count_image_relationships()`. No logging. |
| `src/infrastructure/tests/adapters/document/__init__.py` existence | Confirmed: exists. No new `__init__.py` files needed anywhere. |
| Additional numeric literals | Fully audited. 9 additional constants beyond the 4 from the proposal. All listed in Section 2 constants table. |
| `VerifyEumicUseCaseWiringForTest` | Excluded. Pattern not established in this project. No concrete scenario requires it in Slice 11. |
