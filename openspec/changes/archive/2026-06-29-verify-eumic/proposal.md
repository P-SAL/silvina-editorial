# SDD Proposal — verify-eumic (Slice 11)

**Change name**: verify-eumic
**Slice**: 11
**Branch**: refactor/hexagonal-migration
**Author**: SDD propose phase
**Date**: 2026-06-29
**Status**: Proposed

---

## 1. Intent

Slice 11 migrates the format-compliance logic of `eumic_verifier.py` (root-level legacy file) into the hexagonal architecture. The result is a new `VerifyEumicUseCase` backed by a `DocumentFormatInspectionPort`, implemented by a `DocxEumicAdapter` that houses all five check methods from the original `EumicVerifier` class.

**Why now**: The hexagonal migration is progressing slice by slice. Grammar checking (Slice 10), citation extraction (Slice 7), and reference extraction were already migrated. EUMIC compliance checking is the next standalone domain capability that can be migrated independently while `main.py` keeps calling the legacy `eumic_verifier.py` (coexistence strategy).

**What success looks like**:
- All five EUMIC verification methods live inside `DocxEumicAdapter` under `src/infrastructure/adapters/document/`.
- A clean `DocumentFormatInspectionPort` defines the contract in the domain layer.
- `VerifyEumicUseCase` acts as thin orchestration: no domain logic, delegates entirely to the port.
- All magic numbers and string literals encoding EUMIC editorial standards are extracted into named module-level constants.
- 459 existing tests continue passing; 11 new test files are added and pass in Strict TDD order (tests written before implementation).
- `eumic_verifier.py` at the root remains unchanged and functional; `main.py` continues calling it without modification.

---

## 2. Scope

### In scope — 11 new files (all greenfield)

#### Source files (5)

| File | Artifact |
|---|---|
| `src/domain/dtos/eumic_violation_dto.py` | `EumicViolationDTO` dataclass |
| `src/domain/document/document_format_inspection_port.py` | `DocumentFormatInspectionPort` ABC |
| `src/application/verify_eumic_use_case.py` | `VerifyEumicUseCase` |
| `src/infrastructure/adapters/document/docx_eumic_adapter.py` | `DocxEumicAdapter` (5 check methods + named constants) |
| `src/infrastructure/wirings/verify_eumic_use_case_wiring.py` | `VerifyEumicUseCaseWiring` |

#### Test files (6)

| File | Artifact |
|---|---|
| `src/domain/tests/dtos/test_eumic_violation_dto.py` | Unit tests for `EumicViolationDTO` |
| `src/domain/tests/document/test_document_format_inspection_port.py` | Unit tests for the port ABC |
| `src/domain/tests/document/fake_document_format_inspection_port.py` | Configurable test double (fake port) |
| `src/application/tests/test_verify_eumic_use_case.py` | Unit tests for the use case |
| `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py` | Integration tests for the wiring |
| `src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py` | Adapter tests (mocked python-docx) |

#### Supporting `__init__.py` files

Any `__init__.py` files needed to make new directories importable (e.g., `src/infrastructure/tests/adapters/`, `src/infrastructure/tests/adapters/document/`) are in scope as mechanical additions with no logic.

### Existing files modified

None. Slice 11 is strictly additive. No existing file is modified. This is the coexistence strategy: the legacy `eumic_verifier.py` and `main.py` continue working without change until Slice 14.

---

## 3. Out of Scope

| Item | Deferred to |
|---|---|
| `format_violations_report()` presentation logic (emoji formatting, severity grouping into a report string) | Slice 14 (controller migration) |
| `main.py` migration — replacing the `verify_eumic_compliance()` call with `VerifyEumicUseCaseWiring` | Slice 14 |
| Modifying `eumic_verifier.py` root file in any way | Not planned; kept as-is indefinitely until Slice 14 deletes it |
| New domain exception class for EUMIC-specific errors | Not needed; `DocumentUnreadable` (already in `document_errors.py`) covers docx open failure; `@generic_error_handler` covers unexpected infrastructure exceptions |
| New `EumicSeverity` enum | Not needed; `SeverityLevel` (already at `src/domain/enums/severity_level.py`) covers the three EUMIC severity levels (INFO / WARNING / CRITICAL) |
| Gradio app integration | `gradio_app.py` does not use `eumic_verifier.py` — no changes needed |
| Additional EUMIC check methods beyond the original five | Post-Slice 11 feature work |

---

## 4. Named Constants for Editorial Standards

### Requirement

All magic numbers and string literals encoding EUMIC editorial policy must be extracted into named module-level constants. No literal `2.5`, `12`, `1000`, `"Times New Roman"`, etc. may appear inside method bodies.

### Placement decision: `docx_eumic_adapter.py` (module level)

These constants configure how the adapter performs its inspections. They are placed at the top of `src/infrastructure/adapters/document/docx_eumic_adapter.py` as module-level constants, before any class definition.

**Justification for adapter placement**:
- The port interface (`DocumentFormatInspectionPort`) defines the contract (`inspect → list[EumicViolationDTO]`) without encoding what the thresholds are. The thresholds are an implementation detail of *this* adapter — they answer "how does the DocxEumicAdapter decide if a margin is wrong?", not "what is the domain model for a violation?".
- Placing them in the adapter respects the YAGNI principle: there is currently one adapter and no planned second adapter for EUMIC inspection. If a second adapter appears, these constants can be extracted into a domain constants module at that time.
- The domain layer remains free of adapter-specific constants, preserving import invariant purity (domain has no knowledge of adapter internals).
- The adapter is allowed to import from the domain layer (allowed by import invariants); the domain is forbidden from importing adapter constants. Placing constants in the adapter is therefore the only direction that is architecturally safe without an additional indirection layer.

### Constant definitions (in `docx_eumic_adapter.py`)

```python
REQUIRED_MARGIN_CM: float = 2.5
REQUIRED_FONT_SIZE_PT: int = 12
ALLOWED_FONTS: frozenset[str] = frozenset({"Times New Roman", "Arial", "Calibri"})
MIN_WORDS_FOR_ABSTRACT_CHECK: int = 1000
```

### Additional constants to identify during implementation

The `sdd-spec` phase must audit the full `EumicVerifier` source and extract any remaining numeric or string literals representing business rules. Known candidates beyond the above four:
- Sequential figure/table numbering expectations (any hardcoded indices or comparison values in `_verify_figures` / `_verify_tables`)
- Formula alignment expected values (if `WD_ALIGN_PARAGRAPH.CENTER` is compared against a hardcoded value)
- Keyword and abstract minimum/maximum count values (if any beyond `MIN_WORDS_FOR_ABSTRACT_CHECK`)

The spec phase will enumerate all remaining literals and assign a constant name for each.

---

## 5. Port Interface

**File**: `src/domain/document/document_format_inspection_port.py`
**Class**: `DocumentFormatInspectionPort(ABC)`

```python
@abstractmethod
def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
    """Return EUMIC format violations found in the document at docx_path."""
```

**Signature rationale**:
- `docx_path: str` — consistent with `CitationExtractionPort` and `ReferenceExtractionPort` conventions in the project. The adapter opens the file internally.
- `word_count: int` — primitive instead of the full `DocumentContentDTO`. The adapter only reads `word_count` from that DTO (in `_verify_abstract_keywords`). Coupling the port to the full DTO would violate Interface Segregation.
- Return type `list[EumicViolationDTO]` — empty list signals no violations (not an exception). Callers handle empty list as "document is compliant".

---

## 6. DTO

**File**: `src/domain/dtos/eumic_violation_dto.py`
**Class**: `EumicViolationDTO(BaseDTO)` — frozen dataclass

Fields:
- `category: str` — which EUMIC check raised this violation (e.g., `"format"`, `"figures"`, `"tables"`, `"formulas"`, `"abstract_keywords"`)
- `message: str` — human-readable violation description
- `severity: SeverityLevel` — reuses `src/domain/enums/severity_level.py`; no new enum
- `details: str = ""` — optional extra context (default empty string)

**No new enum**: The existing `SeverityLevel` (INFO / WARNING / ERROR / CRITICAL) covers the three EUMIC severity levels used in the original code (INFO / WARNING / CRITICAL). The original `EumicSeverity` enum with emoji-prefixed string values (`"🔴 CRÍTICO"`) is a presentation concern and is not migrated into the domain.

---

## 7. Use Case

**File**: `src/application/verify_eumic_use_case.py`
**Class**: `VerifyEumicUseCase`

```python
class VerifyEumicUseCase:
    def __init__(self, format_inspection_port: DocumentFormatInspectionPort) -> None:
        self._format_inspection_port = format_inspection_port

    @generic_error_handler
    def execute(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        """Execute EUMIC compliance inspection and return violations."""
        return self._format_inspection_port.inspect(
            docx_path=docx_path,
            word_count=word_count,
        )
```

**Design decisions**:
- `@generic_error_handler` on `execute()`: wraps infrastructure exceptions from the adapter, re-raises `BaseSrcError` subclasses as-is, wraps unexpected exceptions in `SrcGenericError`. Consistent with all other use cases in the project.
- No domain computation: the use case is purely orchestration. All five check methods are inside the adapter; the use case delegates entirely.
- Single port dependency: `DocumentFormatInspectionPort`. No second port, no repository.

---

## 8. Adapter

**File**: `src/infrastructure/adapters/document/docx_eumic_adapter.py`
**Class**: `DocxEumicAdapter(DocumentFormatInspectionPort)`

**Responsibilities**:
- Implements `DocumentFormatInspectionPort`.
- Opens the docx file via `python-docx` (`Document(docx_path)`).
- Raises `DocumentUnreadable` (from `src/domain/exceptions/document_errors.py`) if the file cannot be opened.
- Delegates to five private check methods, collecting `EumicViolationDTO` results.
- `inspect()` is decorated with `@generic_error_handler`.

**Five check methods migrated from `EumicVerifier`**:

| Method | Checks |
|---|---|
| `_verify_format(document)` | Margins (uses `REQUIRED_MARGIN_CM`), fonts (uses `ALLOWED_FONTS`), font size (uses `REQUIRED_FONT_SIZE_PT`), text justification |
| `_verify_figures(document)` | Image relationship count vs captioned figure count; sequential figure numbering |
| `_verify_tables(document)` | Table title count vs table count; sequential table numbering |
| `_verify_formulas(document)` | OMath XML detection via `run._element.xml`; formula paragraph alignment |
| `_verify_abstract_keywords(document, word_count)` | Abstract presence and minimum word count; keyword presence and count (only when `word_count >= MIN_WORDS_FOR_ABSTRACT_CHECK`) |

**Local import fixes (SKILL §3 compliance)**:
The original `eumic_verifier.py` uses local imports inside methods (`from docx.shared import Pt, Cm` and `from docx.enum.text import WD_ALIGN_PARAGRAPH`). All imports are moved to the module top-level in `DocxEumicAdapter`.

**Replacement of legacy types**:
- `EumicViolation` (old dataclass) → `EumicViolationDTO`
- `EumicSeverity.INFO / WARNING / CRITICAL` → `SeverityLevel.INFO / WARNING / CRITICAL`

**Preserved verbatim**:
- The `run._element.xml` access with `isinstance(xml_str, bytes)` decode guard (internal python-docx attribute; behavior must be preserved for parity).
- The `doc.part.rels` access pattern in `_verify_figures`.

---

## 9. Wiring

**File**: `src/infrastructure/wirings/verify_eumic_use_case_wiring.py`
**Class**: `VerifyEumicUseCaseWiring`

```python
class VerifyEumicUseCaseWiring:
    def get_verify_eumic_use_case(self) -> VerifyEumicUseCase:
        """Return a fully assembled VerifyEumicUseCase."""
        return VerifyEumicUseCase(
            format_inspection_port=self._get_document_format_inspection_port(),
        )

    def _get_document_format_inspection_port(self) -> DocumentFormatInspectionPort:
        return DocxEumicAdapter()
```

**Convention notes**:
- Public method named `get_verify_eumic_use_case()` per SKILL §8 convention (`get_<use_case_snake_case>()`).
- Private method returns the port type (`DocumentFormatInspectionPort`), not the concrete adapter.
- No business logic — only object creation and wiring.

---

## 10. Test Plan (Strict TDD)

Tests are written **before** the corresponding implementation file. The implementation may only exist once its test file is written and failing.

### TDD order

```
1.  test_eumic_violation_dto.py          → eumic_violation_dto.py
2.  test_document_format_inspection_port.py + fake_document_format_inspection_port.py  → document_format_inspection_port.py
3.  test_verify_eumic_use_case.py        → verify_eumic_use_case.py
4.  test_docx_eumic_adapter.py           → docx_eumic_adapter.py
5.  test_verify_eumic_use_case_wiring.py → verify_eumic_use_case_wiring.py
```

### Test file details

#### `test_eumic_violation_dto.py`
- `test_creates_with_required_fields` — category, message, severity can be set
- `test_details_defaults_to_empty_string` — default value
- `test_is_frozen` — `dataclasses.FrozenInstanceError` on mutation attempt
- `test_is_subclass_of_base_dto`
- `test_severity_accepts_severity_level_enum`

#### `test_document_format_inspection_port.py`
- `test_cannot_instantiate_directly` — `TypeError` on `DocumentFormatInspectionPort()`
- `test_fake_port_satisfies_interface` — `FakeDocumentFormatInspectionPort` instantiates and returns `list[EumicViolationDTO]`

#### `fake_document_format_inspection_port.py`
- Configurable: accepts a list of `EumicViolationDTO` to return, or an exception to raise on `inspect()`.
- No I/O — pure test double.

#### `test_verify_eumic_use_case.py`
- `test_returns_empty_list_when_no_violations` — fake port returns `[]`
- `test_returns_violations_from_port` — fake port returns a list; use case propagates it
- `test_propagates_exception_from_port` — fake port raises; `@generic_error_handler` wraps it in `SrcGenericError`

#### `test_docx_eumic_adapter.py`
- Uses `unittest.mock.MagicMock` to mock `python-docx` `Document` objects (following pattern from the existing `tests/test_eumic_verifier.py`)
- Covers all five check methods with representative scenarios:
  - Correct format → no violations
  - Margin below threshold → INFO/WARNING violation returned
  - Disallowed font → violation returned
  - Wrong font size → violation returned
  - Figure count mismatch → violation returned
  - Table count mismatch → violation returned
  - OMath detected but not center-aligned → violation returned
  - Document shorter than `MIN_WORDS_FOR_ABSTRACT_CHECK` → abstract check skipped
  - Abstract missing on long document → violation returned
  - Keywords missing on long document → violation returned
- `test_raises_document_unreadable_when_docx_path_is_invalid` — verifies the error mapping

#### `test_verify_eumic_use_case_wiring.py`
- `test_wiring_returns_verify_eumic_use_case_instance`
- `test_wiring_injects_docx_eumic_adapter_as_format_inspection_port`

---

## 11. Risks and Mitigations

| Risk | Severity | Mitigation |
|---|---|---|
| **Local imports in source** — `from docx.shared import Pt, Cm` and `from docx.enum.text import WD_ALIGN_PARAGRAPH` are inside methods in `eumic_verifier.py` (SKILL §3 violation) | Medium | Move all to module top in `DocxEumicAdapter`. Spec phase lists exact import lines. |
| **`run._element.xml` internal attribute** — accesses python-docx internals; may break on version upgrade | Low | Preserve the existing defensive guard (`isinstance(xml_str, bytes)` decode). Document this access as a known dependency on python-docx internals in the adapter's docstring. |
| **Bare `except: pass` in `_verify_figures`** — swallows errors silently | Low | Replace with `except Exception:` and optionally log. Behavior parity preserved (do not re-raise). Spec phase decides exact replacement. |
| **`margin_value.cm` attribute mocking** — docx `Length` objects need careful mocking in unit tests | Medium | Follow the `MagicMock` pattern from `tests/test_eumic_verifier.py`. Spec phase defines exact mock setup for `_verify_format` tests. |
| **`__init__.py` gaps in new test subdirectory** — `src/infrastructure/tests/adapters/document/` may not have `__init__.py` files, causing test discovery failures | Low | Create `__init__.py` files for each new subdirectory as part of the additive scope. |
| **Regression risk** — 459 tests currently passing; adapter port contract must not break callers | High | Coexistence strategy guarantees no modifications to existing files. Run `python -m pytest src/ -q` after each TDD step. |
| **`format_violations_report` has no hexagonal home** — main.py still calls the old `verify_eumic_compliance()` → `format_violations_report()` chain | Accepted | Deferred to Slice 14 by design. The old code path remains untouched. No risk to Slice 11. |

---

## 12. Non-Goals

The following are explicitly NOT part of Slice 11, even if they are related:

- **No modification to any existing file** (no `eumic_verifier.py`, no `main.py`, no `document_errors.py`, no `severity_level.py`, no existing `__init__.py` logic — only new `__init__.py` files in new directories)
- **No new exception class** — `DocumentUnreadable` (existing) is sufficient
- **No new enum** — `SeverityLevel` (existing) is sufficient
- **No `format_violations_report` migration** — presentation concern; Slice 14
- **No Gradio app changes** — `gradio_app.py` does not use EUMIC verification
- **No change to `main.py`** — coexistence strategy; Slice 14
- **No additional EUMIC check methods** — only migrate the existing five
- **No database or external service** — pure docx file inspection, no persistence

---

## Appendix: Dependency and Call Flow

```
main.py (legacy, unchanged)
  └─ eumic_verifier.py (legacy, unchanged) ← Slice 14 replaces this

VerifyEumicUseCaseWiring
  └─ VerifyEumicUseCase
       └─ DocumentFormatInspectionPort (ABC)
            └─ DocxEumicAdapter (implements)
                 ├─ REQUIRED_MARGIN_CM, REQUIRED_FONT_SIZE_PT, ALLOWED_FONTS, MIN_WORDS_FOR_ABSTRACT_CHECK
                 ├─ _verify_format()
                 ├─ _verify_figures()
                 ├─ _verify_tables()
                 ├─ _verify_formulas()
                 └─ _verify_abstract_keywords()
                      └─ EumicViolationDTO (uses SeverityLevel enum)
```

---

## Proposal Question Round

Before finalizing, the following questions remain open. The `sdd-spec` phase should resolve them:

1. **Bare `except: pass` in `_verify_figures`**: replace with `except Exception: pass` silently, or add a logging call? The spec should decide the exact replacement and whether a `DocumentProcessingWarning` (or similar) is appropriate.

2. **`src/infrastructure/tests/adapters/document/`**: Does this subdirectory already exist with `__init__.py` files from a prior slice? The spec phase should confirm directory state before generating file lists.

3. **Additional numeric literals**: Are there comparison values in `_verify_figures` / `_verify_tables` (e.g., hardcoded expected sequential numbering) that require named constants beyond the four listed? The spec phase must audit all five check methods exhaustively.

4. **Wiring test double**: Does Slice 11 need a `VerifyEumicUseCaseWiringForTest` (SKILL §8 pattern)? No application-level integration test has been planned that would need to swap the adapter. Spec phase should confirm whether to include it or defer.
