# Exploration: verify-eumic

**Change**: verify-eumic — Slice 11 hexagonal migration of `eumic_verifier.py`
**Date**: 2026-06-29
**Status**: Complete

---

## eumic_verifier.py Analysis

**5 responsibilities** (5 internal check methods):
1. `_verify_format(doc)` — margins (2.5cm required), fonts (Times/Arial/Calibri), font size (12pt), text justification
2. `_verify_figures(doc)` — counts image relationships vs captioned figures, checks sequential numbering
3. `_verify_tables(doc)` — table title count vs table count, sequential numbering
4. `_verify_formulas(doc)` — detects OMath XML in runs, checks formula paragraph alignment
5. `_verify_abstract_keywords(doc, document_content)` — abstract presence and word count, keyword presence and count (only for docs > 1000 words)

**Public API used by callers:**
- `verify_eumic_compliance(doc, document_content) -> str` — convenience function called only by `main.py` (line 256)
- `gradio_app.py` does NOT use eumic_verifier at all

**`document_content` dependency:** only `.word_count` is read (in `_verify_abstract_keywords`)

**Local imports (SKILL §3 violation to fix in adapter):**
- `from docx.shared import Pt, Cm` — inside `_verify_format`
- `from docx.enum.text import WD_ALIGN_PARAGRAPH` — inside `_verify_format` and `_verify_formulas`

**`format_violations_report(violations)` method:** PRESENTATION logic — formats violations into a report string with emojis. This is a CONTROLLER concern in hexagonal. The use case returns `list[EumicViolationDTO]`; main.py handles formatting.

---

## Port Design: DocumentFormatInspectionPort

File: `src/domain/document/document_format_inspection_port.py`

```python
from abc import ABC, abstractmethod
from src.domain.dtos.eumic_violation_dto import EumicViolationDTO

class DocumentFormatInspectionPort(ABC):
    """Port for inspecting document format compliance."""

    @abstractmethod
    def inspect(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        """Return EUMIC format violations found in the document."""
```

**Signature rationale:** `docx_path: str` (adapter opens file internally, consistent with existing ports). `word_count: int` primitive — only this field is needed from `document_content`. Full DTO coupling would violate ISP.

---

## New DTO: EumicViolationDTO

File: `src/domain/dtos/eumic_violation_dto.py`

```python
@dataclass(frozen=True)
class EumicViolationDTO(BaseDTO):
    category: str
    message: str
    severity: SeverityLevel  # reuse existing src/domain/enums/severity_level.py
    details: str = ""
```

**No new enum needed**: existing `SeverityLevel` (INFO/WARNING/ERROR/CRITICAL) covers the three EUMIC severity levels. The old `EumicSeverity` enum with emoji string values is a presentation concern — discarded.

---

## Use Case Design: VerifyEumicUseCase

File: `src/application/verify_eumic_use_case.py`

```python
class VerifyEumicUseCase:
    def __init__(self, format_inspection_port: DocumentFormatInspectionPort) -> None:
        self._format_inspection_port = format_inspection_port

    @generic_error_handler
    def execute(self, docx_path: str, word_count: int) -> list[EumicViolationDTO]:
        return self._format_inspection_port.inspect(docx_path=docx_path, word_count=word_count)
```

Thin orchestration — delegates entirely to the port. `@generic_error_handler` handles infrastructure exceptions.

---

## Adapter Design: DocxEumicAdapter

File: `src/infrastructure/adapters/document/docx_eumic_adapter.py`

- Implements `DocumentFormatInspectionPort`
- Opens the docx file internally via python-docx
- Migrates all 5 check methods from `EumicVerifier` verbatim (with fixes: move local imports to top, replace `EumicViolation` with `EumicViolationDTO`, replace `EumicSeverity` with `SeverityLevel`)
- Decorates `inspect()` with `@generic_error_handler`
- Raises `DocumentUnreadable` on docx open failure

---

## Wiring Design

File: `src/infrastructure/wirings/verify_eumic_use_case_wiring.py`

```python
class VerifyEumicUseCaseWiring:
    def create_use_case(self) -> VerifyEumicUseCase:
        return VerifyEumicUseCase(format_inspection_port=self._get_document_format_inspection_port())

    def _get_document_format_inspection_port(self) -> DocumentFormatInspectionPort:
        return DocxEumicAdapter()
```

---

## Test Strategy (Strict TDD — write tests FIRST)

**11 new files total:**

Domain tests (pure, no I/O):
- `src/domain/tests/dtos/test_eumic_violation_dto.py`
- `src/domain/tests/document/test_document_format_inspection_port.py`
- `src/domain/tests/document/fake_document_format_inspection_port.py`

Application tests (use case with fake port):
- `src/application/tests/test_verify_eumic_use_case.py`

Infrastructure tests:
- `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py`
- `src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py`

---

## Risks

1. **Local import violation**: `Pt, Cm` and `WD_ALIGN_PARAGRAPH` are inside method bodies. Must move to module top in adapter (SKILL §3).
2. **`format_violations_report` is presentation logic**: Deferred to Slice 14 (CLI controller migration). Not a blocker for Slice 11.
3. **`margin_value.cm` attribute access**: Needs careful mocking in unit tests. Follow pattern from `tests/test_eumic_verifier.py`.
4. **`doc.part.rels` bare `except: pass`**: Make more specific in adapter while preserving behavior.
5. **OMath XML detection uses `run._element.xml`**: Internal python-docx attribute. Bytes/str defensive check must be preserved verbatim.
6. **Integration test complexity**: Adapter tests need mocked python-docx Document objects (follow existing `tests/test_eumic_verifier.py` pattern).

---

## File Inventory

New files (all greenfield):
- `src/domain/dtos/eumic_violation_dto.py`
- `src/domain/document/document_format_inspection_port.py`
- `src/application/verify_eumic_use_case.py`
- `src/infrastructure/adapters/document/docx_eumic_adapter.py`
- `src/infrastructure/wirings/verify_eumic_use_case_wiring.py`
- `src/domain/tests/dtos/test_eumic_violation_dto.py`
- `src/domain/tests/document/test_document_format_inspection_port.py`
- `src/domain/tests/document/fake_document_format_inspection_port.py`
- `src/application/tests/test_verify_eumic_use_case.py`
- `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py`
- `src/infrastructure/tests/adapters/document/test_docx_eumic_adapter.py`

No modifications to existing files (coexistence strategy).
No new enum file. No new exception file.
