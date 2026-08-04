# Design: check-grammar (Slice 10 — Hexagonal Migration)

**Change name**: check-grammar
**Slice**: 10 of N (incremental hexagonal migration)
**Date**: 2026-06-28
**Status**: designed

---

## Technical Approach

Approach A: `GrammarCheckPort` returns raw `list[GrammarErrorDTO]`; scoring and feedback logic live in `CheckGrammarUseCase`. `LanguageToolAdapter` uses lazy initialization (`if self._tool is None`) to defer Java startup to the first `check()` call. 18 new files, 0 modified, single PR targeting `refactor/hexagonal-migration`.

---

## Architecture Decisions

| Decision | Choice | Rejected | Rationale |
|----------|--------|----------|-----------|
| Scoring placement | `CheckGrammarUseCase._calculate_score()` / `_build_feedback()` | Adapter (Approach B), domain service (Approach C) | Scoring is business logic; adapters return raw data; 3 thresholds do not warrant a dedicated domain service |
| Lazy LanguageTool init | `self._tool = None` in `__init__`; initialized inside `check()` on first call | Eager `__init__` init | Eager init starts Java on test collection, breaking isolation; local import violates SKILL §3 |
| `language` constructor param | `language: str = "es"` in `LanguageToolAdapter.__init__` | Hardcoded `"es"` | Wiring uses default; parameter allows future extensibility at no cost |
| Port location | `src/domain/grammar/grammar_check_port.py` | `src/domain/ports/` (Ollama-era, pre-Slice 5) | Entity-scoped ports are the established convention from Slices 5–9 |
| `@generic_error_handler` scope | Adapter `check()` only | Also on use case `execute()` | Proposal is explicit; use case logic (threshold, DTO construction) cannot raise unexpected exceptions; double-wrapping is redundant |
| Wiring public method | `create_use_case()` | `get_check_grammar_use_case()` (SKILL §8 text) | All 6 existing wirings use `create_use_case()`; codebase convention overrides SKILL §8 text |
| Fake location | `src/domain/tests/grammar/fake_grammar_check_port.py` | `infrastructure/tests/test_doubles/` | All existing fakes live under `domain/tests/<entity>/`; no test_doubles folder exists |
| Init failure exception | `GrammarCheckUnavailable` caught inside `_initialize_tool_if_needed()` | Let `SrcGenericError` wrap it via decorator | LanguageTool init failure is a predictable, non-fatal condition matching `SrcBaseWarning` semantics |
| `_tool.check()` failure exception | `GrammarCheckUnavailable` in try/except inside `check()` | Delegate to `@generic_error_handler` as `SrcGenericError` | Runtime Java errors during a check are also non-fatal and expected |

---

## Internal Flow Diagram

```
CheckGrammarUseCaseWiring.create_use_case()
  └── _get_grammar_check_port() -> LanguageToolAdapter(language="es")

[paragraphs: list[str]]
        |
        v
CheckGrammarUseCase.execute(paragraphs)
        |
        └──> LanguageToolAdapter.check(paragraphs)   [@generic_error_handler]
                 |
                 ├── _initialize_tool_if_needed()
                 │       if _tool is None:
                 │           try: _tool = LanguageTool("es")
                 │           except: raise GrammarCheckUnavailable()
                 |
                 ├── _build_sample_text(paragraphs)
                 │       "\n".join(paragraphs[:20])[:5000]
                 |
                 ├── try: _tool.check(text) -> raw_matches
                 │   except: raise GrammarCheckUnavailable()
                 |
                 ├── filter(ruleIssueType != "misspelling")
                 ├── [:10] limit
                 └── _map_to_dto(index, match) -> GrammarErrorDTO
                         match.message          -> message
                         match.context          -> context
                         match.offset           -> offset
                         match.errorLength      -> length
                         match.replacements[:3] -> replacements
                 |
                 v
        list[GrammarErrorDTO]
        |
        ├── _calculate_score(len(errors)) -> float
        │       == 0  -> 10.0
        │       <= 5  -> 8.5
        │       <= 15 -> 7.0
        │       else  -> 5.0
        |
        ├── _build_feedback(len(errors)) -> str
        │       == 0  -> "Sin errores gramaticales"
        │       <= 5  -> "Pocos errores gramaticales"
        │       <= 15 -> "Errores gramaticales moderados"
        │       else  -> "Muchos errores gramaticales"
        |
        v
GrammarCheckResultDTO(score=..., feedback=..., errors=[...])
```

---

## Class Signatures

### `GrammarCheckPort` — `src/domain/grammar/grammar_check_port.py`

```python
from abc import ABC, abstractmethod
from src.domain.dtos.grammar_error_dto import GrammarErrorDTO

class GrammarCheckPort(ABC):
    """Port for grammar checking services."""

    @abstractmethod
    def check(self, paragraphs: list[str]) -> list[GrammarErrorDTO]:
        """Return grammar issues found in the given paragraphs."""
```

### `GrammarErrorDTO` — `src/domain/dtos/grammar_error_dto.py`

```python
from dataclasses import dataclass
from src.domain.dtos.base_dto import BaseDTO

@dataclass(frozen=True)
class GrammarErrorDTO(BaseDTO):
    """A single grammar error found in the text."""

    number: int
    message: str
    context: str
    offset: int
    length: int
    replacements: list[str]
```

### `GrammarCheckResultDTO` — `src/domain/dtos/grammar_check_result_dto.py`

```python
from dataclasses import dataclass
from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.grammar_error_dto import GrammarErrorDTO

@dataclass(frozen=True)
class GrammarCheckResultDTO(BaseDTO):
    """The aggregate result of a grammar check analysis."""

    score: float
    feedback: str
    errors: list[GrammarErrorDTO]
```

### `grammar_errors.py` — `src/domain/exceptions/grammar_errors.py`

```python
from src.domain.exceptions.base_src_error import BaseSrcError, SrcBaseWarning

class GrammarError(BaseSrcError):
    """Base class for all grammar check exceptions."""

class GrammarCheckUnavailable(SrcBaseWarning):
    """Raised when the grammar checker backend is unavailable."""

    MESSAGE = "The grammar check service is unavailable."
```

### `FakeGrammarCheckPort` — `src/domain/tests/grammar/fake_grammar_check_port.py`

```python
from src.domain.dtos.grammar_error_dto import GrammarErrorDTO
from src.domain.grammar.grammar_check_port import GrammarCheckPort

class FakeGrammarCheckPort(GrammarCheckPort):
    """Test double for GrammarCheckPort with configurable errors list or exception."""

    def __init__(
        self,
        errors: list[GrammarErrorDTO] | None = None,
        error: Exception | None = None,
    ) -> None:
        self._errors = errors or []
        self._error = error

    def check(self, paragraphs: list[str]) -> list[GrammarErrorDTO]:
        """Return the configured errors or raise the configured exception."""
        if self._error is not None:
            raise self._error
        return self._errors
```

### `CheckGrammarUseCase` — `src/application/check_grammar_use_case.py`

```python
from src.domain.dtos.grammar_check_result_dto import GrammarCheckResultDTO
from src.domain.grammar.grammar_check_port import GrammarCheckPort

class CheckGrammarUseCase:
    """Orchestrates grammar checking and scoring for a set of text paragraphs."""

    # public
    def __init__(self, grammar_port: GrammarCheckPort) -> None:
        self._grammar_port = grammar_port

    def execute(self, paragraphs: list[str]) -> GrammarCheckResultDTO:
        """Check grammar and return a scored result DTO."""
        errors = self._grammar_port.check(paragraphs)
        error_count = len(errors)
        return GrammarCheckResultDTO(
            score=self._calculate_score(error_count),
            feedback=self._build_feedback(error_count),
            errors=errors,
        )

    # private
    def _build_feedback(self, error_count: int) -> str:
        """Return a human-readable message for the given error count."""
        if error_count == 0:
            return "Sin errores gramaticales"
        if error_count <= 5:
            return "Pocos errores gramaticales"
        if error_count <= 15:
            return "Errores gramaticales moderados"
        return "Muchos errores gramaticales"

    def _calculate_score(self, error_count: int) -> float:
        """Return a numeric score based on the number of grammar errors."""
        if error_count == 0:
            return 10.0
        if error_count <= 5:
            return 8.5
        if error_count <= 15:
            return 7.0
        return 5.0
```

### `LanguageToolAdapter` — `src/infrastructure/adapters/grammar/language_tool_adapter.py`

```python
import language_tool_python
from typing import Any

from src.domain.dtos.grammar_error_dto import GrammarErrorDTO
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.exceptions.grammar_errors import GrammarCheckUnavailable
from src.domain.grammar.grammar_check_port import GrammarCheckPort

_MAX_PARAGRAPHS = 20
_MAX_CHARS = 5000
_MAX_ERRORS = 10


class LanguageToolAdapter(GrammarCheckPort):
    """Grammar checking adapter backed by LanguageTool (requires Java runtime)."""

    # public
    def __init__(self, language: str = "es") -> None:
        self._language = language
        self._tool: language_tool_python.LanguageTool | None = None

    @generic_error_handler
    def check(self, paragraphs: list[str]) -> list[GrammarErrorDTO]:
        """Return grammar errors in the given paragraphs, excluding misspellings."""
        self._initialize_tool_if_needed()
        sample_text = self._build_sample_text(paragraphs)
        try:
            raw_matches = self._tool.check(sample_text)
        except Exception as exc:
            raise GrammarCheckUnavailable() from exc
        grammar_matches = [m for m in raw_matches if m.ruleIssueType != "misspelling"]
        return [
            self._map_to_dto(index, match)
            for index, match in enumerate(grammar_matches[:_MAX_ERRORS], start=1)
        ]

    # private (alphabetical)
    def _build_sample_text(self, paragraphs: list[str]) -> str:
        """Join the first MAX_PARAGRAPHS paragraphs and truncate to MAX_CHARS."""
        return "\n".join(paragraphs[:_MAX_PARAGRAPHS])[:_MAX_CHARS]

    def _initialize_tool_if_needed(self) -> None:
        """Initialize the LanguageTool Java instance on the first call."""
        if self._tool is None:
            try:
                self._tool = language_tool_python.LanguageTool(self._language)
            except Exception as exc:
                raise GrammarCheckUnavailable() from exc

    def _map_to_dto(self, number: int, match: Any) -> GrammarErrorDTO:
        """Map a LanguageTool match object to a GrammarErrorDTO."""
        return GrammarErrorDTO(
            number=number,
            message=match.message,
            context=match.context,
            offset=match.offset,
            length=match.errorLength,
            replacements=list(match.replacements[:3]),
        )
```

### `CheckGrammarUseCaseWiring` — `src/infrastructure/wirings/check_grammar_use_case_wiring.py`

```python
from src.application.check_grammar_use_case import CheckGrammarUseCase
from src.domain.grammar.grammar_check_port import GrammarCheckPort
from src.infrastructure.adapters.grammar.language_tool_adapter import LanguageToolAdapter

class CheckGrammarUseCaseWiring:
    """Factory for building a ready-to-use CheckGrammarUseCase."""

    # public
    def create_use_case(self) -> CheckGrammarUseCase:
        """Return a fully assembled CheckGrammarUseCase."""
        return CheckGrammarUseCase(grammar_port=self._get_grammar_check_port())

    # private
    def _get_grammar_check_port(self) -> GrammarCheckPort:
        """Return the concrete grammar check port implementation."""
        return LanguageToolAdapter()
```

---

## TDD Order (strict TDD mode active)

| Step | Test file (RED first) | Implementation file (GREEN) | Blocking |
|------|-----------------------|-----------------------------|----------|
| 1 | `test_grammar_error.py` + `test_grammar_check_unavailable.py` | `grammar_errors.py` | — |
| 2 | `test_grammar_error_dto.py` + `test_grammar_check_result_dto.py` | `grammar_error_dto.py` + `grammar_check_result_dto.py` | Step 1 (exception imports) |
| 3 | `test_grammar_check_port.py` | `grammar_check_port.py` + `fake_grammar_check_port.py` | Step 2 (DTO in port signature) |
| 4 | `test_check_grammar_use_case.py` | `check_grammar_use_case.py` | Step 3 (port + fake) |
| 5 | (adapter integration — `@skipIf` Java guard) within `test_check_grammar_use_case_wiring.py` | `language_tool_adapter.py` | Step 3 |
| 6 | `test_check_grammar_use_case_wiring.py` (type assertions) | `check_grammar_use_case_wiring.py` | Steps 4 + 5 |

**Note**: Steps 4 and 5 can proceed in parallel — both depend only on Step 3. The proposal's 18-file scope does not include a standalone `test_language_tool_adapter.py`; the adapter's type contract is validated through the wiring integration test. The tasks phase should evaluate adding a dedicated adapter unit test for lazy-init and error-handling behavior.

---

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/domain/grammar/__init__.py` | Create | Package marker for grammar entity folder |
| `src/domain/grammar/grammar_check_port.py` | Create | `GrammarCheckPort(ABC)` — abstract `check()` |
| `src/domain/dtos/grammar_error_dto.py` | Create | `GrammarErrorDTO` frozen dataclass |
| `src/domain/dtos/grammar_check_result_dto.py` | Create | `GrammarCheckResultDTO` frozen dataclass |
| `src/domain/exceptions/grammar_errors.py` | Create | `GrammarError` + `GrammarCheckUnavailable` |
| `src/domain/tests/grammar/__init__.py` | Create | Package marker |
| `src/domain/tests/grammar/fake_grammar_check_port.py` | Create | Configurable test double for `GrammarCheckPort` |
| `src/domain/tests/grammar/test_grammar_check_port.py` | Create | Unit tests for port ABC |
| `src/domain/tests/dtos/test_grammar_error_dto.py` | Create | Unit tests for `GrammarErrorDTO` |
| `src/domain/tests/dtos/test_grammar_check_result_dto.py` | Create | Unit tests for `GrammarCheckResultDTO` |
| `src/domain/tests/exceptions/test_grammar_error.py` | Create | Unit tests for `GrammarError` |
| `src/domain/tests/exceptions/test_grammar_check_unavailable.py` | Create | Unit tests for `GrammarCheckUnavailable` |
| `src/application/check_grammar_use_case.py` | Create | `CheckGrammarUseCase` with scoring logic |
| `src/application/tests/test_check_grammar_use_case.py` | Create | Unit tests via `FakeGrammarCheckPort` |
| `src/infrastructure/adapters/grammar/__init__.py` | Create | Package marker |
| `src/infrastructure/adapters/grammar/language_tool_adapter.py` | Create | `LanguageToolAdapter` with lazy init |
| `src/infrastructure/wirings/check_grammar_use_case_wiring.py` | Create | `CheckGrammarUseCaseWiring` |
| `src/infrastructure/tests/test_check_grammar_use_case_wiring.py` | Create | Wiring integration test (type assertions) |

---

## Implementation Notes

1. **Lazy init guard — non-obvious pattern**: `_initialize_tool_if_needed()` is a plain private method (no decorator). It catches `Exception` from `LanguageTool(self._language)` and raises `GrammarCheckUnavailable`. Because the `@generic_error_handler` on `check()` sees `GrammarCheckUnavailable` (a `SrcBaseWarning`), it re-raises it as-is. This is intentional — init failure is non-fatal.

2. **`_tool.check()` failure**: After successful lazy init, a runtime check failure (Java crash, network) is caught separately inside `check()` with another `try/except Exception` block and raised as `GrammarCheckUnavailable`. Do not rely solely on `@generic_error_handler` for this, because the decorator would wrap it as `SrcGenericError` instead.

3. **LanguageTool field mapping** (from legacy `gramatica_checker.py`):
   - `match.message` → description of the rule violation
   - `match.context` → surrounding text with error highlighted
   - `match.offset` → byte offset in the **full submitted text**
   - `match.errorLength` → character span of the error
   - `match.replacements[:3]` → top 3 suggested replacements
   - `match.ruleIssueType` → filter discriminator (`"misspelling"` excluded)

4. **`_map_to_dto` parameter type**: annotate as `Any` (from `typing`). The `language_tool_python` library does not export its internal match class as a public type; coupling to the private type would be fragile.

5. **Module-level constants**: `_MAX_PARAGRAPHS = 20`, `_MAX_CHARS = 5000`, `_MAX_ERRORS = 10` as module-level constants in `language_tool_adapter.py`. This avoids magic numbers and matches the pattern in other adapters with configuration values.

6. **Guard clause ordering in `_calculate_score` and `_build_feedback`**: use early returns from `== 0` → `<= 5` → `<= 15` → bare `return` (else). No `elif` chains. Matches CLAUDE.md guard clause rule.

7. **`GrammarCheckResultDTO.errors` field**: `list[GrammarErrorDTO]` in a `frozen=True` dataclass means the list reference is immutable — the list contents are mutable. Acceptable for an output DTO in this domain.

8. **Wiring integration test is safe without Java**: `LanguageToolAdapter()` construction does NOT start Java (lazy init). `isinstance` assertions pass immediately. The wiring test can run in any environment.

9. **`FakeGrammarCheckPort` default**: `errors=None` defaults to `[]` via `errors or []`. This simplifies the common case: `FakeGrammarCheckPort()` returns zero errors, representing a clean document.

---

## Consistency with Existing Adapters

| Aspect | `OllamaGeneratorAdapter` | `LanguageToolAdapter` |
|--------|--------------------------|----------------------|
| Module-level import | `import ollama` | `import language_tool_python` |
| Constructor params | `model_name: str, base_url: str` | `language: str = "es"` |
| Decorator | `@generic_error_handler` on `generate()` | `@generic_error_handler` on `check()` |
| Lazy init | No (stateless HTTP client) | Yes — Java process requires deferred startup |
| Private helpers | None | `_build_sample_text`, `_initialize_tool_if_needed`, `_map_to_dto` |
| Specific exception raised | `LanguageModelUnavailable` | `GrammarCheckUnavailable` |
| Exception parent | `SrcBaseWarning` | `SrcBaseWarning` |

Wiring pattern is identical to all 6 existing wirings: `create_use_case()` calls one `_get_<port>()` private method per dependency, returns the port interface type.

Fake pattern is identical to `FakeCharacterCountPort`, `FakeCitationExtractionPort`, etc.: `__init__(result | errors=None, error=None)`, configurable for success or failure paths.

---

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Domain exceptions | `GrammarError` is subclass of `BaseSrcError`; `GrammarCheckUnavailable` is subclass of `SrcBaseWarning`; `MESSAGE` constant value | `unittest.TestCase` |
| Domain DTOs | Frozen dataclass fields, equality, `as_dict()`, `from_dict()`, nested `errors` field | `unittest.TestCase` |
| Domain port | ABC cannot be instantiated directly; `FakeGrammarCheckPort` satisfies the contract | `unittest.TestCase` |
| Application use case | All 4 score thresholds (0, ≤5, ≤15, >15); all 4 feedback strings; delegation to port; returns `GrammarCheckResultDTO` | `unittest.TestCase` + `FakeGrammarCheckPort` |
| Infrastructure wiring | `create_use_case()` returns `CheckGrammarUseCase`; `._grammar_port` is `LanguageToolAdapter` | `unittest.TestCase` (integration, no Java needed) |

---

## Open Questions

- **Dedicated adapter test file**: The 18-file proposal scope does not include `test_language_tool_adapter.py`. The adapter's lazy-init behavior, misspelling filter, 10-error cap, and text sampling logic are not unit-tested in isolation. The tasks phase should decide: add a 19th file with `@skipIf` Java guard, or accept that the wiring test covers the integration boundary and adapter behavior is implicitly validated by end-to-end coexistence with the legacy layer.
