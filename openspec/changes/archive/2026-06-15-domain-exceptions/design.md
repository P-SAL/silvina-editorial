# Technical Design: Domain Exceptions (Slice 1)

> Slice 1 of the hexagonal migration (`docs/plan-migracion-hexagonal.md`).
> Normative guide: `.agent/skills/clean-architecture/SKILL.md`.
> Scope: PURE domain only — exception group files. No ports, adapters, use cases, wirings.

## 1. Architecture Approach

This slice fills five exception group files in the existing
`src/domain/exceptions/` skeleton. No structural changes are made.

The pattern for each file (SKILL §5):

```python
from src.domain.exceptions.base_src_error import SrcBaseNotFound, SrcBaseWarning


class DocumentNotFound(SrcBaseNotFound):
    """Raised when a document file cannot be located."""

    MESSAGE = "The document file could not be found."


class DocumentEmpty(SrcBaseWarning):
    """Raised when a document has no readable content."""

    MESSAGE = "The document has no readable content."
```

**Key structural decisions:**

1. **Multiple exceptions per file** — SKILL §4 requires one class per file for
   all domain artifacts, but §5 explicitly carves out `src/domain/exceptions/`
   as the one place where multiple related exception classes may coexist in a
   single file. Each file groups exceptions by domain area, matching the
   per-group naming convention in plan §7.

2. **`MESSAGE` class attribute** — all custom exceptions define a `MESSAGE: str`
   class-level attribute. `BaseSrcError.dict()` uses it as the error payload.
   Exceptions that do not need a message at `BaseSrcError` level still define
   `MESSAGE` for consistency and to satisfy the base contract.

3. **Base type selection** — two base types are used:
   - `SrcBaseNotFound` — for "entity/resource cannot be located" errors
     (`DocumentNotFound`). These are non-recoverable at the domain level: the
     resource simply isn't there.
   - `SrcBaseWarning` — for recoverable operational failures: the resource
     exists (or might exist) but a processing step failed. All other exceptions
     in this slice use `SrcBaseWarning`.

4. **No `__init__` override** — none of the exception classes in this slice
   take constructor arguments. The `BaseSrcError.__init__` (zero-arg, sets
   `was_error_logged`) is inherited as-is.

5. **Docstrings required** — SKILL §7 requires docstrings on all public classes.
   Each exception has a single-line docstring describing when it is raised.

## 2. File Layout

```
src/domain/exceptions/
├── base_src_error.py              (existing — not modified)
├── document_errors.py             (new — Slice 1)
├── citation_errors.py             (new — Slice 1)
├── classification_errors.py       (new — Slice 1)
├── quality_errors.py              (new — Slice 1)
├── language_model_errors.py       (new — Slice 1)
└── decorators/
    └── generic_error_handler.py   (existing — not modified)

src/domain/tests/exceptions/
├── test_base_src_error.py         (existing)
├── test_document_errors.py        (new — Slice 1)
├── test_citation_errors.py        (new — Slice 1)
├── test_classification_errors.py  (new — Slice 1)
├── test_quality_errors.py         (new — Slice 1)
├── test_language_model_errors.py  (new — Slice 1)
└── decorators/
    └── test_generic_error_handler.py  (existing)
```

## 3. Exception Classes per Group

### `document_errors.py`

| Class | Base | Rationale |
|---|---|---|
| `DocumentNotFound` | `SrcBaseNotFound` | File path does not exist; non-recoverable |
| `DocumentEmpty` | `SrcBaseWarning` | File found but has no readable content; recoverable (caller may skip or warn) |
| `DocumentUnreadable` | `SrcBaseWarning` | File found but parsing/reading failed; recoverable (caller may use fallback) |

`document_errors.py` is the only file in this slice with more than one base
type (`SrcBaseNotFound` for not-found, `SrcBaseWarning` for reading warnings).
Both base types are imported from `base_src_error.py` at the top of the file.

### `citation_errors.py`

| Class | Base | Rationale |
|---|---|---|
| `CitationParsingFailed` | `SrcBaseWarning` | Citation text found but could not be parsed to a structured type; caller proceeds with partial data |

### `classification_errors.py`

| Class | Base | Rationale |
|---|---|---|
| `ClassificationFailed` | `SrcBaseWarning` | LLM or rules engine could not classify the article; caller may default to `ArticleType.UNKNOWN` |

### `quality_errors.py`

| Class | Base | Rationale |
|---|---|---|
| `QualityAnalysisFailed` | `SrcBaseWarning` | LLM call for quality analysis failed; caller may skip or surface a degraded result |

### `language_model_errors.py`

| Class | Base | Rationale |
|---|---|---|
| `LanguageModelUnavailable` | `SrcBaseWarning` | Ollama or the configured LLM backend cannot be reached; caller decides on fallback |

> `LanguageModelUnavailable` is `SrcBaseWarning` (not a more severe type)
> because the adapter policy is to degrade gracefully — the use case decides
> whether to propagate or return a default. The generic error handler re-raises
> `SrcBaseWarning` as-is (SKILL §5 handler table).

## 4. Import Convention

Each group file imports only what it needs from `base_src_error.py`:

```python
# document_errors.py (needs both base types)
from src.domain.exceptions.base_src_error import SrcBaseNotFound, SrcBaseWarning

# citation_errors.py (single base type)
from src.domain.exceptions.base_src_error import SrcBaseWarning
```

No other imports. These files have zero external dependencies — they import
only from `src.domain.exceptions.base_src_error` (a sibling in the same
package). No circular imports are possible.

## 5. Test Strategy

**Framework:** `unittest.TestCase`, MANDATORY (SKILL §6).
**Runner:** `python -m pytest src/`.

Per-group test file covers:
- Each exception class `issubclass` of its direct base type
  (`SrcBaseNotFound` or `SrcBaseWarning`).
- Each exception is catchable as `BaseSrcError` (via `with self.assertRaises`).

```python
from unittest import TestCase

from src.domain.exceptions.base_src_error import BaseSrcError, SrcBaseWarning
from src.domain.exceptions.citation_errors import CitationParsingFailed


class TestCitationParsingFailed(TestCase):
    def test_is_subclass_of_src_base_warning(self):
        self.assertTrue(issubclass(CitationParsingFailed, SrcBaseWarning))

    def test_is_catchable_as_base_src_error(self):
        with self.assertRaises(BaseSrcError):
            raise CitationParsingFailed()
```

Tests are pure Python — no I/O, no external libraries. They run as part of the
full `python -m pytest src/` suite.

## 6. Coexistence Guarantee

- `base_src_error.py` and `generic_error_handler.py` are NOT modified.
- No legacy file outside `src/` is read, modified, or deleted.
- The legacy test suite under `tests/` continues to pass.
- The new exception classes are defined but not yet raised by any production
  caller — wiring them in is the responsibility of each later slice.

## 7. ADR-Style Decisions

**ADR-1: Multiple exceptions per group file is intentional (not a SKILL violation).**
- SKILL §4 states one class per file for all domain artifacts.
- SKILL §5 explicitly overrides this for `src/domain/exceptions/`: "files inside
  `src/domain/exceptions/` MAY contain multiple related exception classes."
- Decision: each group file holds all related exceptions for that domain area.
  This avoids a proliferation of single-line exception files and keeps the
  exception grouping legible.

**ADR-2: `SrcBaseWarning` for all operational failures.**
- `ClassificationFailed`, `QualityAnalysisFailed`, `LanguageModelUnavailable`,
  `CitationParsingFailed`, `DocumentEmpty`, `DocumentUnreadable` are all
  `SrcBaseWarning` because they represent degraded-but-recoverable outcomes:
  the use case can decide to surface a partial result or a default.
- `SrcBaseNotFound` is reserved for `DocumentNotFound` (file literally absent).
- Rejected: using `SrcGenericError` for LLM failures — that type wraps
  unexpected infrastructure exceptions; operational LLM unavailability is a
  known, typed failure.

**ADR-3: `MESSAGE` attribute on every exception class.**
- `BaseSrcError.dict()` returns `self.MESSAGE` when set, or a default message
  including the class name. Defining `MESSAGE` on every exception class keeps
  the `.dict()` output predictable and human-readable at every layer.
- Rejected: relying on the default message — the default names the class, which
  leaks implementation names to entry points.
