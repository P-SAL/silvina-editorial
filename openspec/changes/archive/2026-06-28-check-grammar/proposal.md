# Proposal: check-grammar (Slice 10 — Hexagonal Migration)

**Change name**: check-grammar
**Slice**: 10 of N (incremental hexagonal migration)
**Date**: 2026-06-28
**Status**: proposed

---

## 1. Intent

### Problem

`business_logic/gramatica_checker.py` exposes a module-level function
`check_gramatica(paragraphs)` that couples grammar checking, scoring logic, text
sampling, and error filtering into a single infrastructure routine. It creates a
fresh `LanguageTool('es')` instance on every call (no reuse), embeds the scoring
formula (thresholds: 0, 5, 15 errors) inside the infrastructure layer, and uses a
broad `except` clause that swallows failures silently. The function is a
module-level entity, not a class — violating the OOP rule from the architecture
convention.

No hexagonal path exists for grammar checking. Any caller in `src/` that needs
grammar results must reach into `business_logic/`, breaking the dependency rule and
making the feature untestable through port substitution.

### Why now

Slice 10 follows the established rhythm: domain ports are stable and understood,
the adapter complexity (lazy Java init, text sampling) is bounded, and the scoring
formula is compact enough to live in the use case without a dedicated domain
service. Delivering this slice completes another column of the hexagonal migration
plan and makes grammar checking testable in isolation via a fake adapter.

### Success criteria

- A caller in `src/` can invoke `CheckGrammarUseCase.execute(paragraphs)` and
  receive a `GrammarCheckResultDTO(score, feedback, errors)` without importing
  anything from `business_logic/`.
- `LanguageTool` is not instantiated during import or test collection — only on the
  first real call (lazy init).
- All new code is covered by `unittest.TestCase` tests following the established
  port/fake-double/integration pattern.
- The legacy `business_logic/gramatica_checker.py` remains untouched and the system
  continues to work end-to-end (coexistence invariant).

---

## 2. Scope

### In scope — Files to create (18 new files, 0 modified)

**Domain — port**
- `src/domain/grammar/__init__.py`
- `src/domain/grammar/grammar_check_port.py`
  `GrammarCheckPort(ABC)` — single abstract method `check(paragraphs: list[str]) -> list[GrammarErrorDTO]`

**Domain — DTOs**
- `src/domain/dtos/grammar_error_dto.py`
  `GrammarErrorDTO` — frozen dataclass: `number: int`, `message: str`, `context: str`,
  `offset: int`, `length: int`, `replacements: list[str]`
- `src/domain/dtos/grammar_check_result_dto.py`
  `GrammarCheckResultDTO` — frozen dataclass: `score: float`, `feedback: str`,
  `errors: list[GrammarErrorDTO]`

**Domain — exceptions**
- `src/domain/exceptions/grammar_errors.py`
  `GrammarError(BaseSrcError)` base class + `GrammarCheckUnavailable(SrcBaseWarning)`
  with `MESSAGE = "The grammar check service is unavailable."`

**Domain — tests**
- `src/domain/tests/grammar/__init__.py`
- `src/domain/tests/grammar/fake_grammar_check_port.py`
  `FakeGrammarCheckPort(GrammarCheckPort)` — returns a configurable list of
  `GrammarErrorDTO`s; no Java or external calls
- `src/domain/tests/grammar/test_grammar_check_port.py`
- `src/domain/tests/dtos/test_grammar_error_dto.py`
- `src/domain/tests/dtos/test_grammar_check_result_dto.py`
- `src/domain/tests/exceptions/test_grammar_error.py`
- `src/domain/tests/exceptions/test_grammar_check_unavailable.py`

**Application — use case**
- `src/application/check_grammar_use_case.py`
  `CheckGrammarUseCase` — receives `GrammarCheckPort` via constructor injection;
  `execute(paragraphs: list[str]) -> GrammarCheckResultDTO`; scoring and feedback
  are computed here (not in the adapter):
  - 0 errors → 10.0 / "Sin errores gramaticales"
  - ≤ 5 errors → 8.5 / "Pocos errores gramaticales"
  - ≤ 15 errors → 7.0 / "Errores gramaticales moderados"
  - > 15 errors → 5.0 / "Muchos errores gramaticales"

**Application — tests**
- `src/application/tests/test_check_grammar_use_case.py`
  Unit tests using `FakeGrammarCheckPort` (imported from `src/domain/tests/grammar/`)

**Infrastructure — adapter**
- `src/infrastructure/adapters/grammar/__init__.py`
- `src/infrastructure/adapters/grammar/language_tool_adapter.py`
  `LanguageToolAdapter(GrammarCheckPort)` — module-level `import language_tool_python`;
  lazy init: `self._tool: language_tool_python.LanguageTool | None = None`,
  initialized on first `check()` call; samples first 20 paragraphs up to 5000
  chars; filters `rule_issue_type == 'misspelling'`; limits to first 10 errors;
  raises `GrammarCheckUnavailable` on failure; `@generic_error_handler` on `check()`

**Infrastructure — wiring**
- `src/infrastructure/wirings/check_grammar_use_case_wiring.py`
  `CheckGrammarUseCaseWiring.create_use_case() -> CheckGrammarUseCase`
  private `_get_grammar_check_port() -> GrammarCheckPort` returning `LanguageToolAdapter`

**Infrastructure — tests**
- `src/infrastructure/tests/test_check_grammar_use_case_wiring.py`
  Integration test: instantiates wiring, calls `create_use_case()`, asserts
  `isinstance(use_case, CheckGrammarUseCase)` and private port attribute types

### Out of scope

- Any modification to `business_logic/gramatica_checker.py` — kept alive for coexistence
- Spelling check errors (`rule_issue_type == 'misspelling'`) — filtered out in
  the adapter, consistent with legacy behavior; no separate port or feature
- LanguageTool language parameterization beyond `'es'` — deferred; constructor
  parameter `language: str = 'es'` may be added for testability at spec time
- Domain service `GrammarScoreCalculator` — scoring thresholds are compact enough
  in the use case; no dedicated service per migration plan §6
- Wiring into the application entry points (Gradio UI) — future integration step
- `__init__.py` changes in existing packages — follow project convention;
  verify at spec time

---

## 3. Approach

### Architecture

```
[paragraphs: list[str]]
         |
         v
LanguageToolAdapter ──implements──> GrammarCheckPort (ABC, domain/grammar/)
         |  lazy-init LanguageTool('es'), sample, filter, map -> list[GrammarErrorDTO]
         v
CheckGrammarUseCase.execute(paragraphs)
         └── _grammar_port.check(paragraphs)  -> list[GrammarErrorDTO]
         └── _calculate_score(error_count)    -> float
         └── _build_feedback(error_count)     -> str
         └── returns GrammarCheckResultDTO(score, feedback, errors)
         |
         v
CheckGrammarUseCaseWiring.create_use_case()
         └── _get_grammar_check_port() -> LanguageToolAdapter
```

**Port location**: `src/domain/grammar/` — entity-scoped, consistent with Slices 5–9.
Older `domain/ports/` folder (Ollama pattern, Slice < 5) is NOT used.

**Scoring in use case**: scoring thresholds are domain/application logic, not
infrastructure. The adapter returns raw `list[GrammarErrorDTO]`; the use case
calculates `score` and `feedback`. This is Approach A from the exploration.

**Lazy LanguageTool init**: `import language_tool_python` at module level
(no local imports per SKILL §3). The `LanguageTool('es')` instance is created
inside `check()` on first call only (`if self._tool is None`). This prevents Java
from starting during import or test collection.

**Error handling**: `@generic_error_handler` on `LanguageToolAdapter.check()` only.
ABC methods carry no decorator. `GrammarCheckUnavailable` propagates as a
`SrcBaseWarning` (non-fatal), consistent with the exception hierarchy.

**Wiring method**: `create_use_case()` — actual project convention (all existing
wirings use this; SKILL §8 text says `get_<name>()` but the codebase overrides it).

### TDD order (strict TDD mode active)

1. Exception classes + tests (no dependencies)
2. DTOs + tests (`GrammarErrorDTO`, `GrammarCheckResultDTO`)
3. Port ABC + test + fake double (`FakeGrammarCheckPort`)
4. Use case unit tests (with fake) → use case implementation (RED → GREEN → REFACTOR)
5. Adapter integration test (real LanguageTool, skipped if Java unavailable) → adapter
6. Wiring test → wiring implementation

---

## 4. Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Java not present in CI/test env → LanguageTool fails on init | Medium | Lazy init defers Java startup to first `check()` call; integration tests may use `@unittest.skipIf` guard; `GrammarCheckUnavailable` propagates cleanly |
| LanguageTool startup latency (5–10 s on cold start) slows integration tests | High | `FakeGrammarCheckPort` for unit tests; integration test clearly labelled and separated; real adapter test can be marked slow |
| Scoring logic accidentally moves into adapter (violates layer boundary) | Low | Code review + verify phase checks that `LanguageToolAdapter.check()` returns only `list[GrammarErrorDTO]` |
| `grammar_errors.py` not listed in plan §7 exception catalog | Low | Self-contained new group; no cross-slice impact; note discrepancy in spec |
| Lazy init thread-safety under concurrent Gradio requests | Low | Wiring creates a new adapter per invocation (no singleton); low concurrency in editorial tool |
| `language_tool_python` version incompatibility | Low | Pin version in `requirements.txt`; verify against existing legacy usage |

---

## 5. PR Shape

Single PR targeting `refactor/hexagonal-migration`.
- 18 new files, 0 modified files
- Estimated changed lines: ~280–350 (well under 400-line budget)
- No chained PRs required

---

## 6. Definition of Done

Per `docs/plan-migracion-hexagonal.md` §8 — a slice is done when all of the
following are checked:

- [ ] `GrammarCheckPort(ABC)` defined in `src/domain/grammar/` with `check()` returning `list[GrammarErrorDTO]`
- [ ] `GrammarErrorDTO` and `GrammarCheckResultDTO` (frozen dataclasses) with `unittest.TestCase` coverage
- [ ] `GrammarError` and `GrammarCheckUnavailable` in `src/domain/exceptions/grammar_errors.py` with tests
- [ ] `CheckGrammarUseCase.execute()` implements scoring and feedback (Approach A thresholds) — tests use `FakeGrammarCheckPort`
- [ ] `LanguageToolAdapter` has lazy LanguageTool init; module-level import; `@generic_error_handler` on `check()`
- [ ] `CheckGrammarUseCaseWiring.create_use_case()` returns a fully wired instance; wiring integration test passes
- [ ] No imports from `business_logic/` in any `src/` file
- [ ] Domain imports nothing from infrastructure or application (hexagonal invariant)
- [ ] No local imports; no wildcard imports; no `print()` statements
- [ ] One class per file (exception: `grammar_errors.py` may contain multiple exception classes)
- [ ] `python -m pytest src/ -q` passes with all tests green
- [ ] `business_logic/gramatica_checker.py` untouched; system works end-to-end

---

## 7. Dependencies

- `language_tool_python` — already installed (used by legacy layer)
- Java runtime — required by LanguageTool; assumed present in production; optional for unit tests
- `src/domain/exceptions/base_src_error.py` — `BaseSrcError`, `SrcBaseWarning` (exist)
- `src/domain/exceptions/decorators/generic_error_handler.py` — `@generic_error_handler` (exists)
- `src/domain/dtos/base_dto.py` — `BaseDTO` base class (exists)
