# CheckGrammar Specification (Slice 10 — Hexagonal Migration)

## Purpose

Introduce a complete hexagonal grammar-checking path under `src/` that requires no imports from `business_logic/`. The legacy `business_logic/gramatica_checker.py` MUST remain unmodified (coexistence invariant).

## Requirements

### Requirement: Grammar DTOs

The system MUST provide two frozen dataclasses:
- `GrammarErrorDTO(frozen=True)` — fields: `number: int`, `message: str`, `context: str`, `offset: int`, `length: int`, `replacements: list[str]`
- `GrammarCheckResultDTO(frozen=True)` — fields: `score: float`, `feedback: str`, `errors: list[GrammarErrorDTO]`

Both MUST inherit from `BaseDTO` (or `object` if project convention allows). Field mutation MUST raise `FrozenInstanceError`.

#### Scenario: DTO construction and immutability

- GIVEN valid field values for each dataclass
- WHEN constructed and then any field is reassigned
- THEN construction succeeds AND reassignment raises `FrozenInstanceError`

### Requirement: GrammarCheckPort

`GrammarCheckPort(ABC)` MUST declare a single `@abstractmethod`: `check(self, paragraphs: list[str]) -> list[GrammarErrorDTO]`. The class MUST NOT be directly instantiable.

#### Scenario: Abstract class cannot be instantiated

- GIVEN `GrammarCheckPort` with no concrete subclass
- WHEN instantiation is attempted
- THEN `TypeError` is raised

#### Scenario: FakeGrammarCheckPort satisfies the contract

- GIVEN `FakeGrammarCheckPort(GrammarCheckPort)` implementing `check()` with a configurable error list
- WHEN instantiated and `check([])` is called
- THEN the configured list of `GrammarErrorDTO` instances is returned

### Requirement: Grammar Exceptions

`src/domain/exceptions/grammar_errors.py` MUST define:
- `GrammarError(BaseSrcError)` — base class for all grammar exceptions
- `GrammarCheckUnavailable(SrcBaseWarning)` — with `MESSAGE = "The grammar check service is unavailable."`

#### Scenario: Correct inheritance chain

- GIVEN `GrammarError` and `GrammarCheckUnavailable`
- WHEN `issubclass` checks execute
- THEN `GrammarError` is a `BaseSrcError` AND `GrammarCheckUnavailable` is a `SrcBaseWarning` AND `GrammarCheckUnavailable.MESSAGE` equals `"The grammar check service is unavailable."`

### Requirement: CheckGrammarUseCase Scoring

`CheckGrammarUseCase.__init__` MUST accept `grammar_port: GrammarCheckPort`. `execute(paragraphs: list[str]) -> GrammarCheckResultDTO` MUST delegate error detection to the port and compute score/feedback by error count:

| Error count | Score | Feedback |
|-------------|-------|----------|
| 0 | 10.0 | "Sin errores gramaticales" |
| 1–5 | 8.5 | "Pocos errores gramaticales" |
| 6–15 | 7.0 | "Errores gramaticales moderados" |
| >15 | 5.0 | "Muchos errores gramaticales" |

#### Scenario: Zero errors — perfect score

- GIVEN `FakeGrammarCheckPort` returning `[]`
- WHEN `execute(paragraphs)` is called
- THEN result is `GrammarCheckResultDTO(score=10.0, feedback="Sin errores gramaticales", errors=[])`

#### Scenario: Five errors — boundary 8.5

- GIVEN `FakeGrammarCheckPort` returning 5 `GrammarErrorDTO` instances
- WHEN `execute(paragraphs)` is called
- THEN `result.score == 8.5` AND `result.feedback == "Pocos errores gramaticales"`

#### Scenario: Fifteen errors — boundary 7.0

- GIVEN `FakeGrammarCheckPort` returning 15 `GrammarErrorDTO` instances
- WHEN `execute(paragraphs)` is called
- THEN `result.score == 7.0` AND `result.feedback == "Errores gramaticales moderados"`

#### Scenario: Sixteen errors — threshold 5.0

- GIVEN `FakeGrammarCheckPort` returning 16 `GrammarErrorDTO` instances
- WHEN `execute(paragraphs)` is called
- THEN `result.score == 5.0` AND `result.feedback == "Muchos errores gramaticales"`

### Requirement: LanguageToolAdapter

`LanguageToolAdapter(GrammarCheckPort)` MUST:
- Import `language_tool_python` at module level
- Store `self._tool: language_tool_python.LanguageTool | None = None`; initialize `LanguageTool('es')` inside `check()` on first call only
- Sample first 20 paragraphs, truncated to 5000 chars total before passing to LanguageTool
- Filter out matches where `rule_issue_type == 'misspelling'`
- Return at most the first 10 errors as `list[GrammarErrorDTO]`
- Do NOT decorate `check()` with `@generic_error_handler`; error wrapping is handled at the use case layer

#### Scenario: Lazy init — no Java on import

- GIVEN the `language_tool_adapter` module is imported
- WHEN no `check()` call has been made
- THEN `self._tool` is `None` on any adapter instance

#### Scenario: Misspelling results are filtered

- GIVEN LanguageTool returns 3 grammar errors and 2 misspellings
- WHEN `check(paragraphs)` processes the results
- THEN exactly 3 `GrammarErrorDTO` instances are returned

#### Scenario: Output capped at 10 errors

- GIVEN LanguageTool returns 12 grammar errors (no misspellings)
- WHEN `check(paragraphs)` processes the results
- THEN exactly 10 `GrammarErrorDTO` instances are returned

#### Scenario: Backend failure propagates GrammarCheckUnavailable

- GIVEN `LanguageTool('es')` raises any exception during init or check
- WHEN `check(paragraphs)` is called
- THEN `GrammarCheckUnavailable` is raised (explicitly by the adapter, not via decorator)

### Requirement: CheckGrammarUseCaseWiring

`CheckGrammarUseCaseWiring` MUST expose `create_use_case() -> CheckGrammarUseCase` and a private `_get_grammar_check_port() -> GrammarCheckPort` returning a `LanguageToolAdapter`. No business logic in the wiring class.

#### Scenario: Wiring produces correctly typed instance

- GIVEN `CheckGrammarUseCaseWiring()` is instantiated
- WHEN `create_use_case()` is called
- THEN `isinstance(result, CheckGrammarUseCase)` is `True` AND `result._grammar_port` is a `LanguageToolAdapter`

## File Inventory

| File | Artifact |
|------|----------|
| `src/domain/grammar/__init__.py` | package |
| `src/domain/grammar/grammar_check_port.py` | `GrammarCheckPort` |
| `src/domain/dtos/grammar_error_dto.py` | `GrammarErrorDTO` |
| `src/domain/dtos/grammar_check_result_dto.py` | `GrammarCheckResultDTO` |
| `src/domain/exceptions/grammar_errors.py` | `GrammarError`, `GrammarCheckUnavailable` |
| `src/domain/tests/grammar/__init__.py` | package |
| `src/domain/tests/grammar/fake_grammar_check_port.py` | `FakeGrammarCheckPort` |
| `src/domain/tests/grammar/test_grammar_check_port.py` | port tests |
| `src/domain/tests/dtos/test_grammar_error_dto.py` | DTO tests |
| `src/domain/tests/dtos/test_grammar_check_result_dto.py` | DTO tests |
| `src/domain/tests/exceptions/test_grammar_error.py` | exception tests |
| `src/domain/tests/exceptions/test_grammar_check_unavailable.py` | exception tests |
| `src/application/check_grammar_use_case.py` | `CheckGrammarUseCase` |
| `src/application/tests/test_check_grammar_use_case.py` | use case unit tests |
| `src/infrastructure/adapters/grammar/__init__.py` | package |
| `src/infrastructure/adapters/grammar/language_tool_adapter.py` | `LanguageToolAdapter` |
| `src/infrastructure/wirings/check_grammar_use_case_wiring.py` | `CheckGrammarUseCaseWiring` |
| `src/infrastructure/tests/test_check_grammar_use_case_wiring.py` | wiring integration test |

## Invariants

1. No `src/` file MAY import from `business_logic/`.
2. `src/domain/` MUST NOT import from `src/application/` or `src/infrastructure/`.
3. `src/application/` MUST NOT import from `src/infrastructure/`.
4. `business_logic/gramatica_checker.py` MUST remain unmodified.
5. All imports MUST be at module level (no local or wildcard imports).
6. One class per file — `grammar_errors.py` is the sole exception (multiple exception classes allowed per SKILL §4).
7. No `print()` statements in production code.
8. All tests MUST use `unittest.TestCase`. No pytest-specific constructs.
