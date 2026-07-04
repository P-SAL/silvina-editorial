# Tasks: CheckGrammar — Slice 10 (Hexagonal Migration)

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~340–400 |
| 400-line budget risk | Medium |
| Chained PRs recommended | No |
| Suggested split | Single PR → `refactor/hexagonal-migration` |
| Delivery strategy | single-pr |
| Chain strategy | pending |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Medium

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | All 19 files (domain + application + infrastructure) | Single PR | Within single-pr budget; 19th file adds ~40 lines to original estimate |

---

## Phase 1: Domain Exceptions

Spec ref: *Requirement: Grammar Exceptions*

- [x] 1.1 [RED] Create `src/domain/tests/exceptions/test_grammar_error.py` — assert `GrammarError` is subclass of `BaseSrcError`; run `pytest` → RED
- [x] 1.2 [RED] Create `src/domain/tests/exceptions/test_grammar_check_unavailable.py` — assert `SrcBaseWarning` subclass + `MESSAGE == "The grammar check service is unavailable."`; run → RED
- [x] 1.3 [GREEN] Create `src/domain/exceptions/grammar_errors.py` — `GrammarError(BaseSrcError)` + `GrammarCheckUnavailable(SrcBaseWarning)` with `MESSAGE`; run 1.1 + 1.2 → GREEN

## Phase 2: Domain DTOs

Spec ref: *Requirement: Grammar DTOs*

- [x] 2.1 [RED] Create `src/domain/tests/dtos/test_grammar_error_dto.py` — assert construction with 6 fields + `FrozenInstanceError` on mutation; run → RED
- [x] 2.2 [RED] Create `src/domain/tests/dtos/test_grammar_check_result_dto.py` — assert construction with 3 fields + `FrozenInstanceError` on mutation; run → RED
- [x] 2.3 [GREEN] Create `src/domain/dtos/grammar_error_dto.py` (`GrammarErrorDTO(BaseDTO)`, `frozen=True`) + `grammar_check_result_dto.py` (`GrammarCheckResultDTO(BaseDTO)`, `frozen=True`); run 2.1 + 2.2 → GREEN

## Phase 3: Port ABC and Fake

Spec ref: *Requirement: GrammarCheckPort*

- [x] 3.1 [RED] Create `src/domain/tests/grammar/test_grammar_check_port.py` — assert `TypeError` on direct `GrammarCheckPort()` instantiation; assert `FakeGrammarCheckPort(errors=[...]).check([])` returns configured list; run → RED
- [x] 3.2 [GREEN] Create `src/domain/grammar/__init__.py` + `grammar_check_port.py` (`GrammarCheckPort(ABC)`, `@abstractmethod check()`) + `src/domain/tests/grammar/__init__.py` + `fake_grammar_check_port.py` (`FakeGrammarCheckPort` with `errors | None` and `error | None`); run 3.1 → GREEN

## Phase 4: Use Case *(depends on 1–3; parallel-capable with Phase 5)*

Spec ref: *Requirement: CheckGrammarUseCase Scoring*

- [x] 4.1 [RED] Create `src/application/tests/test_check_grammar_use_case.py` — 4 boundary scenarios using `FakeGrammarCheckPort`: 0 errors → `score=10.0`, 5 → `8.5`, 15 → `7.0`, 16 → `5.0`; run → RED
- [x] 4.2 [GREEN] Create `src/application/check_grammar_use_case.py` — `__init__(grammar_port)`, `execute()`, `_calculate_score()`, `_build_feedback()` with guard-clause early returns (`== 0` → `<= 5` → `<= 15` → bare return); run 4.1 → GREEN

## Phase 5: Adapter *(depends on 1–3; parallel-capable with Phase 4)*

Spec ref: *Requirement: LanguageToolAdapter* — 19th file (added in tasks phase per design open question)

- [x] 5.1 [RED] Create `src/infrastructure/tests/test_language_tool_adapter.py` — class-level `@skipIf(shutil.which('java') is None, "Java not available")`; 4 tests: `_tool is None` after construction, misspelling matches filtered, output capped at 10, `GrammarCheckUnavailable` on backend failure; run → RED (or SKIP if Java absent)
- [x] 5.2 [GREEN] Create `src/infrastructure/adapters/grammar/__init__.py` + `language_tool_adapter.py` — module constants `_MAX_PARAGRAPHS=20`, `_MAX_CHARS=5000`, `_MAX_ERRORS=10`; `__init__(language: str = "es")`; `@generic_error_handler` on `check()`; helpers `_initialize_tool_if_needed`, `_build_sample_text`, `_map_to_dto(number, match: Any)`; run 5.1 → GREEN/SKIP

## Phase 6: Wiring *(depends on 4 + 5)*

Spec ref: *Requirement: CheckGrammarUseCaseWiring*

- [x] 6.1 [RED] Create `src/infrastructure/tests/test_check_grammar_use_case_wiring.py` — assert `isinstance(result, CheckGrammarUseCase)` + `isinstance(result._grammar_port, LanguageToolAdapter)`; run → RED
- [x] 6.2 [GREEN] Create `src/infrastructure/wirings/check_grammar_use_case_wiring.py` — `create_use_case() -> CheckGrammarUseCase`, `_get_grammar_check_port() -> GrammarCheckPort` returning `LanguageToolAdapter()`; run 6.1 → GREEN

## Phase 7: Verification

- [x] 7.1 Run `python -m pytest src/ -q` — 454 passed, 6 skipped; no regressions in existing slices
- [x] 7.2 Verify no `src/` file imports from `business_logic/`; confirm `business_logic/gramatica_checker.py` is unmodified (coexistence invariant)
- [ ] 7.3 Commit all 19 files: `feat(check-grammar): add hexagonal grammar-checking path (Slice 10)` — PENDING USER REVIEW

---

## File Inventory (19 files — all new)

| # | File | Phase |
|---|------|-------|
| 1 | `src/domain/exceptions/grammar_errors.py` | 1 |
| 2 | `src/domain/tests/exceptions/test_grammar_error.py` | 1 |
| 3 | `src/domain/tests/exceptions/test_grammar_check_unavailable.py` | 1 |
| 4 | `src/domain/dtos/grammar_error_dto.py` | 2 |
| 5 | `src/domain/dtos/grammar_check_result_dto.py` | 2 |
| 6 | `src/domain/tests/dtos/test_grammar_error_dto.py` | 2 |
| 7 | `src/domain/tests/dtos/test_grammar_check_result_dto.py` | 2 |
| 8 | `src/domain/grammar/__init__.py` | 3 |
| 9 | `src/domain/grammar/grammar_check_port.py` | 3 |
| 10 | `src/domain/tests/grammar/__init__.py` | 3 |
| 11 | `src/domain/tests/grammar/fake_grammar_check_port.py` | 3 |
| 12 | `src/domain/tests/grammar/test_grammar_check_port.py` | 3 |
| 13 | `src/application/check_grammar_use_case.py` | 4 |
| 14 | `src/application/tests/test_check_grammar_use_case.py` | 4 |
| 15 | `src/infrastructure/adapters/grammar/__init__.py` | 5 |
| 16 | `src/infrastructure/adapters/grammar/language_tool_adapter.py` | 5 |
| 17 | `src/infrastructure/tests/test_language_tool_adapter.py` | 5 (new, 19th file) |
| 18 | `src/infrastructure/wirings/check_grammar_use_case_wiring.py` | 6 |
| 19 | `src/infrastructure/tests/test_check_grammar_use_case_wiring.py` | 6 |
