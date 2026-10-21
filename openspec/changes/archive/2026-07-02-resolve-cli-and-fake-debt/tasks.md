# Tasks: Resolve CLI and Fake Debt

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 50-70 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | single-pr |
| Chain strategy | size-exception |

Decision needed before apply: Yes
Chained PRs recommended: No
Chain strategy: size-exception
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Refactor fake LLM double and implement CLI exit code propagation | PR 1 | Base branch: main; includes all tests and verification |

## Phase 1: Refactoring the Quality Double

- [x] 1.1 Create new test double class [FakeLlmGeneratorAdapter](file:///E:/Python/silvina-editorial/src/domain/tests/quality/fake_llm_generator_adapter.py) in [fake_llm_generator_adapter.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/fake_llm_generator_adapter.py) inheriting from [LlmGeneratorPort](file:///E:/Python/silvina-editorial/src/domain/ports/llm_generator_port.py) and implementing `generate(self, prompt: str, options: dict | None = None)`.
- [x] 1.2 Update imports and instantiations of the LLM generator double in [test_quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_analyzer.py) to use the new `FakeLlmGeneratorAdapter`.
- [x] 1.3 Delete the outdated double file [fake_llm_generator_port.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/fake_llm_generator_port.py).
- [x] 1.4 Verify quality domain unit tests pass by running `.venv/Scripts/pytest src/domain/tests/quality/`.

## Phase 2: CLI Exit Code Propagation (Strict TDD)

- [x] 2.1 RED: Add failing test `test_exits_1_when_save_word_report_fails` in [test_main_cli_args.py](file:///E:/Python/silvina-editorial/tests/test_main_cli_args.py) asserting that the process exits with code `1` and prints `"Error: No se pudo guardar el reporte de Word (DOCX)."` to stdout/stderr while ensuring `save_json_report` is still called.
- [x] 2.2 GREEN: In [main.py](file:///E:/Python/silvina-editorial/main.py), capture return value of `save_word_report()`, run `save_json_report()`, and if the Word report failed to save, print `"Error: No se pudo guardar el reporte de Word (DOCX)."` and exit with `sys.exit(1)`.
- [x] 2.3 REFACTOR: Review modifications in [main.py](file:///E:/Python/silvina-editorial/main.py), clean up logic, ensure compliance with project coding standards, and verify exit code propagation pathways.

## Phase 3: Verification

- [x] 3.1 Run tests with `.venv/Scripts/pytest tests/test_main_cli_args.py` to verify exit code propagation passes.
- [x] 3.2 Run quality domain tests with `.venv/Scripts/pytest src/domain/tests/quality/` to verify refactoring didn't introduce regression.
- [x] 3.3 Run ruff linter and formatter validation using `ruff check .` and `ruff format --check .`.
