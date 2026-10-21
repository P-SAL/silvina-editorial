# Proposal: Resolve CLI and Fake Debt

## Intent

Resolve [TECHNICAL_DEBT.md](file:///E:/Python/silvina-editorial/TECHNICAL_DEBT.md) Items 2 and 5:
- CLI must exit with status `1` and propagate errors when saving the Word (DOCX) report fails, ensuring CLI reliability.
- Refactor the quality double `FakeLlmGeneratorPort` to follow hexagonal naming conventions (`FakeLlmGeneratorAdapter`), inherit from its port, and match the target signature.

## Scope

### In Scope
- **CLI failure handling**: Handle `save_word_report()` failure in [main.py](file:///E:/Python/silvina-editorial/main.py) by printing `"Error: No se pudo guardar el reporte de Word (DOCX)."` and calling `sys.exit(1)` (after saving the JSON report to prevent data loss).
- **Refactor quality fake double**:
  - Rename `src/domain/tests/quality/fake_llm_generator_port.py` to `fake_llm_generator_adapter.py`.
  - Rename the class `FakeLlmGeneratorPort` to `FakeLlmGeneratorAdapter`, make it inherit from `LlmGeneratorPort`, and support the `options` parameter.
  - Delete the old file `src/domain/tests/quality/fake_llm_generator_port.py`.
- **Update test imports**: Update imports/references in `src/domain/tests/quality/test_quality_analyzer.py` to use `FakeLlmGeneratorAdapter`.
- **Unit testing**: Add a TDD unit test in `tests/test_main_cli_args.py` asserting exit code `1` when `save_word_report()` fails.

### Out of Scope
- Changing existing domain rules or validation criteria.
- Modifying behavior or location of classification fake doubles.

## Capabilities

### New Capabilities
None

### Modified Capabilities
None

## Approach

1. **Word Save Check in CLI**: Update [main.py](file:///E:/Python/silvina-editorial/main.py) to capture the boolean return of `save_word_report()`. Run `save_json_report()` first, then check the boolean; if `False`, print the error message and exit with `1`.
2. **Refactor Double**: Apply the file rename/class rename under quality tests. Update the signature and inherit from `LlmGeneratorPort`. Update imports in `test_quality_analyzer.py`.
3. **Tests**: Use `unittest.mock.patch` to mock `save_word_report()` return value to `False` and verify `sys.exit(1)` in `TestMainExitCodes`.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| [main.py](file:///E:/Python/silvina-editorial/main.py) | Modified | Check `save_word_report` status and exit with 1 on failure. |
| `src/domain/tests/quality/fake_llm_generator_port.py` | Removed | Delete old fake double. |
| `src/domain/tests/quality/fake_llm_generator_adapter.py` | New | Corrected hexagonal adapter name, inheriting from port. |
| `src/domain/tests/quality/test_quality_analyzer.py` | Modified | Import and use `FakeLlmGeneratorAdapter`. |
| [tests/test_main_cli_args.py](file:///E:/Python/silvina-editorial/tests/test_main_cli_args.py) | Modified | Add TDD test for CLI exit code on Word save failure. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| JSON report not saved when Word report fails | Low | Execute JSON report save before executing the failure exit. |
| Broken test execution due to missing import | Low | Verify all quality tests run successfully before finalizing. |

## Rollback Plan

Revert all changes using git:
```bash
git checkout main.py src/domain/tests/quality/test_quality_analyzer.py tests/test_main_cli_args.py
git checkout HEAD -- src/domain/tests/quality/fake_llm_generator_port.py
rm src/domain/tests/quality/fake_llm_generator_adapter.py
```

## Dependencies

None.

## Success Criteria

- [ ] `pytest tests/test_main_cli_args.py` runs successfully, including the new failure exit code test.
- [ ] Quality unit tests run successfully using the new `FakeLlmGeneratorAdapter`.
- [ ] No regression in report generation under normal operation.
