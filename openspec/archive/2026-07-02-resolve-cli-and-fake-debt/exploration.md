## Exploration: resolve-cli-and-fake-debt

### Current State
1. **CLI Exit Code Propagation (TECHNICAL_DEBT.md Item 2)**:
   - In [main.py](file:///E:/Python/silvina-editorial/main.py#L267), `silvina.save_word_report(results, str(word_report_path))` is called but its return value (a boolean) is ignored.
   - `SilvinaEditorialAssistant.save_word_report` catches exceptions internally and returns `False` on failure. However, because the return value is ignored in `main()`, the CLI continues execution, prints "ANÁLISIS COMPLETADO", and exits with code `0`.
   - The CLI exit code behavior is tested in [tests/test_main_cli_args.py](file:///E:/Python/silvina-editorial/tests/test_main_cli_args.py) via the `TestMainExitCodes` class, but there is currently no test covering a `save_word_report` failure.

2. **FakeLlmGeneratorPort Naming (TECHNICAL_DEBT.md Item 5)**:
   - The test double [fake_llm_generator_port.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/fake_llm_generator_port.py) defines the `FakeLlmGeneratorPort` class.
   - This class does not inherit from `LlmGeneratorPort`, using pure duck-typing instead. According to hexagonal terminology, a test fake is an Adapter, not a Port.
   - A similar issue was already resolved in the classification tests by introducing `FakeLlmGeneratorAdapter(LlmGeneratorPort)` under [fake_llm_generator_adapter.py](file:///E:/Python/silvina-editorial/src/domain/tests/classification/fake_llm_generator_adapter.py), which also conforms to the `LlmGeneratorPort` signature including the `options` parameter.
   - `FakeLlmGeneratorPort` is imported and instantiated in [test_quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_analyzer.py).

### Affected Areas
- [main.py](file:///E:/Python/silvina-editorial/main.py) — Must check the returned boolean from `save_word_report()` and exit with code `1` if it is `False` (saving the JSON report first to avoid data loss).
- [tests/test_main_cli_args.py](file:///E:/Python/silvina-editorial/tests/test_main_cli_args.py) — Needs a unit test in `TestMainExitCodes` asserting that `main()` exits with status `1` when `save_word_report` returns `False`.
- [src/domain/tests/quality/fake_llm_generator_port.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/fake_llm_generator_port.py) — Should be renamed to `fake_llm_generator_adapter.py`. The class must be renamed to `FakeLlmGeneratorAdapter`, inherit from `LlmGeneratorPort`, and support the `options` parameter in its signature.
- [src/domain/tests/quality/test_quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_analyzer.py) — Needs to update its import statement and usages of the fake class.

### Approaches
1. **Handle save_word_report() failure in main()**
   - **Description**: Assign the result of `save_word_report(...)` to a variable, proceed with saving the JSON report (so the user doesn't lose the JSON data), and then call `sys.exit(1)` if saving the Word report failed.
   - Pros: Simple, ensures exit code correctness without interrupting other tasks, keeps JSON backup.
   - Cons: None.
   - Effort: Low.

2. **Refactor FakeLlmGeneratorPort to FakeLlmGeneratorAdapter**
   - **Description**: Rename the file and class, inherit from `LlmGeneratorPort`, implement the signature with `options: dict | None = None`, and update usages in `test_quality_analyzer.py`.
   - Pros: Restores proper hexagonal naming terminology, achieves alignment with the classification domain, prevents future signature drift.
   - Cons: None.
   - Effort: Low.

### Recommendation
Implement both approaches. For the CLI, capture the success of `save_word_report` and check it after `save_json_report` runs, then exit with code `1` if it is `False`. For the fake double, perform the file rename, rename the class, inherit from the port, and align the method signatures and imports.

### Risks
- **JSON Save Bypass**: Exiting immediately on Word report failure could bypass saving the JSON report. The recommended approach is to run `save_json_report` first, then call `sys.exit(1)` if `save_word_report` failed.

### Ready for Proposal
Yes. The changes are very clear and can move directly to the proposal phase. The orchestrator should prompt the user to proceed.
