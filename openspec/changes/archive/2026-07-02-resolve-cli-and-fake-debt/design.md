# Design: Resolve CLI and Fake Debt

## Technical Approach

We will resolve technical debt items 2 and 5 by ensuring proper CLI exit code propagation on Word report save failure, and refactoring the quality test fake to conform to hexagonal naming/architectural rules.

1. **CLI Failure Handling**: Capture the return value of `save_word_report()`. If `False`, we print `"Error: No se pudo guardar el reporte de Word (DOCX)."` and exit with `sys.exit(1)` *after* executing `save_json_report()` to prevent data loss.
2. **Hexagonal Test Double Refactoring**: Rename the fake LLM double file to `fake_llm_generator_adapter.py`, rename class to `FakeLlmGeneratorAdapter`, make it inherit from `LlmGeneratorPort`, and update its signature to match `generate(self, prompt: str, options: dict | None = None)`. Update all imports and usages in `test_quality_analyzer.py`.

## Architecture Decisions

| Option | Tradeoff | Decision |
|--------|----------|----------|
| Exiting immediately on Word save failure vs. saving JSON report first | Exiting immediately is simpler but causes potential data loss of the JSON report. Saving JSON first preserves the data. | **Save JSON report first**, then check Word save status and exit with code `1` if `save_word_report` failed. |
| Inheritance vs. Duck Typing for Fake | Duck typing requires no imports but allows signature/contract drift. Inheritance guarantees compatibility with `LlmGeneratorPort`. | **Inherit from `LlmGeneratorPort`** and align name to `FakeLlmGeneratorAdapter` per hexagonal naming conventions. |
| Mocking `SilvinaEditorialAssistant` vs. using actual use case | Mocking is faster and cleaner for CLI testing. actual class testing is covered in other layers. | **Mock the assistant class** and assert its mock method calls and the console output using `unittest.mock`. |

## Data Flow

```
main() (CLI entry point)
  │
  ├──► silvina.analyze_document() ──► results
  │
  ├──► word_saved = silvina.save_word_report(results, word_path)
  ├──► silvina.save_json_report(results, json_path)
  │
  └──► If not word_saved:
         ├──► Print error message
         └──► sys.exit(1)
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `main.py` | Modify | Capture return value of `save_word_report()`. Save JSON first, then conditional exit if saving Word report failed. |
| `src/domain/tests/quality/fake_llm_generator_port.py` | Delete | Remove the outdated class/file name. |
| `src/domain/tests/quality/fake_llm_generator_adapter.py` | Create | New hexagonal adapter implementing `LlmGeneratorPort` with options signature. |
| `src/domain/tests/quality/test_quality_analyzer.py` | Modify | Update imports and usages to reference `FakeLlmGeneratorAdapter`. |
| `tests/test_main_cli_args.py` | Modify | Add TDD test `test_exits_1_when_save_word_report_fails` asserting exit code `1` on Word save failure. |

## Interfaces / Contracts

### `FakeLlmGeneratorAdapter`

```python
from src.domain.ports.llm_generator_port import LlmGeneratorPort

class FakeLlmGeneratorAdapter(LlmGeneratorPort):
    def __init__(self, responses: list[str]) -> None:
        self._responses = responses
        self.call_count = 0
        self.received_prompts: list[str] = []
        self.received_options: list[dict | None] = []

    def generate(self, prompt: str, options: dict | None = None) -> str:
        self.received_prompts.append(prompt)
        self.received_options.append(options)
        response = self._responses[self.call_count]
        self.call_count += 1
        return response
```

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit (CLI) | CLI exit code and JSON save preservation when Word report fails. | Mock `SilvinaEditorialAssistant`, returning `False` for `save_word_report`, assert `SystemExit` with `1`, assert `save_json_report` was called, and verify stderr/stdout contains `"Error: No se pudo guardar el reporte de Word (DOCX)."`. |
| Unit (Quality Domain) | Quality analysis using the refactored `FakeLlmGeneratorAdapter`. | Run `pytest src/domain/tests/quality/test_quality_analyzer.py` to ensure imports and instantiations function correctly. |

## Migration / Rollout

No migration required.

## Open Questions

None.
