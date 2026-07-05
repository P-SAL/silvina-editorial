## Exploration: resolve-generic-error-handler-scope

### Current State
Currently, the `@generic_error_handler` decorator is applied to both the application use cases (`AnalyzeDocumentUseCase`, `ExportReportUseCase`) and various infrastructure adapters. This violates the clean architecture boundary where error interception, generic wrapping into `SrcGenericError`, and logging should be handled at the application layer/use case boundary. When both layers are decorated, unexpected errors can be caught and wrapped early, leading to redundant logging or incorrect exception types propagation.

The target adapters handle errors and custom exceptions as follows:
- **`DocxTextAdapter`**: Raises `DocumentNotFound` if the path doesn't exist, and `DocumentUnreadable` (wrapping python-docx exceptions) if parsing fails.
- **`DocxCitationAdapter`**: Does not raise custom exceptions (only imports `CitationParsingFailed` without raising it). It delegates loading to the text port and parses using regex.
- **`DocxReferenceAdapter`**: Does not raise custom exceptions (only imports `ReferenceParsingFailed` without raising it). It delegates loading to the text port and parses using regex.
- **`DocxEumicAdapter`**: Does not raise custom exceptions. It uses python-docx `Document` to load/verify formatting.
- **`OllamaGeneratorAdapter`**: Catches Ollama client/connection errors and wraps/raises them as `LanguageModelUnavailable`.

### Affected Areas
- `src/infrastructure/adapters/document/docx_text_adapter.py` — Contains `@generic_error_handler` on the `read_paragraphs` method.
- `src/infrastructure/adapters/document/docx_citation_adapter.py` — Contains `@generic_error_handler` on the `extract_citations` method.
- `src/infrastructure/adapters/document/docx_reference_adapter.py` — Contains `@generic_error_handler` on the `extract_references` method.
- `src/infrastructure/adapters/document/docx_eumic_adapter.py` — Contains `@generic_error_handler` on the `inspect` method.
- `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py` — Contains `@generic_error_handler` on the `generate` method.

### Approaches
1. **Remove `@generic_error_handler` from adapters and let exceptions propagate** — Completely remove the decorator imports and annotations from the target adapters. Let all domain exceptions (e.g. `DocumentNotFound`, `LanguageModelUnavailable`) and unexpected exceptions bubble up naturally. Since the use cases are decorated with `@generic_error_handler`, they will log and wrap unexpected errors at the application layer boundary.
   - Pros: Correct application of clean hexagonal architecture principles. Simpler adapter code, no redundant logging, and cleaner stack traces.
   - Cons: None.
   - Effort: Low.

2. **Implement adapter-specific error logging decorators** — Replace `@generic_error_handler` in the adapters with a lightweight logger-only decorator that doesn't wrap exceptions in `SrcGenericError`.
   - Pros: Immediate logging at the infrastructure boundary.
   - Cons: Violates separation of concerns. Logging at the adapter boundary adds noise and doesn't benefit from use case orchestration context.
   - Effort: Medium.

### Recommendation
I recommend **Approach 1**. It is standard for clean architecture: adapters should perform their concrete operations and raise either specific domain exceptions or propagate third-party exceptions. The use case boundary (which represents the application transaction/orchestration layer) is the correct place to run `@generic_error_handler` to log unhandled errors, wrap unexpected infrastructure failures into `SrcGenericError`, and let expected domain errors pass through to the delivery layer.

### Risks
- **Test cases asserting adapter behavior**: Unit tests for the adapters (e.g., `TestDocxTextAdapter`, `TestOllamaGeneratorAdapter`) check that custom exceptions are raised. Fortunately, these exceptions (`DocumentNotFound`, `DocumentUnreadable`, `LanguageModelUnavailable`) are explicitly raised inside the adapter method logic, not by the decorator. Therefore, removing the decorator will not break these tests.
- **Uncaught errors in direct adapter invocation**: If an adapter is invoked directly outside a decorated usecase (e.g., in a standalone script), unhandled exceptions will not be logged/wrapped. This is desired behavior since external scripts should handle raw exceptions.

### Ready for Proposal
Yes — The orchestrator should proceed to create the proposal and task list for removing the `@generic_error_handler` from the five target adapters.
