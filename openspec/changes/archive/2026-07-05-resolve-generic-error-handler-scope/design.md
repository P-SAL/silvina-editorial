# Design: Resolve Generic Error Handler Scope

## Technical Approach

The goal of this change is to enforce clean hexagonal architecture boundaries by removing the `@generic_error_handler` decorator and its import statement from five infrastructure adapters:
1. `DocxTextAdapter` (`src/infrastructure/adapters/document/docx_text_adapter.py`)
2. `DocxCitationAdapter` (`src/infrastructure/adapters/document/docx_citation_adapter.py`)
3. `DocxReferenceAdapter` (`src/infrastructure/adapters/document/docx_reference_adapter.py`)
4. `DocxEumicAdapter` (`src/infrastructure/adapters/document/docx_eumic_adapter.py`)
5. `OllamaGeneratorAdapter` (`src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`)

Infrastructure adapters should perform concrete actions and either raise specific domain exceptions or propagate third-party/system errors. Application use cases (the orchestration boundary) remain decorated with `@generic_error_handler`, which handles logging unexpected exceptions and wrapping them in `SrcGenericError`. Removing it from adapters prevents redundant logging and incorrect wrapping before exceptions reach the orchestrator.

This implementation maps to Approach 1 of the proposal and complies with the updated specifications in `openspec/specs/extract-citations/spec.md`.

## Architecture Decisions

### Decision: Scope of generic_error_handler

| Option | Tradeoff | Decision |
|---|---|---|
| **Option 1: Apply to both Adapters and Use Cases** | Redundant decoration, duplicate log entries, and premature exception wrapping. Violates clean architecture boundaries. | Rejected |
| **Option 2: Apply only to Application Use Cases** | Keeps adapter code simple and focused. Ensures unexpected errors are uniformly logged and wrapped in `SrcGenericError` once at the use case boundary. | **Selected** |
| **Option 3: Adapter-specific lightweight loggers** | Adds unnecessary complexity and logging noise without any design benefit. | Rejected |

## Data Flow

Before this change, unhandled errors inside adapters were intercepted by the adapter's `@generic_error_handler`, logged, and raised as `SrcGenericError`. This `SrcGenericError` was caught again by the usecase's `@generic_error_handler`, leading to double logs.

After this change, exceptions bubble up directly from the adapters to the usecase layer, where they are caught, logged, and wrapped:

```
[Controller / CLI]
       │
       ▼
[Use Case (@generic_error_handler)]
       │
       ▼
[Infrastructure Adapter] (No decorator; propagates domain/raw exceptions)
       │
       ▼
[External Library / Resource] (e.g., python-docx, Ollama client)
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/infrastructure/adapters/document/docx_text_adapter.py` | Modify | Remove `generic_error_handler` import and decorator from `read_paragraphs`. |
| `src/infrastructure/adapters/document/docx_citation_adapter.py` | Modify | Remove `generic_error_handler` import and decorator from `extract_citations`. |
| `src/infrastructure/adapters/document/docx_reference_adapter.py` | Modify | Remove `generic_error_handler` import and decorator from `extract_references`. |
| `src/infrastructure/adapters/document/docx_eumic_adapter.py` | Modify | Remove `generic_error_handler` import and decorator from `inspect`. |
| `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py` | Modify | Remove `generic_error_handler` import and decorator from `generate`. |
| `openspec/specs/extract-citations/spec.md` | Modify | Remove `@generic_error_handler` requirements from `R7` and `R8`. |

## Interfaces / Contracts

No changes to public port definitions, method signatures, return values, or DTO contracts. The interfaces remain:
- `DocumentTextPort.read_paragraphs(path: str) -> list[str]`
- `CitationExtractionPort.extract_citations(docx_path: str) -> list[CitationDTO]`
- `ReferenceExtractionPort.extract_references(docx_path: str) -> tuple[list[ReferenceDTO], str]`
- `DocumentFormatInspectionPort.inspect(docx_path: str, word_count: int) -> list[EumicViolationDTO]`
- `LlmGeneratorPort.generate(prompt: str, options: dict | None = None) -> str`

## Testing Strategy

All existing tests for these adapters expect specific domain errors or verify correct behavior under normal/exceptional conditions. Since `@generic_error_handler` was only wrapping unexpected exceptions (which the tests do not raise, or mock client errors are already translated explicitly to domain exceptions inside the adapters like `LanguageModelUnavailable` or `DocumentUnreadable`), removing the decorator does not break any tests.

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | Verify adapters still raise expected domain exceptions under failure conditions. | Run existing adapter tests: `test_docx_text_adapter.py`, `test_docx_citation_adapter.py`, `test_docx_reference_adapter.py`, `test_docx_eumic_adapter.py`, `test_ollama_generator_adapter.py`. |
| Integration / E2E | Verify use cases catch raw adapter exceptions and wrap them using `@generic_error_handler` at the boundary. | Run all use case, E2E, and smoke tests. |

## Migration / Rollout

No data migration, feature flags, or phased rollout required. All adapter code changes will be deployed in a single atomic release.

## Open Questions

None.
