# Proposal: Resolve Generic Error Handler Scope

## Intent

Clean up architectural boundary violations where the `@generic_error_handler` decorator was applied to infrastructure adapters instead of being restricted to the application layer / use case boundary. This ensures unexpected infrastructure exceptions are logged and wrapped in `SrcGenericError` once at the outermost application boundary rather than redundantly, keeping adapter code simpler and stack traces cleaner.

## Scope

### In Scope
- Remove the `@generic_error_handler` decorator and its import from:
  1. `DocxTextAdapter.read_paragraphs` in `src/infrastructure/adapters/document/docx_text_adapter.py`
  2. `DocxCitationAdapter.extract_citations` in `src/infrastructure/adapters/document/docx_citation_adapter.py`
  3. `DocxReferenceAdapter.extract_references` in `src/infrastructure/adapters/document/docx_reference_adapter.py`
  4. `DocxEumicAdapter.inspect` in `src/infrastructure/adapters/document/docx_eumic_adapter.py`
  5. `OllamaGeneratorAdapter.generate` in `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`
- Update system specification for citation extraction (`openspec/specs/extract-citations/spec.md`) to remove `@generic_error_handler` requirements at the adapter level.

### Out of Scope
- Introducing new custom domain exceptions.
- Modifying `@generic_error_handler` implementation or its use case decorations.
- Altering other adapters or application-layer services.

## Capabilities

### New Capabilities
None

### Modified Capabilities
- `extract-citations`: Remove requirements specifying `@generic_error_handler` decorations on `DocxCitationAdapter.extract_citations` and `DocxReferenceAdapter.extract_references` methods.

## Approach

1. Remove `@generic_error_handler` annotations and imports from the 5 target adapter files.
2. Update the specification file `openspec/specs/extract-citations/spec.md` to remove the `@generic_error_handler` requirement on `DocxCitationAdapter` and `DocxReferenceAdapter`.
3. Run the existing test suite using `.venv/Scripts/pytest` to verify that all tests pass.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/infrastructure/adapters/document/docx_text_adapter.py` | Modified | Remove `@generic_error_handler` from `read_paragraphs`. |
| `src/infrastructure/adapters/document/docx_citation_adapter.py` | Modified | Remove `@generic_error_handler` from `extract_citations`. |
| `src/infrastructure/adapters/document/docx_reference_adapter.py` | Modified | Remove `@generic_error_handler` from `extract_references`. |
| `src/infrastructure/adapters/document/docx_eumic_adapter.py` | Modified | Remove `@generic_error_handler` from `inspect`. |
| `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py` | Modified | Remove `@generic_error_handler` from `generate`. |
| `openspec/specs/extract-citations/spec.md` | Modified | Remove adapter-level decorator requirements. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Adapter tests expecting `SrcGenericError` | Low | Unit tests pass because they expect specific domain exceptions or do not mock the decorator. |
| Raw python-docx exceptions bubbling up to UI | Low | Use cases remain decorated with `@generic_error_handler` and will wrap unhandled third-party errors into `SrcGenericError` at the orchestrator boundary. |

## Rollback Plan

Revert codebase to the git commit prior to applying these changes:
```bash
git checkout HEAD -- src/infrastructure/adapters/
git checkout HEAD -- openspec/specs/extract-citations/spec.md
```

## Dependencies

- None

## Success Criteria

- [ ] All 5 target adapter methods are free of `@generic_error_handler` decorators and imports.
- [ ] Spec `openspec/specs/extract-citations/spec.md` is updated to remove adapter decorator requirements.
- [ ] All 589 tests in the test suite pass successfully (`.venv/Scripts/pytest`).
