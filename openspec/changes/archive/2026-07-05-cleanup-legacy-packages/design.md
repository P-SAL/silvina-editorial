# Design: Cleanup Legacy Packages

## Technical Approach
We will clean up the project structure by deleting all legacy source directories and root-level files that have been fully replaced by the new clean hexagonal architecture in `src/`.
To prevent regression and preserve test coverage, we will adapt the smoke/parity tests in `tests/smoke/` to execute directly against the corresponding new ports and use cases under `src/`, rather than comparing them against legacy code.
Finally, we will update the E2E test suites to remove mocks/patches targeting legacy packages, and update `pytest.ini` to remove legacy test paths.

## Architecture Decisions

### Decision: Smoke Tests Retention Strategy
| Option | Tradeoff | Decision |
|---|---|---|
| Delete all smoke tests entirely | Simplifies the cleanup, but reduces test coverage on real document fixtures. | Rejected. |
| Adapt smoke tests to assert against `src/` classes directly | Preserves valuable integration/smoke coverage on actual documents, but requires rewriting the tests to remove legacy imports and compare against fixed expected outputs. | **Chosen**. Provides robust regression testing for the migrated codebase. |

### Decision: Win32com Patching in E2E Tests
| Option | Tradeoff | Decision |
|---|---|---|
| Remove legacy win32com mocks only | Simple, but might cause Gradio E2E tests to fail on systems without Word/win32com if they trigger the new Word count adapter. | Rejected. |
| Replace legacy `data_access` patches with `src` patches | Ensures that E2E tests do not attempt to call live win32com on CI/non-Windows systems for the new adapters, preserving test portability. | **Chosen**. Replace `data_access.word_counter.WIN32COM_AVAILABLE` patches with `src.infrastructure.adapters.document.win32com_word_count_adapter.WIN32COM_AVAILABLE`. |

## Data Flow
Since this is a cleanup change, no new data flows are introduced. The data flow through the new hexagonal architecture remains as defined:

```
  [User / CLI / Gradio]
         │
         ▼
  [main.py / gradio_app.py] (Controllers)
         │
         ▼
  [src/infrastructure/wirings] (Wiring Assemblies)
         │
         ▼
  [src/application/use_cases] (Use Cases)
      /      \
     ▼        ▼
[src/domain]  [src/infrastructure/adapters] (Ports/Adapters)
```

Data flow for the adapted smoke tests:
```
[Smoke Test Suite] ──(loads fixture docx)──> [DocumentTextPort]
       │                                             │
       │ (returns paragraphs) <──────────────────────┘
       ▼
[Smoke Test Suite] ──(invokes)──> [UseCase/Port/Adapter] ──> [Assert Outputs]
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `domain/` | Delete | Legacy domain models and enums, replaced by `src/domain/` |
| `data_access/` | Delete | Legacy I/O adapters, replaced by `src/infrastructure/adapters/` |
| `business_logic/` | Delete | Legacy services, replaced by `src/application/` use cases |
| `presentation/` | Delete | Legacy report generation and configuration, replaced by `src/infrastructure/adapters/report/` |
| `apa_validator.py` | Delete | Legacy APA validator script, replaced by `src/domain/citation/apa_validator.py` |
| `eumic_verifier.py` | Delete | Legacy EUMIC verifier script, replaced by `src/application/verify_eumic_use_case.py` |
| `config.py` | Delete | Legacy configuration script, replaced by environment variable configuration |
| `main_legacy.py` | Delete | Legacy entry point, replaced by `main.py` |
| `tests/legacy/` | Delete | Unit and integration tests for legacy packages |
| `tests/smoke/test_classify_article_parity.py` | Modify | Adapt to test `src.infrastructure.adapters.llm_generator.ollama_generator_adapter` directly with canned response mocks |
| `tests/smoke/test_extract_content_parity.py` | Modify | Adapt to test `src.infrastructure.adapters.document.docx_content_extraction_adapter` directly |
| `tests/smoke/test_read_document_parity.py` | Modify | Adapt to test `src.infrastructure.adapters.document.docx_text_adapter` directly |
| `tests/smoke/test_validate_structure_parity.py` | Modify | Adapt to test `src.domain.structure.structure_validator` directly |
| `tests/e2e/test_cli_e2e.py` | Modify | Remove legacy mock `data_access.word_counter.WIN32COM_AVAILABLE` |
| `tests/e2e/test_gradio_e2e.py` | Modify | Remove legacy mock `data_access.word_counter.WIN32COM_AVAILABLE` and ensure `src.infrastructure.adapters.document.win32com_word_count_adapter.WIN32COM_AVAILABLE` is patched to `False` |
| `pytest.ini` | Modify | Remove `tests/legacy` from `norecursedirs` |

## Interfaces / Contracts
No new interfaces or contracts are introduced by this cleanup. The smoke tests will interact directly with existing domain DTOs and ports under `src/`.

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit / Domain | Verify domain logic | Run existing `src/domain/tests/` to ensure 100% pass rate. |
| Integration | Verify `src/` adapters using actual docx files | Run the updated `tests/smoke/` tests. |
| E2E | Verify CLI and Gradio app flow | Run `tests/e2e/test_cli_e2e.py` and `tests/e2e/test_gradio_e2e.py`. |

## Migration / Rollout
No migration required.

## Open Questions
None.
