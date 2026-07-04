# Design: Refactor Slice 14 CLI (refactor-slice-14-cli)

## Technical Approach

Refactor [main.py](file:///E:/Python/silvina-editorial/main.py) to act as a clean, lightweight entry point (driving adapter). The file will be stripped of manual business logic orchestration, hardcoded thresholds, and legacy parsing. Instead, it will:
1. Parse CLI arguments using `argparse`.
2. Instantiate the hexagonal use cases using their production wirings:
   - [AnalyzeDocumentUseCase](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) via `AnalyzeDocumentUseCaseWiring`.
   - [ExportReportUseCase](file:///E:/Python/silvina-editorial/src/application/export_report_use_case.py) via `ExportReportWiring`.
3. Wrap execution in a backwards-compatible `SilvinaEditorialAssistant` shim class to prevent breaking [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) and the E2E tests.
4. Detect Ollama/LLM offline status at startup or during execution and fail fast.
5. Catch domain-specific exceptions inheriting from `BaseSrcError` and map them to appropriate CLI exit codes.

---

## Architecture Decisions

### Decision: Backwards-compatible SilvinaEditorialAssistant wrapper class

* **Choice**: Shim class converting DTOs to legacy dict format for [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) and E2E tests.
* **Alternatives considered**: Rewriting [gradio_app.py](file:///E:/Python/silvina-editorial/gradio_app.py) and E2E tests to consume `ReportInputDTO` directly.
* **Rationale**: Decoupling the legacy UI and tests from the internal domain migration reduces the scope of this refactoring slice and prevents regressions, while allowing downstream consumer migrations to happen independently.

### Decision: Configurable output paths

* **Choice**: Add argparse arguments (`--output-dir`, `--word-report-path`, `--json-report-path`) with the current behavior (same folder as input docx) as the default.
* **Alternatives considered**: Restricting output directories to hardcoded paths or environment variables.
* **Rationale**: Giving CLI users full control over output locations improves usability in automated scripting pipelines, while defaulting to the input file's folder maintains backwards compatibility for existing manual runs.

### Decision: Immediate failure on Ollama offline

* **Choice**: If a `LanguageModelUnavailable` (or other Ollama connection issue) is raised or detected, fail and abort the CLI execution immediately.
* **Alternatives considered**: Mocking the LLM responses or printing warnings but generating incomplete reports.
* **Rationale**: Document classification and quality analysis rely heavily on the LLM backend. Proceeding with dummy values results in invalid reports, which degrades trust. Failing fast is the safest, most transparent approach.

### Decision: Exception mapping to exit codes

* **Choice**: Catch `BaseSrcError` and its subclasses, print clean messages to `stderr`, and exit with `1` (or `2` for argument issues).
* **Alternatives considered**: Letting raw exceptions propagate to the console with full Python tracebacks.
* **Rationale**: Standard CLI design dictates hiding internal technical stack traces from end-users on expected validation/domain failures, providing clean, actionable error messages instead. Full tracebacks will be reserved for unexpected generic exceptions.

---

## Data Flow

```
   [ CLI Command / gradio_app.py ]
                 │
                 ▼
     [ SilvinaEditorialAssistant ] (Shim Class)
                 │
      ┌──────────┴──────────┐
      ▼                     ▼
[ AnalyzeDocumentUseCase ]  [ ExportReportUseCase ]
      │ (Executes Pipeline) │ (Writes Reports)
      ▼                     ▼
[ ReportInputDTO ] ──→ [ legacy dict ] ──→ [ .docx & .json files ]
```

---

## File Changes

| File | Action | Description |
|------|--------|-------------|
| [main.py](file:///E:/Python/silvina-editorial/main.py) | Modify | Refactor to use `AnalyzeDocumentUseCaseWiring` and `ExportReportWiring`. Implement the DTO-to-legacy mapping helper and custom argument parser. |
| `main_legacy.py` | Create | Preserved snapshot of the old `main.py` implementation (done in previous step). |

---

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | DTO to legacy dictionary converter | Test helper functions in `main.py` with mock DTO inputs to assert key parity. |
| Integration | Argument parsing and file generation paths | Verify argument parser configuration and correct output path defaults. |
| E2E | End-to-end execution of refactored CLI | Run [test_cli_e2e.py](file:///E:/Python/silvina-editorial/tests/e2e/test_cli_e2e.py) with mocked Ollama endpoints to verify zero regressions. |

---

## Migration / Rollout

No data migration required. In case of unexpected production issues, standard rollback is to replace `main.py` with the copy in `main_legacy.py`.

---

## Open Questions

* **None**: The target scope is well-defined and backwards-compatibility requirements are fully mapped.
