# Proposal: Add Unittest Suite

## Intent

Establish a robust automated test suite to verify code correctness, prevent regressions across citation parsing and classification components, and guarantee cross-platform execution (Windows, Linux, CI) without external service dependencies.

## Scope

### In Scope
- Unit tests for all modules in `business_logic/`, `data_access/`, and root scripts using Python's standard `unittest` library.
- Integration tests using local `.docx` fixtures to verify parsing and metadata extraction.
- End-to-end (E2E) tests for the CLI (`main.py`) via mock interfaces.
- Gradio UI E2E tests using the native Gradio testing client.
- Mock wrappers for Ollama APIs, LanguageTool Java client, and MS Word COM (`win32com`).

### Out of Scope
- Testing physical MS Word execution on non-Windows operating systems.
- Live HTTP/API connections to Ollama or LanguageTool instances during test execution.

## Capabilities

### New Capabilities
- `testing-suite`: Implement a full unit, integration, and E2E testing framework using Python's standard `unittest` library.

### Modified Capabilities
- None

## Approach

We will use a hybrid approach combining unit, integration, and E2E tests. We will use the built-in `unittest` and `unittest.mock` libraries to isolate components. Standard mock objects will stub external LLM, grammar checks, and COM-based actions. The CLI and Gradio UI will run E2E scenarios using fake inputs and the Gradio testing client.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `tests/` | Modified | Add comprehensive suite of unit, integration, and E2E tests. |
| `tests/fixtures/` | New | Add lightweight `.docx` file fixtures for parsing testing. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| CI Failure | Med | Ensure absolute isolation of `win32com` and Java requirements in CI environment via conditional mocks/skips. |
| Mock Mismatch | Med | Validate mock responses against actual Ollama and LanguageTool schemas to avoid drift. |

## Rollback Plan

Revert all additions under the `tests/` directory. Since no application/production source code is modified, rollback has zero risk and zero production impact.

## Dependencies

- Python 3.10+
- `gradio` and `python-docx` packages installed in the local environment.

## Success Criteria

- [ ] All unit, integration, and E2E tests in the `tests/` directory pass successfully.
- [ ] The CLI workflow execution is validated end-to-end with mock endpoints.
- [ ] The Gradio interface is verified using its native testing client.
