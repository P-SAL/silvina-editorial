## Exploration: add-unittest-suite

### Current State
The system has a single test file `tests/test_structure_validator.py` which verifies `StructureValidator` using basic unit tests. The rest of the modules under `business_logic/`, `data_access/`, and root scripts (`apa_validator.py`, `eumic_verifier.py`) are not covered by any automated unit tests. Some modules like `word_counter.py` rely on Windows COM automation (win32com) which fails on non-Windows/CI systems, others like `quality_analyzer.py` and `article_classifier.py` depend on external LLM services (Ollama) and local endpoints, and `gramatica_checker.py` loads `language_tool_python` which requires a Java environment.

### Affected Areas
- `tests/test_apa_validator.py` — New test suite to verify citation format validation rules deterministically.
- `tests/test_structure_analyzer.py` — New test suite to verify structure analysis without mocks.
- `tests/test_citation_matcher.py` — New test suite for matching citations to reference lists using in-memory mock lists.
- `tests/test_word_reader.py` — New test suite for docx reading, requiring mocks for python-docx elements.
- `tests/test_citation_parser.py` — New test suite for XML-based citation extraction, requiring zipfile / XML mocking.
- `tests/test_reference_parser.py` — New test suite for reference list parsing, requiring XML/zipfile mocking.
- `tests/test_content_extractor.py` — New test suite for layout and metadata extraction, requiring mocks for `WordCounter`.
- `tests/test_article_classifier.py` — New test suite for hybrid classification rules, mocking the Ollama API client.
- `tests/test_quality_analyzer.py` — New test suite for LLM-based quality analysis, mocking Ollama API responses.
- `tests/test_gramatica_checker.py` — New test suite for grammar checking, mocking LanguageTool.
- `tests/test_eumic_verifier.py` — New test suite for EUMIC standards verification, mocking the `docx` document structure.
- `tests/e2e/test_cli_e2e.py` — New E2E test suite to execute `main.py` CLI workflow end-to-end with mock endpoints and output verification.
- `tests/e2e/test_gradio_e2e.py` — New E2E test suite to run the Gradio interface in test mode using the Gradio test client to simulate file uploads and feedback submissions.

### Approaches

1. **Pure Unit Testing with Mocks and Test Doubles** — Implement all new tests as isolated unit tests using Python's built-in `unittest` and `unittest.mock`. Fully mock file operations, python-docx Document XML structures, zipfile reading, win32com client, LanguageTool, and the Ollama client.
   - **Pros**:
     - Fast execution (under 1 second).
     - Deterministic and independent of the operating system (runs on Linux/Windows CIs) or local software (no MS Word, Java, or Ollama server required).
     - No need to maintain physical file fixtures.
   - **Cons**:
     - Mocking complex objects like python-docx XML or zipfile is tedious and prone to drift if the underlying libraries update or behave differently.
   - **Effort**: Medium

2. **Hybrid Testing (Unit + Lightweight Integration Tests + E2E)** — Keep deterministic in-memory logic under pure unit tests. Use real `.docx` fixtures in `tests/fixtures/` for parsers. Add end-to-end (E2E) tests for both the CLI (`main.py` via in-process execution or `subprocess` with mocked endpoints) and the Gradio app (using `gradio`'s built-in testing client), while strictly mocking external services (Ollama, Windows Word COM) to ensure CI compatibility.
   - **Pros**:
     - Tests for document parsing and E2E workflows (CLI execution, UI interaction) are realistic and verify complete system integration.
     - Protects against regression across the entire pipeline.
   - **Cons**:
     - Requires managing binary `.docx` file fixtures.
     - E2E tests are slower and require more complex setup (running mock Ollama server, simulating Gradio events).
   - **Effort**: Medium-High

### Recommendation
We recommend **Approach 2 (Hybrid Testing with E2E)**. This combines fast, isolated unit tests for core logic, small `.docx` fixtures for parsing verification, and automated E2E tests for the CLI (`main.py`) and Gradio interface (`gradio_app.py`) to verify the full pipeline. External dependencies (Ollama LLM, Windows Word COM API, Java LanguageTool) will be stubbed/mocked during tests to guarantee cross-platform compatibility and fast execution in CI.

### Risks
- **Flaky/Incomplete Mocks**: If the mock return values for Ollama or LanguageTool do not accurately reflect real API structures, tests might pass but the app could fail in production.
- **CI Environments**: If tests accidentally trigger `win32com.client` or Java execution for LanguageTool without fallback checks, CI builds will break.
- **Gradios Testing Client Overhead**: Gradio E2E testing relies on port availability and proper server startup/shutdown cycles.
- **Fixture Maintenance**: Corrupted or out-of-date `.docx` fixtures can cause tests to fail for reasons unrelated to code logic.

### Ready for Proposal
Yes — The orchestrator should proceed to define the proposal and design specs for the `add-unittest-suite` change using Approach 2 (incorporating E2E).

