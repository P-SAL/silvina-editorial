# Tasks: Add Unittest Suite

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 600-800 lines |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | PR 1: Unit + Integration Tests; PR 2: E2E Tests |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: stacked-to-main
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Setup, Unit, and Integration Tests | PR 1 | Copies docx fixture, mocks win32com, and tests validators, classifiers, readers, and parsers. |
| 2 | End-to-End CLI and UI Validation | PR 2 | Implements E2E tests for the CLI (main.py) and UI (gradio_app.py). |

## Phase 1: Setup and Fixtures

- [ ] 1.1 Copy fixture `E:\Python\silvina-doc\capacidades_razonamiento_emergente_LLMs.docx` to `tests/fixtures/capacidades_razonamiento_emergente_LLMs.docx`.
- [ ] 1.2 Set up sys.modules mocks for `win32com`, `win32com.client`, and `pythoncom` inside `tests/__init__.py`.

## Phase 2: Unit Tests

- [ ] 2.1 Implement unit tests for `apa_validator.py` and `eumic_verifier.py` in `tests/test_apa_validator.py` and `tests/test_eumic_verifier.py`.
- [ ] 2.2 Implement unit tests for `business_logic/structure_validator.py` and `business_logic/structure_analyzer.py` in `tests/test_structure_validator.py` and `tests/test_structure_analyzer.py`.
- [ ] 2.3 Implement unit tests for `business_logic/citation_matcher.py` in `tests/test_citation_matcher.py`.
- [ ] 2.4 Implement unit tests for `data_access/content_extractor.py` and `data_access/word_counter.py` in `tests/test_content_extractor.py` and `tests/test_word_counter.py`.
- [ ] 2.5 Implement mock-based unit tests for `business_logic/article_classifier.py` and `business_logic/quality_analyzer.py` in `tests/test_article_classifier.py` and `tests/test_quality_analyzer.py` using Ollama client mocks.
- [ ] 2.6 Implement mock-based unit tests for `business_logic/gramatica_checker.py` in `tests/test_gramatica_checker.py` using mock LanguageTool.

## Phase 3: Integration Tests

- [ ] 3.1 Implement integration tests for `data_access/word_reader.py` in `tests/test_word_reader.py` using the `.docx` fixture.
- [ ] 3.2 Implement integration tests for `data_access/citation_parser.py` in `tests/test_citation_parser.py` using the `.docx` fixture.
- [ ] 3.3 Implement integration tests for `data_access/reference_parser.py` in `tests/test_reference_parser.py` using the `.docx` fixture.

## Phase 4: E2E Tests

- [ ] 4.1 Implement E2E tests for the CLI `main.py` in `tests/e2e/test_cli_e2e.py` by mocking `sys.argv` and validating output reports.
- [ ] 4.2 Implement E2E tests for the UI `gradio_app.py` in `tests/e2e/test_gradio_e2e.py` using the Gradio testing client.
