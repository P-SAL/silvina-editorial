# Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~500 lines (additions + deletions) |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 (Domain Services) → PR 2 (Validators) → PR 3 (UseCase & Wiring) |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: High

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Create 4 new domain services and unit tests | PR 1 | Base branch (feature tracker or main); tests included |
| 2 | Enhance Apa & Structure validators and unit tests | PR 2 | Depends on PR 1; tests included |
| 3 | Refactor Orchestrator & Wiring, and update tests | PR 3 | Depends on PR 2; integrates the complete pipeline |

## Phase 1: Domain Services Foundation (TDD)

- [x] 1.1 (RED) Write unit tests for `DocumentContentExtractor` in `src/domain/tests/document/test_document_content_extractor.py`.
- [x] 1.2 (GREEN) Create `DocumentContentExtractor` in `src/domain/document/document_content_extractor.py` to extract content and fallback counts.
- [x] 1.3 (RED) Write unit tests for `CitationExtractor` in `src/domain/tests/citation/test_citation_extractor.py`.
- [x] 1.4 (GREEN) Create `CitationExtractor` in `src/domain/citation/citation_extractor.py` to extract citations/references.
- [x] 1.5 (RED) Write unit tests for `DocumentFormatInspector` in `src/domain/tests/document/test_document_format_inspector.py`.
- [x] 1.6 (GREEN) Create `DocumentFormatInspector` in `src/domain/document/document_format_inspector.py` to wrap format inspection.
- [x] 1.7 (RED) Write unit tests for `GrammarChecker` in `src/domain/tests/grammar/test_grammar_checker.py`.
- [x] 1.8 (GREEN) Create `GrammarChecker` in `src/domain/grammar/grammar_checker.py` to check grammar and score errors.

## Phase 2: Enhanced Validators (TDD)

- [x] 2.1 (RED) Write unit tests in `src/domain/tests/citation/test_apa_validator_skip_patterns.py` for new AUTHOR_YEAR filtering and paragraph text retrieval.
- [x] 2.2 (GREEN) Modify `ApaValidator` in `src/domain/citation/apa_validator.py` to filter citations and fetch paragraph previews with fallback.
- [x] 2.3 (RED) Write unit tests in `src/domain/tests/structure/test_structure_validator_scientific.py` for empty check and section exclusions.
- [x] 2.4 (GREEN) Modify `StructureValidator` in `src/domain/structure/structure_validator.py` to raise `DocumentEmpty` and filter missing sections.

## Phase 3: Orchestrator & Wiring Refactoring (TDD)

- [x] 3.1 (RED) Update orchestrator unit tests in `src/application/tests/test_analyze_document_use_case.py` to mock 10 services and expect sequential calls.
- [x] 3.2 (GREEN) Refactor `AnalyzeDocumentUseCase` in `src/application/analyze_document_use_case.py` to coordinate the 10 services.
- [x] 3.3 (RED) Update wiring integration tests in `src/infrastructure/tests/test_analyze_document_use_case_wiring.py` to assert new dependency graph.
- [x] 3.4 (GREEN) Refactor `AnalyzeDocumentUseCaseWiring` in `src/infrastructure/wirings/analyze_document_use_case_wiring.py` to wire all 10 domain services.

## Phase 4: Cleanup & Verification

- [x] 4.1 Run linter and formatter using `ruff check` and `ruff format` on all new and modified files.
- [x] 4.2 Verify all tests pass by running pytest using `.venv/Scripts/pytest`.
