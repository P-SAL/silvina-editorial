## Exploration: Refactor Document Analysis Wiring

### Current State
Currently, `AnalyzeDocumentUseCase` orchestrates the academic document analysis workflow by delegating tasks to 10 sub-use cases (`ReadDocumentUseCase`, `ExtractContentUseCase`, etc.). In the wiring layer, `AnalyzeDocumentUseCaseWiring` instantiates the orchestrator by delegating instantiation to 10 separate sub-wiring classes, each corresponding to one sub-use case.

This structure introduces excessive delegation layers (middleman classes) with no independent business logic, violating clean architecture simplicity guidelines where a use case should represent a complete, user-initiated action.

### Affected Areas
#### 1. Core Application & Orchestration
- `src/application/analyze_document_use_case.py` — Refactor constructor to inject ports and domain services directly instead of sub-use cases, and update orchestration logic.
- `src/infrastructure/wirings/analyze_document_use_case_wiring.py` — Refactor to build and inject concrete adapters and domain services directly into the orchestrator.

#### 2. Files to be Eliminated (Use Cases)
- `src/application/read_document_use_case.py`
- `src/application/extract_content_use_case.py`
- `src/application/extract_citations_use_case.py`
- `src/application/validate_apa_use_case.py`
- `src/application/check_grammar_use_case.py`
- `src/application/classify_article_use_case.py`
- `src/application/analyze_quality_use_case.py`
- `src/application/validate_structure_use_case.py`
- `src/application/match_citations_use_case.py`
- `src/application/verify_eumic_use_case.py`

#### 3. Files to be Eliminated (Wirings)
- `src/infrastructure/wirings/read_document_use_case_wiring.py`
- `src/infrastructure/wirings/extract_content_use_case_wiring.py`
- `src/infrastructure/wirings/extract_citations_use_case_wiring.py`
- `src/infrastructure/wirings/validate_apa_wiring.py`
- `src/infrastructure/wirings/check_grammar_use_case_wiring.py`
- `src/infrastructure/wirings/classify_article_use_case_wiring.py`
- `src/infrastructure/wirings/analyze_quality_use_case_wiring.py`
- `src/infrastructure/wirings/validate_structure_wiring.py`
- `src/infrastructure/wirings/match_citations_use_case_wiring.py`
- `src/infrastructure/wirings/verify_eumic_use_case_wiring.py`

#### 4. Tests to be Eliminated
- `src/application/tests/test_read_document_use_case.py`
- `src/application/tests/test_extract_content_use_case.py`
- `src/application/tests/test_extract_citations_use_case.py`
- `src/application/tests/test_validate_apa_use_case.py`
- `src/application/tests/test_check_grammar_use_case.py`
- `src/application/tests/test_classify_article_use_case.py`
- `src/application/tests/test_analyze_quality_use_case.py`
- `src/application/tests/test_validate_structure_use_case.py`
- `src/application/tests/test_match_citations_use_case.py`
- `src/application/tests/test_verify_eumic_use_case.py`
- `src/infrastructure/tests/test_read_document_use_case_wiring.py`
- `src/infrastructure/tests/test_extract_content_use_case_wiring.py`
- `src/infrastructure/tests/test_extract_citations_use_case_wiring.py`
- `src/infrastructure/tests/test_validate_apa_wiring.py`
- `src/infrastructure/tests/test_check_grammar_use_case_wiring.py`
- `src/infrastructure/tests/test_classify_article_use_case_wiring.py`
- `src/infrastructure/tests/test_analyze_quality_use_case_wiring.py`
- `src/infrastructure/tests/test_validate_structure_wiring.py`
- `src/infrastructure/tests/test_match_citations_use_case_wiring.py`
- `src/infrastructure/tests/test_verify_eumic_use_case_wiring.py`

#### 5. Tests to be Refactored
- `src/application/tests/test_analyze_document_use_case.py` — Update mocks and verify direct ports/services injection.
- `src/infrastructure/tests/test_analyze_document_use_case_wiring.py` — Verify direct instantiation of ports and domain services.

#### 6. Export Report Flow (Optional cleanup to maintain "Only One Use Case" rule)
- `src/application/export_report_use_case.py` — Eliminate.
- `src/infrastructure/wirings/export_report_wiring.py` — Eliminate or refactor to return the report export adapter.
- `src/application/tests/test_export_report_use_case_error_propagation.py` & `test_export_report_use_case_success.py` — Eliminate.
- `src/infrastructure/tests/test_export_report_wiring.py` — Eliminate or refactor.
- `main.py` — Call adapter directly.
- `gradio_app.py` — Call adapter directly.

### Approaches
1. **Full Hexagonal Clean Refactoring (Eliminate Sub-Use Cases)**
   - **Description**: Delete all 10 sub-use cases and their wirings. Inject the domain services and ports directly into `AnalyzeDocumentUseCase`. `AnalyzeDocumentUseCaseWiring` acts as the single composition root for the analysis pipeline.
   - **Pros**:
     - Reduces class and file count significantly (removes ~40 files).
     - Minimizes delegation and boilerplate.
     - Perfectly conforms to hexagonal architecture intent: use cases model complete user actions, not internal implementation steps.
   - **Cons**:
     - Requires modifying the constructor of `AnalyzeDocumentUseCase` and updating its unit tests.
   - **Effort**: Medium

2. **Wiring-Only Refactoring (Retain Sub-Use Cases, Eliminate Sub-Wirings)**
   - **Description**: Keep the 10 sub-use cases but delete their wiring files. Have `AnalyzeDocumentUseCaseWiring` instantiate all the sub-use cases and sub-dependencies.
   - **Pros**:
     - Avoids changing the constructor signature of `AnalyzeDocumentUseCase`.
   - **Cons**:
     - Leaves 10 redundant pass-through use cases in the codebase, failing to solve the root problem of over-engineering and middleman classes.
   - **Effort**: Low

### Recommendation
**Approach 1** is highly recommended. It fully aligns with clean architecture by eliminating redundant wrappers and keeping the domain/application layers lean and focused. It also matches the explicit intent of the instruction.

Regarding `ExportReportUseCase`: to strictly enforce "the only use case in the system should be AnalyzeDocumentUseCase", `ExportReportUseCase` should be eliminated and the controllers (`main.py` and `gradio_app.py`) should interact directly with `ReportExportPort` via an adapter instantiated by a simple factory or wiring in infrastructure.

### Risks
- **Tight Coupling in `AnalyzeDocumentUseCaseWiring`**: Having one wiring manage all dependencies makes it large, but this is the appropriate role for a composition root.
- **Test Suite Updates**: Extensive tests exist for sub-use cases and sub-wirings; these will need to be cleaned up/removed carefully without losing test coverage for adapter or domain service behavior.
- **Win32Com Word Count Integration**: Ensure word-count refinement logic is preserved correctly in the orchestrator.

### Ready for Proposal
Yes — the next step is to create a detailed proposal detailing the exact signature and implementation updates.
