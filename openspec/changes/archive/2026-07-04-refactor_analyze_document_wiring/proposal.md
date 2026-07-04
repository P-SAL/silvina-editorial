# Proposal: Refactor Document Analysis Wiring

## Intent
Eliminate 10 redundant sub-use cases and their corresponding wirings. Instantiations and invocations of ports and domain services will be done directly within `AnalyzeDocumentUseCase` and its wiring `AnalyzeDocumentUseCaseWiring`. This reduces unnecessary pass-through delegation layers, simplifies system architecture, and aligns with Clean Architecture.

## Scope
### In Scope
- Modify `AnalyzeDocumentUseCase` to directly accept and orchestrate 7 ports, 5 domain services, and 1 builder in its constructor.
- Modify `AnalyzeDocumentUseCaseWiring` to instantiate and inject the 13 direct dependencies.
- Update `test_analyze_document_use_case.py` and `test_analyze_document_use_case_wiring.py` to test the new structure.
- Delete the 10 obsolete sub-use case files, 10 sub-wiring files, and their associated unit tests (approx. 40 files total).
- Preserve `CharacterCountUnavailable` fallback behavior directly inside the orchestrator.

### Out of Scope
- Modification or deletion of `ExportReportUseCase` or `ExportReportWiring`.

## Capabilities
### New Capabilities
- None
### Modified Capabilities
- None

## Approach
1. **Orchestrator Refactor**: Change `AnalyzeDocumentUseCase` to inject ports (`DocumentTextPort`, `ContentExtractionPort`, `CharacterCountPort`, `CitationExtractionPort`, `ReferenceExtractionPort`, `GrammarCheckPort`, `DocumentFormatInspectionPort`), domain services (`ApaValidator`, `ArticleClassifier`, `QualityAnalyzer`, `StructureValidator`, `CitationMatcher`), and the `RecommendationBuilder`.
2. **Wiring Refactor**: Update `AnalyzeDocumentUseCaseWiring` to act as the single composition root for the analysis pipeline, instantiating the required adapters and services.
3. **Obsolete Deletion**: Delete the 10 sub-use cases, 10 wirings, and 20 corresponding test files.
4. **Test Alignment**: Update the orchestrator's tests to mock/fake the 13 direct dependencies.

## Affected Areas
| Area | Impact | Description |
|------|--------|-------------|
| Application Layer | Medium | Refactored `AnalyzeDocumentUseCase`; deleted 10 use case files. |
| Infrastructure Layer | Medium | Refactored `AnalyzeDocumentUseCaseWiring`; deleted 10 wiring files. |
| Test Suite | High | Deleted 20 test files; refactored orchestrator and wiring tests. |

## Risks
| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Regressions in win32com integration | Low | Retain try-except logic for `CharacterCountUnavailable` fallback inside orchestrator. |
| Test coverage gaps | Medium | Refactor `test_analyze_document_use_case.py` carefully to verify exact port/service interactions. |

## Rollback Plan
Revert changes using Git version control to restore deleted use cases, wirings, tests, and original orchestrator signatures.

## Dependencies
- None

## Success Criteria
- [ ] All 10 obsolete sub-use cases and their 10 wirings deleted.
- [ ] `AnalyzeDocumentUseCase` executes successfully with direct ports and services.
- [ ] Word count refinement fallback works as expected.
- [ ] All unit and integration tests pass successfully.
