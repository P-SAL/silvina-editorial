# Proposal: Decouple Use Case Ports

## Intent

Technical debt refactoring of `AnalyzeDocumentUseCase` to comply with hexagonal architecture guidelines by removing all 7 direct infrastructure port dependencies and housing local business rules inside 10 coordinated domain services.

## Scope

### In Scope
- Refactor `AnalyzeDocumentUseCase` to coordinate 10 domain services.
- Create 4 new domain services wrapping port calls: `DocumentContentExtractor`, `CitationExtractor`, `DocumentFormatInspector`, and `GrammarChecker`.
- Enhance `ApaValidator` to handle citation filtering and location tuple creation internally.
- Enhance `StructureValidator` to handle empty document checks and references/development filtering internally.
- Refactor `AnalyzeDocumentUseCaseWiring` to wire the new dependency graph.
- Update unit tests (`test_analyze_document_use_case.py` and `test_analyze_document_use_case_wiring.py`).

### Out of Scope
- Modifying underlying port interfaces or adapter implementations.
- Modifying reporting, export, or CLI/Gradio delivery layers.

## Capabilities

### New Capabilities
- None

### Modified Capabilities
- `analyze-document`: Decouple orchestrator by shifting infrastructure dependencies and local business logic into domain services.
- `validate-apa`: Move AUTHOR_YEAR filtering and paragraph text lookup from orchestrator into `ApaValidator`.
- `validate-structure`: Move empty document validation and development/references section filtering from orchestrator into `StructureValidator`.

## Approach

1. **New Domain Services (SRP & Flat Hierarchies)**:
   - `DocumentContentExtractor`: Wraps text, content, and character count ports. Handles fallback logic.
   - `CitationExtractor`: Wraps citation and reference extraction ports.
   - `DocumentFormatInspector`: Wraps format inspection port.
   - `GrammarChecker`: Wraps grammar checking port and error scoring.
2. **Enhanced Validators**:
   - `ApaValidator`: Receives `list[CitationDTO]` and `list[str]` (paragraphs). Performs filtering for `AUTHOR_YEAR` type and constructs location previews internally.
   - `StructureValidator`: Handles `DocumentEmpty` checks and filters out `DEVELOPMENT` and `REFERENCES` (when references exist).
3. **Orchestrator & Wiring Refactor**:
   - Remove private helper methods. Constructor injects 10 domain services. Update composition root.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/application/analyze_document_use_case.py` | Modified | Coordinate 10 domain services, remove ports/logic |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Modified | Wire new services and dependencies |
| `src/domain/document/document_content_extractor.py` | New | Domain service for content extraction |
| `src/domain/citation/citation_extractor.py` | New | Domain service for citations and references |
| `src/domain/document/document_format_inspector.py` | New | Domain service for format inspection |
| `src/domain/grammar/grammar_checker.py` | New | Domain service for grammar checks and scoring |
| `src/domain/citation/apa_validator.py` | Modified | Encapsulate filtering and lookup |
| `src/domain/structure/structure_validator.py` | Modified | Encapsulate empty check and post-filtering |
| `src/application/tests/test_analyze_document_use_case.py` | Modified | Mock new domain services |
| `src/infrastructure/tests/test_analyze_document_use_case_wiring.py` | Modified | Verify new dependency graph |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Test mock breakage | High | Update tests to assert mock interaction on domain services |
| Wiring complexity | Low | Validate instantiation in wiring integration tests |

## Rollback Plan

Revert git changes to return to direct port dependencies and local orchestrator logic.

## Dependencies

- None

## Success Criteria

- [ ] All 10 domain services are correctly constructed and wired.
- [ ] `AnalyzeDocumentUseCase` has zero direct infrastructure port dependencies.
- [ ] All existing unit tests pass successfully.
