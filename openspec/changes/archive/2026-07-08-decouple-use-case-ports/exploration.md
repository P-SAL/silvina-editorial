## Exploration: Decouple Use Case Ports and Local Logic

### Current State
Currently, [AnalyzeDocumentUseCase](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) directly depends on 7 infrastructure ports:
- [DocumentTextPort](file:///E:/Python/silvina-editorial/src/domain/document/document_text_port.py)
- [ContentExtractionPort](file:///E:/Python/silvina-editorial/src/domain/document/content_extraction_port.py)
- [CharacterCountPort](file:///E:/Python/silvina-editorial/src/domain/document/character_count_port.py)
- [CitationExtractionPort](file:///E:/Python/silvina-editorial/src/domain/document/citation_extraction_port.py)
- [ReferenceExtractionPort](file:///E:/Python/silvina-editorial/src/domain/document/reference_extraction_port.py)
- [GrammarCheckPort](file:///E:/Python/silvina-editorial/src/domain/grammar/grammar_check_port.py)
- [DocumentFormatInspectionPort](file:///E:/Python/silvina-editorial/src/domain/document/document_format_inspection_port.py)

It also houses local business rules/logic:
- Merging character counts with fallback configurations.
- Filtering citations for APA validation and formatting location tuples.
- Checking grammar and mapping error count to a score/feedback level.
- Section completeness validation (filtering out `DEVELOPMENT` and conditionally ignoring `REFERENCES`).
- Converting raw string section types into `SectionName` enums.

This violates clean hexagonal architecture guidelines, which state that a use case must be a thin orchestrator of domain services, must not contain local business rules, and must never inject ports directly.

### Affected Areas
- [analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) — Refactor to remove ports and local logic; inject and coordinate 10 domain services.
- [analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py) — Refactor to build and inject the 10 domain services.
- [document_content_extractor.py](file:///E:/Python/silvina-editorial/src/domain/document/document_content_extractor.py) — Create new domain service wrapping text, content, and count ports.
- [citation_extractor.py](file:///E:/Python/silvina-editorial/src/domain/citation/citation_extractor.py) — Create new domain service wrapping citation and reference extraction ports.
- [document_format_inspector.py](file:///E:/Python/silvina-editorial/src/domain/document/document_format_inspector.py) — Create new domain service wrapping the format inspection port.
- [grammar_checker.py](file:///E:/Python/silvina-editorial/src/domain/grammar/grammar_checker.py) — Create new domain service wrapping the grammar check port and scoring.
- [apa_validator.py](file:///E:/Python/silvina-editorial/src/domain/citation/apa_validator.py) — Enhance with `validate` to filter citations and build location tuples.
- [structure_validator.py](file:///E:/Python/silvina-editorial/src/domain/structure/structure_validator.py) — Enhance with `validate_structure` to check for empty doc and filter sections.
- [test_analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_analyze_document_use_case.py) — Update mocks to focus on domain services.
- [test_analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_analyze_document_use_case_wiring.py) — Update wiring tests to verify new dependency graph.

### Approaches
1. **Port Decoupling via Domain Services (Recommended)**
   - Introduce 4 new domain services ([DocumentContentExtractor](file:///E:/Python/silvina-editorial/src/domain/document/document_content_extractor.py), [CitationExtractor](file:///E:/Python/silvina-editorial/src/domain/citation/citation_extractor.py), [DocumentFormatInspector](file:///E:/Python/silvina-editorial/src/domain/document/document_format_inspector.py), and [GrammarChecker](file:///E:/Python/silvina-editorial/src/domain/grammar/grammar_checker.py)) to wrap port calls and encapsulate helper logic. Enhance [ApaValidator](file:///E:/Python/silvina-editorial/src/domain/citation/apa_validator.py) and [StructureValidator](file:///E:/Python/silvina-editorial/src/domain/structure/structure_validator.py) to absorb local logic.
   - Pros: Conforms perfectly to hexagonal architecture guidelines, keeps the use case clean, and simplifies testing.
   - Cons: Adds 4 new domain service classes/files.
   - Effort: Medium

2. **Wiring Delegation Only**
   - Extract ports into domain services but keep the helper methods (citation filtering, grammar level lookup, structure filtering) inside [AnalyzeDocumentUseCase](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py).
   - Pros: Avoids modifying existing validator interfaces.
   - Cons: Leaves business logic inside the application layer.
   - Effort: Low

### Recommendation
Approach 1 is recommended as it enforces a clean separation of concerns, keeps the use case focused on orchestration, and removes all business rules from the application layer.

### Risks
- **Adapter Mocking in Tests**: Unit tests must be carefully adjusted to mock the new domain services correctly without losing validation logic coverage.
- **Wired Graph Complexity**: The composition root will manage more domain services, but this is aligned with its responsibility.

### Ready for Proposal
Yes. The orchestrator is ready to create the proposal for this change.
