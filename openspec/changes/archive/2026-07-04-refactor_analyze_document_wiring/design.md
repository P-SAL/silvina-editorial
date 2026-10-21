# Design: Refactor Document Analysis Wiring

## Technical Approach
Eliminate the 10 redundant pass-through use cases and their wirings. `AnalyzeDocumentUseCase` will directly orchestrate the 7 ports, 5 domain services, and 1 builder. `AnalyzeDocumentUseCaseWiring` will instantiate and inject all 13 dependencies directly.

Any intermediate validation logic or fallback handling (such as `CharacterCountUnavailable` fallback and structure validation's references filtering) will be moved into `AnalyzeDocumentUseCase`.

## Architecture Decisions
| Decision Title | Choice | Rationale | Alternatives |
|----------------|--------|-----------|--------------|
| Direct Orchestration | Accept 13 dependencies in `AnalyzeDocumentUseCase` | Reduces boilerplate and layers of useless delegation. | Retain sub-use cases with high architectural complexity. |
| Single LLM Generator | Shared `LlmGeneratorPort` instance in wiring | `ArticleClassifier` and `QualityAnalyzer` use identical LLM configurations; reusing one adapter instance reduces overhead. | Separate generator adapters per domain service. |

## Data Flow
```
[Client / Controller]
        │
        ▼ (document_path)
[AnalyzeDocumentUseCase]
        │───► [DocumentTextPort] ──► (paragraphs)
        │───► [ContentExtractionPort] & [CharacterCountPort] ──► (document_content)
        │───► [CitationExtractionPort] & [ReferenceExtractionPort] ──► (citations, references)
        │───► [ApaValidator] ──► (apa_validation)
        │───► [GrammarCheckPort] ──► (grammar)
        │───► [ArticleClassifier] ──► (classification)
        │───► [QualityAnalyzer] ──► (quality)
        │───► [StructureValidator] ──► (structure)
        │───► [CitationMatcher] ──► (matched_citations)
        │───► [DocumentFormatInspectionPort] ──► (eumic_violations)
        │───► [RecommendationBuilder] ──► (recommendations, verdict)
        ▼
 [ReportInputDTO]
```

## File Changes
| File | Action | Description |
|------|--------|-------------|
| [analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) | Modify | Update constructor to accept 13 dependencies. Merge sub-use case logic (fallbacks/filtering) directly into `execute`. |
| [analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py) | Modify | Update `create_use_case` to wire and inject the 13 dependencies directly. |
| [test_analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_analyze_document_use_case.py) | Modify | Update mocks in test helper to match new constructor and verify direct interactions. |
| [test_analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_analyze_document_use_case_wiring.py) | Modify | Assert direct dependency fields on the instantiated use case class. |

### Deleted Files
#### Application Layer Use Cases (10 files)
- [read_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/read_document_use_case.py)
- [extract_content_use_case.py](file:///E:/Python/silvina-editorial/src/application/extract_content_use_case.py)
- [extract_citations_use_case.py](file:///E:/Python/silvina-editorial/src/application/extract_citations_use_case.py)
- [validate_apa_use_case.py](file:///E:/Python/silvina-editorial/src/application/validate_apa_use_case.py)
- [check_grammar_use_case.py](file:///E:/Python/silvina-editorial/src/application/check_grammar_use_case.py)
- [classify_article_use_case.py](file:///E:/Python/silvina-editorial/src/application/classify_article_use_case.py)
- [analyze_quality_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_quality_use_case.py)
- [validate_structure_use_case.py](file:///E:/Python/silvina-editorial/src/application/validate_structure_use_case.py)
- [match_citations_use_case.py](file:///E:/Python/silvina-editorial/src/application/match_citations_use_case.py)
- [verify_eumic_use_case.py](file:///E:/Python/silvina-editorial/src/application/verify_eumic_use_case.py)

#### Infrastructure Layer Wirings (10 files)
- [read_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/read_document_use_case_wiring.py)
- [extract_content_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_content_use_case_wiring.py)
- [extract_citations_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_citations_use_case_wiring.py)
- [validate_apa_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/validate_apa_wiring.py)
- [check_grammar_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/check_grammar_use_case_wiring.py)
- [classify_article_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/classify_article_use_case_wiring.py)
- [analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py)
- [validate_structure_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/validate_structure_wiring.py)
- [match_citations_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/match_citations_use_case_wiring.py)
- [verify_eumic_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/verify_eumic_use_case_wiring.py)

#### Application Use Case Tests (10 files)
- [test_read_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_read_document_use_case.py)
- [test_extract_content_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_extract_content_use_case.py)
- [test_extract_citations_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_extract_citations_use_case.py)
- [test_validate_apa_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_validate_apa_use_case.py)
- [test_check_grammar_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_check_grammar_use_case.py)
- [test_classify_article_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_classify_article_use_case.py)
- [test_analyze_quality_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_analyze_quality_use_case.py)
- [test_validate_structure_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_validate_structure_use_case.py)
- [test_match_citations_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_match_citations_use_case.py)
- [test_verify_eumic_use_case.py](file:///E:/Python/silvina-editorial/src/application/tests/test_verify_eumic_use_case.py)

#### Infrastructure Wiring Tests (10 files)
- [test_read_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_read_document_use_case_wiring.py)
- [test_extract_content_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_extract_content_use_case_wiring.py)
- [test_extract_citations_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_extract_citations_use_case_wiring.py)
- [test_validate_apa_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_validate_apa_wiring.py)
- [test_check_grammar_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_check_grammar_use_case_wiring.py)
- [test_classify_article_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_classify_article_use_case_wiring.py)
- [test_analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_analyze_quality_use_case_wiring.py)
- [test_validate_structure_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_validate_structure_wiring.py)
- [test_match_citations_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_match_citations_use_case_wiring.py)
- [test_verify_eumic_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_verify_eumic_use_case_wiring.py)

## Interfaces / Contracts

### AnalyzeDocumentUseCase Signature
```python
def __init__(
    self,
    document_text_port: DocumentTextPort,
    content_extraction_port: ContentExtractionPort,
    character_count_port: CharacterCountPort,
    citation_extraction_port: CitationExtractionPort,
    reference_extraction_port: ReferenceExtractionPort,
    grammar_check_port: GrammarCheckPort,
    document_format_inspection_port: DocumentFormatInspectionPort,
    apa_validator: ApaValidator,
    article_classifier: ArticleClassifier,
    quality_analyzer: QualityAnalyzer,
    structure_validator: StructureValidator,
    citation_matcher: CitationMatcher,
    recommendation_builder: RecommendationBuilder,
) -> None:
```

### AnalyzeDocumentUseCaseWiring Interface
```python
class AnalyzeDocumentUseCaseWiring:
    def create_use_case(self) -> AnalyzeDocumentUseCase: ...
    def _get_document_text_port(self) -> DocumentTextPort: ...
    def _get_content_extraction_port(self) -> ContentExtractionPort: ...
    def _get_character_count_port(self) -> CharacterCountPort: ...
    def _get_citation_extraction_port(self) -> CitationExtractionPort: ...
    def _get_reference_extraction_port(self) -> ReferenceExtractionPort: ...
    def _get_grammar_check_port(self) -> GrammarCheckPort: ...
    def _get_document_format_inspection_port(self) -> DocumentFormatInspectionPort: ...
    def _get_apa_validator(self) -> ApaValidator: ...
    def _get_article_classifier(self) -> ArticleClassifier: ...
    def _get_quality_analyzer(self) -> QualityAnalyzer: ...
    def _get_structure_validator(self) -> StructureValidator: ...
    def _get_citation_matcher(self) -> CitationMatcher: ...
    def _get_recommendation_builder(self) -> RecommendationBuilder: ...

    # Internal wiring helpers
    def _get_llm_generator(self) -> LlmGeneratorPort: ...
    def _get_article_size_classifier(self) -> ArticleSizeClassifier: ...
    def _get_article_size_thresholds(self) -> ArticleSizeThresholdsDTO: ...
    def _get_quality_level_resolver(self) -> QualityLevelResolver: ...
    def _get_quality_level_thresholds(self) -> QualityLevelThresholdsDTO: ...
    def _get_quality_text_sampler(self) -> QualityTextSampler: ...
```

## Testing Strategy
| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit / Application | Mocks check in `TestAnalyzeDocumentUseCase` | Update to inject and verify invocation parameters on mock ports, services, and builder. |
| Integration / Infrastructure | Assembly in `TestAnalyzeDocumentUseCaseWiring` | Assert that wiring resolves to `AnalyzeDocumentUseCase` and internal fields hold correctly configured adapters. |

## Migration / Rollout
No migration required.

## Open Questions
- None
