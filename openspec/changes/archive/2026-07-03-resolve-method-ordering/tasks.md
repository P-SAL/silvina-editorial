# Tasks: Method Ordering Resolution

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 100-130 |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-on-risk |
| Chain strategy | pending |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: pending
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Reorder private methods in all 6 target files | PR 1 | Base branch: main; tests included |

## Phase 1: Domain Refactoring

- [x] 1.1 Reorder private methods in [quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_analyzer.py) to place [_ensure_call_produced_usable_content](file:///E:/Python/silvina-editorial/src/domain/quality/quality_analyzer.py#L87) before [_render_prompt](file:///E:/Python/silvina-editorial/src/domain/quality/quality_analyzer.py#L84).
- [x] 1.2 Reorder private methods in [quality_response_parser.py](file:///E:/Python/silvina-editorial/src/domain/quality/quality_response_parser.py) alphabetically:
  1. [_extract_feedback](file:///E:/Python/silvina-editorial/src/domain/quality/quality_response_parser.py#L86)
  2. [_extract_score](file:///E:/Python/silvina-editorial/src/domain/quality/quality_response_parser.py#L68)
  3. [_infer_score_from_narrative](file:///E:/Python/silvina-editorial/src/domain/quality/quality_response_parser.py#L79)
  4. [_map_block_to_dimension](file:///E:/Python/silvina-editorial/src/domain/quality/quality_response_parser.py#L103)

## Phase 2: Infrastructure Wiring Refactoring

- [x] 2.1 Reorder private methods in [analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py) alphabetically:
  1. [_get_analyze_quality_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L64)
  2. [_get_check_grammar_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L58)
  3. [_get_classify_article_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L61)
  4. [_get_extract_citations_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L52)
  5. [_get_extract_content_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L49)
  6. [_get_match_citations_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L70)
  7. [_get_read_document_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L46)
  8. [_get_recommendation_builder](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L76)
  9. [_get_validate_apa_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L55)
  10. [_get_validate_structure_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L67)
  11. [_get_verify_eumic_use_case](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_document_use_case_wiring.py#L73)
- [x] 2.2 Reorder private methods in [analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py) alphabetically:
  1. [_get_llm_generator](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py#L39)
  2. [_get_quality_analyzer](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py#L25)
  3. [_get_text_sampler](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py#L44)
- [x] 2.3 Reorder private methods in [extract_citations_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_citations_use_case_wiring.py) alphabetically:
  1. [_get_citation_port](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_citations_use_case_wiring.py#L17)
  2. [_get_document_text_port](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_citations_use_case_wiring.py#L23)
  3. [_get_reference_port](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_citations_use_case_wiring.py#L20)
- [x] 2.4 Reorder private methods in [extract_content_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_content_use_case_wiring.py) alphabetically:
  1. [_get_count_port](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_content_use_case_wiring.py#L22)
  2. [_get_extraction_port](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/extract_content_use_case_wiring.py#L19)

## Phase 3: Testing & Verification

- [x] 3.1 Run quality domain unit tests: [test_quality_analyzer.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_analyzer.py) and [test_quality_response_parser.py](file:///E:/Python/silvina-editorial/src/domain/tests/quality/test_quality_response_parser.py).
- [x] 3.2 Run wiring integration tests:
  - [test_analyze_document_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_analyze_document_use_case_wiring.py)
  - [test_analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_analyze_quality_use_case_wiring.py)
  - [test_extract_citations_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_extract_citations_use_case_wiring.py)
  - [test_extract_content_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/test_extract_content_use_case_wiring.py)
- [x] 3.3 Execute full test suite via `pytest` to ensure zero regressions.
