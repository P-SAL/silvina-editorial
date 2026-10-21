# Tasks: docx-citation-location-fix

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 70-90 lines |
| 400-line budget risk | Low |
| Chained PRs recommended | No |
| Suggested split | Single PR |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main |

Decision needed before apply: No
Chained PRs recommended: No
Chain strategy: stacked-to-main
400-line budget risk: Low

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Adapter paragraph extraction & tests | PR 1 | Modifies docx_citation_adapter.py and adds tests |

## Phase 1: Adapter Refactoring

- [x] 1.1 Update `DocxCitationAdapter.__init__` in [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py) to make `max_author_name_length` optional, defaulting to 100.
- [x] 1.2 Modify `extract_citations` in [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py) to pass `paragraphs` list to `_extract_citations` directly.
- [x] 1.3 Update signature of `_extract_citations` in [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py) to accept `paragraphs: list[str] | str | None = None` and fallback `full_text: str | None = None`.
- [x] 1.4 Refactor `_extract_citations` to loop over `paragraphs` with `p_idx` and run helper methods for each paragraph.
- [x] 1.5 Update helper signatures (`_collect_parenthetical`, `_collect_multi_author`, `_collect_single_author`) to accept `p_text` and `p_idx`.
- [x] 1.6 Update DTO initialization in helpers to set `location=p_idx`.

## Phase 2: Testing

- [x] 2.1 Add unit test in [test_docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py) for Scenario S7e, verifying paragraph index tracking mapping to 0-based index.
- [x] 2.2 Add unit test in [test_docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/tests/adapters/document/test_docx_citation_adapter.py) for Scenario S7f, verifying fallback logic assigning `location=0` for strings and `full_text`.
- [x] 2.3 Verify all tests pass in the adapter's test suite.
