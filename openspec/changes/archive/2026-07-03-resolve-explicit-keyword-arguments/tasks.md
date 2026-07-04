# Tasks: Resolve Explicit Keyword Arguments

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | 300-500 lines |
| 400-line budget risk | Medium |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 (Use Cases & Wirings) -> PR 2 (Adapters) -> PR 3 (Tests) |
| Delivery strategy | ask-on-risk |
| Chain strategy | stacked-to-main (targeting refactor/hexagonal-migration) |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: stacked-to-main
400-line budget risk: Medium

### Suggested Work Units

| Unit | Goal | Likely PR | Notes |
|------|------|-----------|-------|
| 1 | Refactor application layer use cases and wirings | PR 1 | Base branch; verify with existing tests |
| 2 | Refactor infrastructure adapters | PR 2 | Build on PR 1; verify adapter execution |
| 3 | Refactor all unit and integration tests | PR 3 | Build on PR 2; run pytest suite |

## Phase 1: Foundation & Application Layer

- [x] 1.1 Refactor [analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) to pass sub-use case arguments via keywords.
- [x] 1.2 Refactor [analyze_quality_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_quality_use_case.py) and [classify_article_use_case.py](file:///E:/Python/silvina-editorial/src/application/classify_article_use_case.py) call sites.
- [x] 1.3 Refactor [extract_content_use_case.py](file:///E:/Python/silvina-editorial/src/application/extract_content_use_case.py), [read_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/read_document_use_case.py), [validate_apa_use_case.py](file:///E:/Python/silvina-editorial/src/application/validate_apa_use_case.py), and [validate_structure_use_case.py](file:///E:/Python/silvina-editorial/src/application/validate_structure_use_case.py).
- [x] 1.4 Refactor wirings [analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py) and [classify_article_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/classify_article_use_case_wiring.py).

## Phase 2: Adapters

- [x] 2.1 Refactor [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py) citation helper call sites.
- [x] 2.2 Refactor [docx_eumic_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_eumic_adapter.py) internal verification and violation helper call sites.
- [x] 2.3 Refactor [docx_reference_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_reference_adapter.py) parser and resolve helper calls.
- [x] 2.4 Refactor [paragraph_content_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/paragraph_content_adapter.py) extraction helper call sites.
- [x] 2.5 Refactor [win32com_word_count_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/win32com_word_count_adapter.py) session and count call sites.
- [x] 2.6 Refactor [docx_report_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_adapter.py) section and table builder call sites.
- [x] 2.7 Refactor every custom call site (collaborator methods and private helpers) in `src/domain/classification/article_classifier.py` and `src/domain/quality/quality_analyzer.py` to keyword arguments, plus the `OllamaGeneratorAdapter` test (`src/infrastructure/tests/test_ollama_generator_adapter.py`). Not in original design.md scope; added as a follow-up gap closure. Verified exhaustively via full-file review — no remaining positional custom calls in either file.

## Phase 3: Tests & Verification

- [x] 3.1 Refactor use case invocation arguments under `src/application/tests/`.
- [x] 3.2 Refactor adapter and wiring test calls under `src/infrastructure/tests/`.
- [x] 3.3 Run `pytest` and fix any parameter mismatch/signature errors.
