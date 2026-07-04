# Design: Resolve Explicit Keyword Arguments

## Technical Approach

We will perform a comprehensive refactoring of the application and infrastructure layers to enforce the use of explicit keyword arguments (`arg=value`) for all custom function and method calls. This change covers all internal classes, ports, adapters, helper functions, and their respective unit tests.

Standard library calls (e.g., `os.path.join`, `open`), third-party libraries (e.g., `python-docx`, `ollama`), and standard test assertions (e.g., `self.assertEqual`) are out of scope.

## Architecture Decisions

### Decision: Full Audit and Alignment of All Internal and External Calls

**Choice**: Option 2: Full Audit and Alignment.
**Alternatives considered**: Option 1: Boundary/Public calls only.
**Rationale**: Option 2 ensures complete consistency across the codebase. Enforcing the rule on internal/private helper methods (e.g., in `ParagraphContentAdapter` and `DocxReportAdapter`) removes ambiguity and aligns the entire application with the modern style established in recent slices. Option 1 would leave the codebase in a hybrid state, complicating future development and code reviews.

## Data Flow

Data flow remains structurally unchanged. The only change is how parameters are bound at the call sites:

```
[Client / Wiring]
       │
       ▼ (Uses explicit keyword arguments)
[Use Case .execute(param=...)]
       │
       ├─► [Port/Adapter .method(param=...)]
       │
       └─► [Private Helper ._method(param=...)]
```

## File Changes

| File | Action | Description |
|------|--------|-------------|
| [analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) | Modify | Refactor call sites of sub-use cases (`read_document`, `extract_content`, etc.) to use keyword arguments. |
| [analyze_quality_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_quality_use_case.py) | Modify | Refactor call to `self._analyzer.analyze` to use keyword arguments. |
| [classify_article_use_case.py](file:///E:/Python/silvina-editorial/src/application/classify_article_use_case.py) | Modify | Refactor call to `self._classifier.classify` to use keyword arguments. |
| [extract_content_use_case.py](file:///E:/Python/silvina-editorial/src/application/extract_content_use_case.py) | Modify | Refactor calls to `self._extraction_port.extract` and `self._count_port.count` to use keyword arguments. |
| [read_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/read_document_use_case.py) | Modify | Refactor call to `self._port.read_paragraphs` to use keyword arguments. |
| [validate_apa_use_case.py](file:///E:/Python/silvina-editorial/src/application/validate_apa_use_case.py) | Modify | Refactor call to `self._validator.validate_all_citations` to use keyword arguments. |
| [validate_structure_use_case.py](file:///E:/Python/silvina-editorial/src/application/validate_structure_use_case.py) | Modify | Refactor call to `self._validator.validate` to use keyword arguments. |
| [analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py) | Modify | Refactor calls to `read_text_resource` to use keyword arguments. |
| [classify_article_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/classify_article_use_case_wiring.py) | Modify | Refactor call to `read_text_resource` to use keyword arguments. |
| [docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py) | Modify | Refactor calls to `_extract_citations`, `_collect_parenthetical`, `_collect_multi_author`, and `_collect_single_author` to use keyword arguments. |
| [docx_eumic_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_eumic_adapter.py) | Modify | Refactor all internal verification helper calls and static `EumicViolationFactory` calls to use keyword arguments. |
| [docx_reference_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_reference_adapter.py) | Modify | Refactor calls to `_resolve_section_type`, `_parse_references`, and `_clean_reference` to use keyword arguments. |
| [paragraph_content_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/paragraph_content_adapter.py) | Modify | Refactor all internal title, author, abstract, keywords, and section extraction helper calls to use keyword arguments. |
| [win32com_word_count_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/win32com_word_count_adapter.py) | Modify | Refactor calls to `_measure`, `_word_session`, `_word_count`, and `_char_count` to use keyword arguments. |
| [docx_report_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_adapter.py) | Modify | Refactor all internal section additions, cell population helper calls, and score color resolution to use keyword arguments. |
| `src/application/tests/test_*.py` | Modify | Update use case `.execute(...)` invocations to pass arguments via keywords. |
| `src/infrastructure/tests/**/*.py` | Modify | Update adapter/wiring test executions (e.g., `inspect(...)`, `read_paragraphs(...)`, `generate(...)`) to pass arguments via keywords. |

## Interfaces / Contracts

Existing method/function signatures will remain unchanged to prevent backward compatibility issues. We are only changing call sites. No new interfaces or type definitions are introduced.

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | Use Cases & Adapters | Run all unit tests using `pytest` to verify execution flow and ensure refactored call sites function identically. |
| E2E / Smoke | CLI & Gradio app | Run existing E2E/smoke tests to verify end-to-end correctness of the pipeline. |

## Migration / Rollout

No migration required. This is a pure codebase style refactoring.

## Open Questions

None. All target signatures have been audited and parameter names confirmed.
