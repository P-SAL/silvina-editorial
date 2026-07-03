## Exploration: Explicit Keyword Arguments Audit

### Current State
Currently, multiple files in the application and infrastructure layers (`src/application/`, `src/infrastructure/wirings/`, and `src/infrastructure/adapters/`) violate the codebase convention of using explicit keyword arguments for every method call (even for single-parameter calls). This convention was established starting with Slice 7 (extract-citations) but was never retroactively applied to earlier modules.

A thorough audit has identified several use cases, adapter implementations, internal helper methods, and wiring factory calls that still pass arguments positionally. Standard library calls, third-party libraries, and test assertions are excluded from the scope of this convention.

### Affected Areas
- [src/application/analyze_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_document_use_case.py) — Call sites to other use case `.execute()` methods pass arguments positionally:
  - Line 51: `self._extract_content_use_case.execute` (passes `paragraphs` positionally)
  - Line 61: `self._validate_apa_use_case.execute` (passes `author_year_citations` positionally)
  - Line 63: `self._check_grammar_use_case.execute` (passes `paragraphs` positionally)
  - Line 64: `self._classify_article_use_case.execute` (passes `document_content` positionally)
  - Line 65: `self._analyze_quality_use_case.execute` (passes `document_content` and `article_type` positionally)
  - Line 71: `self._validate_structure_use_case.execute` (passes `document_content` positionally)
  - Line 82: `self._match_citations_use_case.execute` (passes `citations` and `references` positionally)
- [src/application/analyze_quality_use_case.py](file:///E:/Python/silvina-editorial/src/application/analyze_quality_use_case.py) — Calls to `self._analyzer.analyze` pass arguments positionally (Line 11).
- [src/application/classify_article_use_case.py](file:///E:/Python/silvina-editorial/src/application/classify_article_use_case.py) — Calls to `self._classifier.classify` pass `document_content` positionally (Line 11).
- [src/application/extract_content_use_case.py](file:///E:/Python/silvina-editorial/src/application/extract_content_use_case.py) — Passes parameters positionally to `self._extraction_port.extract` (Line 29) and `self._count_port.count` (Line 33).
- [src/application/read_document_use_case.py](file:///E:/Python/silvina-editorial/src/application/read_document_use_case.py) — Passes `path` positionally to `self._port.read_paragraphs` (Line 14).
- [src/application/validate_apa_use_case.py](file:///E:/Python/silvina-editorial/src/application/validate_apa_use_case.py) — Passes `citations` positionally to `self._validator.validate_all_citations` (Line 12).
- [src/application/validate_structure_use_case.py](file:///E:/Python/silvina-editorial/src/application/validate_structure_use_case.py) — Passes parameters positionally to `self._validator.validate` (Line 24).
- [src/infrastructure/wirings/analyze_quality_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/analyze_quality_use_case_wiring.py) — Calls custom utility `read_text_resource` positionally (Lines 30, 33).
- [src/infrastructure/wirings/classify_article_use_case_wiring.py](file:///E:/Python/silvina-editorial/src/infrastructure/wirings/classify_article_use_case_wiring.py) — Calls custom utility `read_text_resource` positionally (Line 41).
- [src/infrastructure/adapters/document/docx_citation_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_citation_adapter.py) — Calls port methods and internal private methods positionally:
  - Line 30: `self._document_text_port.read_paragraphs` (passes `docx_path` positionally)
  - Lines 31, 38-40: Calls to private helper methods (`_extract_citations`, `_collect_parenthetical`, etc.) pass arguments positionally.
- [src/infrastructure/adapters/document/docx_eumic_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_eumic_adapter.py) — A large number of private helper method calls (e.g. `_verify_format`, `_check_margins`) and static calls to `EumicViolationFactory` (Lines 115, 129, 136, etc.) pass arguments positionally.
- [src/infrastructure/adapters/document/docx_reference_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/docx_reference_adapter.py) — Calls port methods and internal helper methods (e.g. `_resolve_section_type`, `_parse_references`, `_clean_reference`) positionally.
- [src/infrastructure/adapters/document/paragraph_content_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/paragraph_content_adapter.py) — Calls internal private helper methods (e.g. `_extract_title`, `_extract_authors`, etc.) positionally.
- [src/infrastructure/adapters/document/win32com_word_count_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/document/win32com_word_count_adapter.py) — Calls internal private methods (e.g. `_measure`, `_word_session`, `_word_count`) positionally.
- [src/infrastructure/adapters/report/docx_report_adapter.py](file:///E:/Python/silvina-editorial/src/infrastructure/adapters/report/docx_report_adapter.py) — Calls internal section builder methods (`_add_header_logo`, `_add_page_numbers`, etc.) and formatting helpers (`_color_for_score`, `_add_markdown_paragraph`) positionally.
- **Test Code in `src/application/tests/`** — Many test execution calls of the use cases are called positionally (e.g., `use_case.execute("test.docx")`). Note that standard unittest assertions (like `self.assertEqual`) should remain positional as they represent the testing framework's standard style.

### Approaches

1. **Strictly Target Boundary/Public Method Calls**
   - *Description*: Apply explicit keyword arguments only to calls crossing layer boundaries (e.g., calls to other use cases, ports, or adapters) and DTO instantiations. Leave internal private helper methods (such as `self._add_header_logo(doc)` in `docx_report_adapter.py`) using positional arguments.
   - *Pros*:
     - Minimizes changes and reduces code churn.
     - Simpler implementation and review.
   - *Cons*:
     - Does not resolve positional parameters for all custom methods, meaning the technical debt registry item remains only partially addressed.
     - Creates inconsistency between public and private methods in the same file.
   - *Effort*: Low

2. **Full Audit and Alignment of All Internal and External Calls**
   - *Description*: Audit and update every call to a user-defined function or class method in the target directories to use keyword arguments. This covers public APIs, ports, adapters, and all internal private helper methods. Built-in functions, third-party library calls (e.g. `python-docx` API), and test suite assertion methods are explicitly excluded.
   - *Pros*:
     - Complete compliance with `TECHNICAL_DEBT.md` Item 1.
     - Ensures consistent design pattern throughout all layers.
     - Highly self-documenting code.
   - *Cons*:
     - High code churn and larger diffs.
     - Can result in very long lines, requiring line wrapping for private method calls with multiple arguments.
   - *Effort*: Medium

### Recommendation
Option 2 is recommended. The codebase convention aims for absolute clarity, and resolving the technical debt item fully ensures future maintainability. Partial fixes will leave the codebase in a hybrid state and complicate future code reviews.

### Risks
- **Runtime Argument Name Mismatch**: If keyword arguments are typed incorrectly (e.g., matching a renamed signature parameter), it will raise a `TypeError` at runtime.
- **Regression in Unaudited Test Files**: Modifying application layer signatures (if necessary to align naming) could break test calls that are still positional. We must ensure signature parameter names match exactly.

### Ready for Proposal
Yes. The orchestrator should proceed to define the design and proposal phase for rewriting the audited call sites using Option 2.
