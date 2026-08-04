## Exploration: resolve-method-ordering

### Current State
Currently, several classes written in earlier slices (Slices 0, 5, 6) do not adhere to the method-ordering convention defined in later reviews. This convention specifies that:
1. Public methods must come before private methods.
2. No interleaving of public and private methods is allowed.
3. Methods within each group (public, private) must be ordered alphabetically.
4. Dunder methods (e.g., `__init__`) are exempt and placed at the top of the class.

A codebase-wide audit of the candidate files revealed that 6 files violate the alphabetical ordering rule for private methods. No files violated the "public before private" or "no interleaving" rules.

### Affected Areas
- `src/domain/quality/quality_analyzer.py` — `QualityAnalyzer` private methods are out of alphabetical order: `_render_prompt` (line 84) is defined before `_ensure_call_produced_usable_content` (line 87).
- `src/domain/quality/quality_response_parser.py` — `QualityResponseParser` private methods are out of alphabetical order: `_extract_score` (line 68) -> `_infer_score_from_narrative` (line 79) -> `_extract_feedback` (line 86) -> `_map_block_to_dimension` (line 103). The correct alphabetical order is `_extract_feedback`, `_extract_score`, `_infer_score_from_narrative`, `_map_block_to_dimension`.
- `src/infrastructure/wirings/analyze_document_use_case_wiring.py` — `AnalyzeDocumentUseCaseWiring` private methods are out of alphabetical order. Correct alphabetical order: `_get_analyze_quality_use_case`, `_get_check_grammar_use_case`, `_get_classify_article_use_case`, `_get_extract_citations_use_case`, `_get_extract_content_use_case`, `_get_match_citations_use_case`, `_get_read_document_use_case`, `_get_recommendation_builder`, `_get_validate_apa_use_case`, `_get_validate_structure_use_case`, `_get_verify_eumic_use_case`.
- `src/infrastructure/wirings/analyze_quality_use_case_wiring.py` — `AnalyzeQualityUseCaseWiring` private methods are out of alphabetical order. Correct alphabetical order: `_get_llm_generator`, `_get_quality_analyzer`, `_get_text_sampler`.
- `src/infrastructure/wirings/extract_citations_use_case_wiring.py` — `ExtractCitationsUseCaseWiring` private methods are out of alphabetical order. Correct alphabetical order: `_get_citation_port`, `_get_document_text_port`, `_get_reference_port`.
- `src/infrastructure/wirings/extract_content_use_case_wiring.py` — `ExtractContentUseCaseWiring` private methods are out of alphabetical order. Correct alphabetical order: `_get_count_port`, `_get_extraction_port`.

*Note: The following candidate files were audited and found to be compliant (no violations or single-method classes):*
- `src/domain/quality/quality_text_sampler.py`
- `src/domain/classification/article_classification_response_parser.py`
- `src/domain/classification/article_classification_text_sampler.py`
- `src/domain/classification/imryd_signal_detector.py`
- `src/domain/classification/article_size_classifier.py`
- `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`
- `src/infrastructure/wirings/check_grammar_use_case_wiring.py`
- `src/infrastructure/wirings/classify_article_use_case_wiring.py`
- `src/infrastructure/wirings/export_report_wiring.py`
- `src/infrastructure/wirings/match_citations_use_case_wiring.py`
- `src/infrastructure/wirings/read_document_use_case_wiring.py`
- `src/infrastructure/wirings/validate_apa_wiring.py`
- `src/infrastructure/wirings/validate_structure_wiring.py`
- `src/infrastructure/wirings/verify_eumic_use_case_wiring.py`

### Approaches
1. **Reorder Methods Alphabetically** — Physically move the method definitions inside each affected class to match the required alphabetical order for their respective groups (mostly private methods).
   - Pros: Fully resolves the convention debt without altering functionality; simple and low risk.
   - Cons: Moves lines around which might disrupt git history blame for those lines (minimal issue).
   - Effort: Low

2. **Automated Sorting via Tooling (e.g., Ruff/Custom script)** — Configure or build a script to automatically sort methods according to the convention.
   - Pros: Repeatable, handles future violations.
   - Cons: Hard to configure custom Python sorting rules (like alphabetical sorting specifically for custom class private methods while exempting dunders and keeping public first) with off-the-shelf tools without writing a custom AST parser script. Overkill for 6 files.
   - Effort: Medium

### Recommendation
Approach 1 (Reorder Methods Alphabetically) is recommended. Since only 6 files are affected, a manual rearrangement is safe, quick, and can be verified by running the existing unit test suite to ensure no syntax/import errors were introduced.

### Risks
- Rearranging methods can lead to reference issues if a class-level variable definition or decorator depends on method ordering, but in Python class methods are bound at class definition time, so reordering standard methods does not affect execution.
- Potential merge conflicts if other feature branches are editing these files simultaneously.

### Ready for Proposal
Yes — The orchestrator should proceed to the proposal phase (`proposal.md`) detailing the specific reordering changes to be applied to the 6 identified files.
