# Design: Method Ordering Resolution

## Technical Approach

Ensure consistent method ordering conventions across domain and infrastructure classes. This change relocates private methods in six affected classes to place them in alphabetical order while preserving dunder methods at the top and public methods preceding private methods.

## Architecture Decisions

### Decision: Relocate Methods Alphabetically

**Choice**: Manually reorder the private methods in the class definitions.
**Alternatives considered**: Automated sorting via custom AST scripts or Ruff configurations.
**Rationale**: With only six affected files, manual rearrangement is extremely low-risk, fast, and does not require writing complex, custom sorting tooling or custom configurations that might have side-effects on other codebases or parts of this codebase.

## Data Flow

No changes to the data flow or business logic are introduced by this change.

## File Changes

| File | Action | Description |
|------|--------|-------------|
| `src/domain/quality/quality_analyzer.py` | Modify | Reorder private methods alphabetically: `_ensure_call_produced_usable_content` before `_render_prompt`. |
| `src/domain/quality/quality_response_parser.py` | Modify | Reorder private methods alphabetically: `_extract_feedback` before `_extract_score` before `_infer_score_from_narrative` before `_map_block_to_dimension`. |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Modify | Reorder private methods alphabetically: `_get_analyze_quality_use_case`, `_get_check_grammar_use_case`, `_get_classify_article_use_case`, `_get_extract_citations_use_case`, `_get_extract_content_use_case`, `_get_match_citations_use_case`, `_get_read_document_use_case`, `_get_recommendation_builder`, `_get_validate_apa_use_case`, `_get_validate_structure_use_case`, `_get_verify_eumic_use_case`. |
| `src/infrastructure/wirings/analyze_quality_use_case_wiring.py` | Modify | Reorder private methods alphabetically: `_get_llm_generator`, `_get_quality_analyzer`, `_get_text_sampler`. |
| `src/infrastructure/wirings/extract_citations_use_case_wiring.py` | Modify | Reorder private methods alphabetically: `_get_citation_port`, `_get_document_text_port`, `_get_reference_port`. |
| `src/infrastructure/wirings/extract_content_use_case_wiring.py` | Modify | Reorder private methods alphabetically: `_get_count_port`, `_get_extraction_port`. |

## Interfaces / Contracts

No changes to external interfaces, API contracts, type definitions, or public methods.

## Testing Strategy

| Layer | What to Test | Approach |
|-------|-------------|----------|
| Unit | Verify method ordering does not break class definition or imports | Execute `pytest` over the suite, particularly target tests in `src/domain/tests` and `src/infrastructure/tests`. |
| Integration | Verify wiring instantiations | Execute `pytest` wiring tests (e.g., `test_analyze_document_use_case_wiring.py`, `test_analyze_quality_use_case_wiring.py`, `test_extract_citations_use_case_wiring.py`, `test_extract_content_use_case_wiring.py`). |
| E2E | End-to-end flow execution | Run the complete test suite using `pytest` to guarantee zero regressions. |

## Migration / Rollout

No migration required.

## Open Questions

None.
