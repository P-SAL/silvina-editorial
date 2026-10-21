# Proposal: Method Ordering Resolution

## Intent

Address technical debt in domain and infrastructure classes where private methods do not adhere to alphabetical ordering. Ensure consistent code layout across all slices.

## Scope

### In Scope
- Reorder private methods alphabetically in:
  - `src/domain/quality/quality_analyzer.py`
  - `src/domain/quality/quality_response_parser.py`
  - `src/infrastructure/wirings/analyze_document_use_case_wiring.py`
  - `src/infrastructure/wirings/analyze_quality_use_case_wiring.py`
  - `src/infrastructure/wirings/extract_citations_use_case_wiring.py`
  - `src/infrastructure/wirings/extract_content_use_case_wiring.py`

### Out of Scope
- Reordering public methods (already correctly ordered before private methods).
- Modifying file interfaces or business logic.
- Formatting files outside the specified affected list.

## Capabilities

### New Capabilities
None

### Modified Capabilities
None

## Approach

Physically relocate private methods to sort them alphabetically in the class body. Dunder methods remain at the top, and public methods precede private methods. Verify the refactoring by executing the test suite.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/domain/quality/quality_analyzer.py` | Modified | Sort private methods alphabetically (`_ensure_call_produced_usable_content` before `_render_prompt`). |
| `src/domain/quality/quality_response_parser.py` | Modified | Sort private methods alphabetically (`_extract_feedback` before `_extract_score`). |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Modified | Sort private methods alphabetically. |
| `src/infrastructure/wirings/analyze_quality_use_case_wiring.py` | Modified | Sort private methods alphabetically. |
| `src/infrastructure/wirings/extract_citations_use_case_wiring.py` | Modified | Sort private methods alphabetically. |
| `src/infrastructure/wirings/extract_content_use_case_wiring.py` | Modified | Sort private methods alphabetically. |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Method reference / import broken | Low | Python runtime lookup is dynamic. Run test suite to verify. |

## Rollback Plan

Discard local changes using `git checkout -- <filepath>`.

## Dependencies

- None

## Success Criteria

- [x] Private methods in the 6 affected files are sorted alphabetically.
- [x] All unit, integration, and E2E tests pass.
