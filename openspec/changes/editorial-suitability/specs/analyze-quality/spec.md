# Delta for analyze-quality

## MODIFIED Requirements

### Requirement: QualityAnalyzer Domain Service Is a Thin Orchestrator

`QualityAnalyzer` (`src/domain/quality/quality_analyzer.py`) MUST be a stateless
domain service, and MUST be the only class defined in that file (one-class-per-file).
Its constructor MUST accept exactly 5 collaborators: an injected `LlmGeneratorPort`
instance, an injected `QualityTextSampler` instance, an injected
`QualityResponseParser` instance, an injected `EditorialSuitabilityAnalyzer` instance,
and the two prompt template strings (`clarity_coherence_prompt_template: str`,
`argumentation_conclusions_prompt_template: str`). It MUST NOT import `ollama`,
anything from `src/infrastructure/`, or perform any file I/O. `analyze()` MUST
delegate text sampling to `QualityTextSampler.build_sample()`, render both prompt
templates via a single private helper that formats an injected template with the
sample text, call `generate()` on the port exactly twice (once per rendered prompt),
delegate parsing of each response to `QualityResponseParser.parse()`, assign dimensions
directly from each call's parsed result, average the 4 final dimension scores into
`overall_score`, map it to a `QualityLevel` via `get_quality_level_from_score()`,
delegate editorial suitability analysis to the `EditorialSuitabilityAnalyzer`,
and return a `QualityResultDTO` containing all scores and the `editorial_suitability` DTO.
It keeps the `QualityAnalysisFailed`-raising validation that checks whether a call
produced any usable content for its relevant dimension pair.
(Previously: QualityAnalyzer constructor took 4 collaborators, and did not perform editorial suitability analysis or include it in its returned QualityResultDTO.)

#### Scenario: Domain service has zero infrastructure imports

- GIVEN the `quality_analyzer.py` source file
- WHEN its import statements are inspected
- THEN none import from `src/infrastructure/` or `ollama`

#### Scenario: Port is called exactly twice per analysis

- GIVEN a fake `LlmGeneratorPort` test double that records call count
- WHEN `QualityAnalyzer.analyze(document_content, article_type)` is invoked once
- THEN the fake port's `generate()` method was called exactly 2 times

#### Scenario: quality_analyzer.py defines exactly one class

- GIVEN the `quality_analyzer.py` source file
- WHEN its top-level class definitions are inspected
- THEN exactly one class, `QualityAnalyzer`, is defined — `_DimensionScore` and
  `_ParsedResponse` no longer live in this file

### Requirement: Overall Score and Quality Level Computation

`QualityAnalyzer` MUST compute `overall_score` as the arithmetic mean of the 4
final dimension scores, and MUST map that mean to a `QualityLevel` using
`get_quality_level_from_score` (`src/domain/enums/quality_level.py`), unchanged:
`>= 9.0` → `EXCELLENT`; `>= 7.0` → `GOOD`; `>= 5.0` → `ACCEPTABLE`; `>= 3.0` →
`NEEDS_IMPROVEMENT`; otherwise `POOR`. The `QualityLevel` enum's body and its
`.value` members (used as printable strings in `QualityResultDTO.__str__`) MUST
NOT be modified by this slice; the 4 numeric thresholds remain expressed as a
named module-level constant used inside `get_quality_level_from_score()`, not as
enum values. The result MUST be returned as a `QualityResultDTO` (updated to support
an optional `editorial_suitability` field), with `dimension_scores` keyed by the 4
dimension string values mirroring the legacy's dict shape.
(Previously: QualityResultDTO did not support the editorial_suitability field.)

#### Scenario: Overall score is the mean of the 4 dimension scores

- GIVEN final dimension scores of `8.0, 6.0, 7.0, 9.0`
- WHEN `overall_score` is computed
- THEN it equals `7.5`

#### Scenario: Quality level boundaries match legacy thresholds

- GIVEN an `overall_score` of exactly `7.0`
- WHEN `quality_level` is computed
- THEN it resolves to `QualityLevel.GOOD`

#### Scenario: QualityLevel enum body is untouched

- GIVEN the `QualityLevel` enum definition
- WHEN this slice is applied
- THEN its 5 members and their string `.value`s are unchanged, and no threshold
  numbers are added as enum values
