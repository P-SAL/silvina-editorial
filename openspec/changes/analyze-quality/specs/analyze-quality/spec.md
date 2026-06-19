# LLM-Backed Quality Analysis Specification

## Purpose

LLM-mediated document quality scoring across 4 semantic dimensions (Claridad,
Coherencia, Argumentación, Conclusiones), produced via two sequential prompts to a
text-generation backend reached only through a port. First capability in the
migration with an external dependency — establishes the port/adapter pattern and
naming convention reused by future LLM-calling slices.

## Requirements

### Requirement: LlmGeneratorPort Contract

`LlmGeneratorPort` (`src/domain/ports/llm_generator_port.py`) MUST be an abstract
interface declaring exactly one method, `generate(prompt: str) -> str`. The
signature MUST contain no Ollama-specific types, parameters, or return shapes —
only a plain string in, plain string out.

#### Scenario: Port has no vendor-specific leakage

- GIVEN the `LlmGeneratorPort` interface definition
- WHEN its method signatures are inspected
- THEN the only method is `generate(prompt: str) -> str` and no parameter or
  return type references Ollama or any other concrete vendor

### Requirement: OllamaGeneratorAdapter Implements the Port

`OllamaGeneratorAdapter` (`src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`)
MUST implement `LlmGeneratorPort`, wrapping a real call to `ollama.generate()`.
It MUST be the only file in the slice that imports `ollama`. On success, it MUST
extract and return `response.get('response', '').strip()` from Ollama's response
dict — the domain layer MUST never see that raw dict shape. Its `generate` method
MUST be decorated with `@generic_error_handler`.

#### Scenario: Successful generation returns the stripped response text

- GIVEN a working Ollama backend that returns `{'response': '  some text  '}`
- WHEN `OllamaGeneratorAdapter.generate(prompt)` is called
- THEN it returns `"some text"` as a plain string

#### Scenario: Backend failure raises LanguageModelUnavailable

- GIVEN the underlying `ollama.generate()` call raises a connection error, timeout,
  or any backend exception
- WHEN `OllamaGeneratorAdapter.generate(prompt)` is called
- THEN it raises `LanguageModelUnavailable`

#### Scenario: Adapter is the sole Ollama import site

- GIVEN the full set of new files introduced by this slice
- WHEN their imports are inspected
- THEN only `ollama_generator_adapter.py` imports `ollama`; no domain or
  application file does

### Requirement: QualityDimension Enum

`QualityDimension` (`src/domain/enums/quality_dimension.py`) MUST define exactly 4
members corresponding to the 4 scored dimensions: `CLARIDAD`, `COHERENCIA`,
`ARGUMENTACION`, `CONCLUSIONES`. It MUST NOT reuse or modify the existing
`AnalysisDimension` enum.

#### Scenario: Enum has exactly the 4 expected members

- GIVEN the `QualityDimension` enum
- WHEN its members are enumerated
- THEN there are exactly 4: `CLARIDAD`, `COHERENCIA`, `ARGUMENTACION`,
  `CONCLUSIONES`

#### Scenario: AnalysisDimension is left untouched

- GIVEN the existing `AnalysisDimension` enum
- WHEN this slice is applied
- THEN `AnalysisDimension`'s definition is unchanged and `QualityDimension` does
  not inherit from or alias it

### Requirement: QualityAnalyzer Domain Service Depends Only on the Port

`QualityAnalyzer` (`src/domain/quality/quality_analyzer.py`) MUST be a stateless
domain service whose only collaborator is an injected `LlmGeneratorPort` instance.
It MUST NOT import `ollama` or anything from `src/infrastructure/`. It MUST
construct the two prompts using the legacy's exact Spanish prompt text and
text-sampling logic (see Text Sampling and Prompt Construction below), call
`generate()` on the port exactly twice (once per prompt), parse each response, and
produce a `QualityResultDTO`.

#### Scenario: Domain service has zero infrastructure imports

- GIVEN the `quality_analyzer.py` source file
- WHEN its import statements are inspected
- THEN none import from `src/infrastructure/` or `ollama`

#### Scenario: Port is called exactly twice per analysis

- GIVEN a fake `LlmGeneratorPort` test double that records call count
- WHEN `QualityAnalyzer.analyze(document_content, article_type)` is invoked once
- THEN the fake port's `generate()` method was called exactly 2 times

### Requirement: Text Sampling and Prompt Construction Preserved Verbatim

The domain service MUST reproduce the legacy's text-sampling heuristic exactly:
title + first 3 paragraphs (intro) + 2 middle paragraphs + up to 3 conclusion
paragraphs (detected via case-insensitive `conclusi` regex match, excluding lines
whose first 80 characters contain `http`, `doi.org`, `https`, or `ISBN`) — or, if no
conclusion paragraphs are found, the last 2 non-reference paragraphs. The
assembled sample MUST be truncated to 8000 characters. If the resulting sample has
fewer than 400 words, the service MUST fall back to using the full joined
paragraph text (also truncated to 8000 characters) instead of the sample. The two
prompt templates (Call 1: Claridad + Coherencia; Call 2: Argumentación +
Conclusiones) MUST use the legacy's exact Spanish wording, instructions, response
format headers, and scoring criteria line.

#### Scenario: Short document triggers full-text fallback instead of sampling

- GIVEN a document whose strategically sampled text totals fewer than 400 words
- WHEN the domain service builds the text sample for prompt construction
- THEN it uses the full joined paragraph text (truncated to 8000 characters)
  instead of the sampled excerpt

#### Scenario: Long document uses the strategic sample, not full text

- GIVEN a document whose strategically sampled text totals 400 words or more
- WHEN the domain service builds the text sample for prompt construction
- THEN it uses the sampled excerpt (title + intro + middle + conclusion or
  fallback tail paragraphs), not the full document text

#### Scenario: Conclusion detection excludes reference-like lines

- GIVEN paragraphs after a detected "conclusi..." paragraph where some lines'
  first 80 characters contain `http`, `doi.org`, `https`, or `ISBN`
- WHEN the domain service collects conclusion paragraphs for the sample
- THEN those reference-like lines are excluded from the collected conclusion
  paragraphs

### Requirement: Per-Dimension Response Parsing Preserved Verbatim

The parsing logic (legacy `_parse_llm_response`, ported into the domain service)
MUST preserve the following rules exactly for a single LLM response string:

- Split the response into blocks on dimension headers matching
  `\*\*(?:\d+\.\s*)?(?:Claridad|Coherencia|Argumentaci[oó]n|Conclusiones)` with
  case-insensitive matching, supporting both numbered (`**1. Claridad`) and
  unnumbered (`**Claridad`) header formats.
- For each non-empty block, search for an explicit score using
  `\[Puntuaci[oó]n:\s*(\d+(?:\.\d+)?)(?:/10)?\]` or `(\d+(?:\.\d+)?)\s*/\s*10`
  (case-insensitive). If found, clamp the parsed float to the `[0.0, 10.0]` range.
- If no explicit score is found, infer one from narrative keywords in the
  lowercased block: `excelente`/`sobresaliente`/`muy bueno` → `8.5`;
  `bueno`/`adecuado`/`correcto` → `7.5`; `aceptable`/`suficiente`/`regular` →
  `6.0`; `deficiente`/`débil`/`pobre`/`insuficiente` → `4.0`; otherwise a neutral
  default score.
- Build feedback from all non-empty lines after the header line, joined by
  spaces, with any `**RECOMENDACIÓN...` tail stripped (case-insensitive, across
  newlines) and whitespace collapsed. If the resulting feedback is shorter than
  10 characters, replace it with a named neutral "not available" feedback
  constant. Truncate feedback to at most 3 sentences (split on `.`, rejoined with
  `. ` and a trailing `.`).
- Map each parsed block to a dimension by inspecting the first 200 characters of
  the block (case-insensitive), checking in this order: `argumentaci` →
  `ARGUMENTACION`; else `conclusi` → `CONCLUSIONES`; else `coherencia` →
  `COHERENCIA`; else `claridad` or `argumento` → `CLARIDAD`. (Order matters: the
  `argumentaci` check MUST run before the `claridad`/`argumento` check to avoid a
  false match on the substring `argumento` inside `argumentacion`.)
- Any dimension not matched by any block in the response keeps a named neutral
  default score and feedback for that dimension only.

#### Scenario: Numbered and unnumbered headers both parse correctly

- GIVEN one response using `**1. Claridad del argumento** [Puntuación: 8/10]` and
  another using `**Claridad** [Puntuación: 8/10]`
- WHEN each is parsed
- THEN both produce a Claridad score of `8.0` with the same feedback extraction
  rules applied

#### Scenario: Score inferred from narrative when explicit score is absent

- GIVEN a block whose text contains "el argumento es bastante bueno y adecuado"
  with no `[Puntuación: X/10]` or `X/10` pattern present
- WHEN the block is parsed
- THEN the inferred score is `7.5`

#### Scenario: Feedback shorter than 10 characters becomes the neutral default

- GIVEN a parsed block whose extracted feedback text is fewer than 10 characters
  long
- WHEN parsing completes for that block
- THEN the feedback is replaced with the named neutral "not available" feedback
  constant

#### Scenario: Feedback longer than 3 sentences is truncated

- GIVEN a parsed block whose feedback contains 5 sentences
- WHEN parsing completes for that block
- THEN the stored feedback contains only the first 3 sentences, joined with
  `". "` and ending in `"."`

#### Scenario: Argumentacion is not misclassified as Claridad

- GIVEN a block whose first 200 characters contain both `argumentaci` and the
  word `argumento`
- WHEN the block is mapped to a dimension
- THEN it is mapped to `ARGUMENTACION`, not `CLARIDAD`

#### Scenario: One missing dimension in an otherwise valid response keeps the rest

- GIVEN a Call 1 response with a valid Claridad block but no Coherencia header
  anywhere in the text
- WHEN the response is parsed
- THEN Claridad reflects the parsed score and feedback, and Coherencia falls back
  to the named neutral default score and feedback — parsing for that call does
  not fail

### Requirement: Direct Per-Call Dimension Assignment, No Cross-Call Heuristic

`QualityAnalyzer` MUST assign `CLARIDAD` and `COHERENCIA` directly from Call 1's
parsed result, and `ARGUMENTACION` and `CONCLUSIONES` directly from Call 2's
parsed result. It MUST NOT apply any cross-call fallback or "prefer whichever
call has real feedback" heuristic — each dimension has exactly one authoritative
source call.

#### Scenario: Claridad and Coherencia always come from Call 1

- GIVEN a Call 1 response containing valid Claridad and Coherencia scores and a
  Call 2 response that (hypothetically) also contains text matching a Claridad
  header
- WHEN dimension scores are assembled
- THEN the final Claridad and Coherencia scores are taken from Call 1's parsed
  result, never from Call 2

#### Scenario: Argumentacion and Conclusiones always come from Call 2

- GIVEN a Call 2 response containing valid Argumentacion and Conclusiones scores
- WHEN dimension scores are assembled
- THEN the final Argumentacion and Conclusiones scores are taken from Call 2's
  parsed result, never from Call 1

### Requirement: Full Per-Call Parse Failure Raises QualityAnalysisFailed

If, for a single call's response, all 4 dimensions for that call's relevant pair
fail to parse such that the entire response yields only neutral-default values
with no genuine extracted content (i.e. the response could not be parsed in any
usable way), `QualityAnalyzer` MUST raise `QualityAnalysisFailed` instead of
returning a result built from defaults. Partial failure — where at least one
dimension in a call's response parses successfully — MUST NOT raise; only the
failing dimension(s) fall back to the named neutral default.

#### Scenario: Both dimensions in one call's response fail to parse

- GIVEN Call 1's response text contains neither a Claridad header nor a
  Coherencia header anywhere, and no parseable score/feedback content for either
- WHEN `QualityAnalyzer.analyze(...)` processes Call 1's response
- THEN it raises `QualityAnalysisFailed`

#### Scenario: Partial failure in one call does not raise

- GIVEN Call 1's response contains a valid, parseable Claridad block but no
  Coherencia header
- WHEN `QualityAnalyzer.analyze(...)` processes Call 1's response
- THEN no exception is raised; Claridad uses the parsed value and Coherencia uses
  the named neutral default

### Requirement: Overall Score and Quality Level Computation

`QualityAnalyzer` MUST compute `overall_score` as the arithmetic mean of the 4
final dimension scores, and MUST map that mean to a `QualityLevel` using the same
thresholds as the legacy `get_quality_level_from_score`: `>= 9.0` → `EXCELLENT`;
`>= 7.0` → `GOOD`; `>= 5.0` → `ACCEPTABLE`; `>= 3.0` → `NEEDS_IMPROVEMENT`;
otherwise `POOR`. The result MUST be returned as a `QualityResultDTO` (reused
as-is, no new DTO), with `dimension_scores` keyed by the 4 dimension string values
mirroring the legacy's dict shape.

#### Scenario: Overall score is the mean of the 4 dimension scores

- GIVEN final dimension scores of `8.0, 6.0, 7.0, 9.0`
- WHEN `overall_score` is computed
- THEN it equals `7.5`

#### Scenario: Quality level boundaries match legacy thresholds

- GIVEN an `overall_score` of exactly `7.0`
- WHEN `quality_level` is computed
- THEN it resolves to `QualityLevel.GOOD`

### Requirement: AnalyzeQualityUseCase Thin Pass-Through

`AnalyzeQualityUseCase` (`src/application/analyze_quality_use_case.py`) MUST
expose `execute(document_content: DocumentContentDTO, article_type) ->
QualityResultDTO`, delegating to `QualityAnalyzer` without adding business logic.
The `article_type` parameter MUST remain in the signature even though the
underlying domain service never reads its value — it is not removed in this
slice.

#### Scenario: Use case returns the domain service's result unchanged

- GIVEN a `DocumentContentDTO` and an `article_type` value
- WHEN `AnalyzeQualityUseCase.execute(document_content, article_type)` is called
- THEN the returned `QualityResultDTO` matches what `QualityAnalyzer.analyze`
  would produce for the same `document_content`

#### Scenario: article_type is accepted but not required to affect the result

- GIVEN two calls to `execute()` with the same `document_content` but different
  `article_type` values
- WHEN both calls succeed
- THEN both calls are accepted without error regardless of the `article_type`
  value passed

### Requirement: AnalyzeQualityUseCaseWiring Assembles Domain Service and Adapter

`AnalyzeQualityUseCaseWiring` (`src/infrastructure/wirings/analyze_quality_use_case_wiring.py`)
MUST expose `create_use_case() -> AnalyzeQualityUseCase` as its single
public method, following the instance-based `_get_*`/`_*` accessor pattern from
Slices 2-4. It MUST additionally assemble an `OllamaGeneratorAdapter` instance via
a private method returning the `LlmGeneratorPort` type, and inject it into
`QualityAnalyzer`. This is the first wiring in the migration that assembles a real
infrastructure adapter rather than only domain objects. The wiring MUST NOT
contain business logic.

#### Scenario: Wiring produces a usable use case instance

- GIVEN an `AnalyzeQualityUseCaseWiring` instance
- WHEN `create_use_case()` is called
- THEN it returns an `AnalyzeQualityUseCase` ready to call `.execute(...)`, backed
  by a real `OllamaGeneratorAdapter`

#### Scenario: Adapter accessor returns the port type

- GIVEN the wiring's private adapter-accessor method
- WHEN its return type annotation is inspected
- THEN it is declared as `LlmGeneratorPort`, not `OllamaGeneratorAdapter`

### Requirement: No print() Statements in Migrated Code

None of the new domain, application, or infrastructure files introduced by this
slice MUST contain a `print()` call. The legacy's 5 Spanish emoji progress
messages are dropped, not replaced with logging, in this slice.

#### Scenario: No print calls anywhere in the new code

- GIVEN the full set of new files introduced by this slice (port, adapter, enum,
  domain service, use case, wiring)
- WHEN their source is inspected
- THEN none contain a `print(` call

## Out of Scope

- `self.client = ollama.Client(...)` field from legacy `__init__` — built but
  never used (only `self.ollama.generate()` is called). Not carried into the
  adapter; tracked in `migration/dead-code-registry`.
- Module-level `analyze_document_quality()` convenience function — confirmed
  broken (calls the instance method with 3 positional arguments against a 2-arg
  signature), unreachable, and excluded entirely.
- Wiring `AnalyzeQualityUseCase` into `main.py` — deferred to the
  caller-switchover slice; legacy `business_logic/quality_analyzer.py` keeps
  serving production calls during coexistence.
- Deleting `business_logic/quality_analyzer.py` — coexistence maintained until
  the caller-switchover slice.
- Replacing dropped `print()` progress messages with structured logging — not
  introduced in this slice.
