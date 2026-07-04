# Proposal: analyze-quality (Slice 5)

## Intent

`business_logic/quality_analyzer.py` (`QualityAnalyzer`, 247 lines) scores document quality
across 4 dimensions (Claridad, Coherencia, Argumentación, Conclusiones) by making 2 sequential
`ollama.generate()` calls with fixed prompts, then parsing the LLM's free-text response into
scores via regex. This is the **first slice requiring real I/O** — Slices 2-4 were pure
computation. The migration must introduce the project's first port/adapter pair, set the
naming precedent for it, and fix the legacy's silent-failure behavior (which swallows any
LLM error and returns a fabricated 7.0/ACCEPTABLE result, hiding real failures from `main.py`).

## Scope

### In Scope

- `src/domain/ports/llm_generator_port.py` — `LlmGeneratorPort`, single method
  `generate(prompt: str) -> str`. **Generic name by design**: a future slice
  (`ArticleClassifier`) also calls Ollama and will reuse this exact port. This establishes the
  port-naming convention for the rest of the migration — ports are named after the
  capability they expose (text generation), not the consuming domain (quality) or the
  concrete vendor (Ollama).
- `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py` —
  `OllamaGeneratorAdapter(LlmGeneratorPort)`, wraps `ollama.generate()`/`ollama.Client`,
  decorated with `@generic_error_handler`. Raises `LanguageModelUnavailable` when the Ollama
  call itself fails (connection error, timeout, backend exception).
- New enum `src/domain/enums/quality_dimension.py` for the 4 actual dimensions (Claridad,
  Coherencia, Argumentación, Conclusiones). Does **not** reuse `AnalysisDimension`
  (8 unrelated English dims — confirmed dead/aspirational, left untouched).
- `src/domain/quality/quality_analyzer.py` — stateless domain service: prompt construction
  (2 templates, verbatim text-sampling logic), response parsing, per-dimension merge, overall
  score averaging, `quality_level` mapping via `get_quality_level_from_score` (ported from
  `domain/enums.py`, not yet in `src/`). Depends only on `LlmGeneratorPort`, never on `ollama`.
  Calls `generate()` twice (once per prompt) — the port stays minimal; the domain owns
  sequencing.
- `src/application/analyze_quality_use_case.py` — `AnalyzeQualityUseCase.execute(document_content: DocumentContentDTO, article_type) -> QualityResultDTO`. Thin pass-through to
  the domain service.
- `src/infrastructure/wirings/analyze_quality_use_case_wiring.py` —
  `AnalyzeQualityUseCaseWiring`, instance-based `_get_*` pattern, now also wiring the adapter
  (`_llm_generator_port()` returns `OllamaGeneratorAdapter`).
- Reuse `QualityResultDTO` and `QualityLevel` as-is — no new result DTO needed, both already
  fit 1:1.
- Domain tests with a fake `LlmGeneratorPort` test double (no real Ollama calls); adapter
  tests mock the `ollama` client to verify the `LanguageModelUnavailable` mapping.

### Out of Scope (dead code — tracked in `migration/dead-code-registry`)

- `self.client = ollama.Client(...)` field in legacy `__init__` — built but never used (only
  `self.ollama.generate()` is called). Not carried into the port/adapter design.
- Module-level `analyze_document_quality()` convenience function — broken (calls the instance
  method with 3 args against a 2-arg signature), unreachable, confirmed dead.
- 5 `print()` calls (Spanish emoji progress messages) — dropped per clean-architecture rule
  (no `print()` in domain/service code), not replaced with logging in this slice.
- `article_type` parameter on `analyze_quality()` / `AnalyzeQualityUseCase.execute()` — kept
  in the signature even though the method body never reads it (legacy never reads it either;
  `main.py` passes the full classification result here). Per user decision, dead parameters
  are not cleaned up slice-by-slice; recorded in `migration/dead-code-registry` for a future
  cross-slice cleanup pass instead of silent removal.
- Deleting `business_logic/quality_analyzer.py` — coexistence maintained until the caller
  switchover slice.
- Wiring `AnalyzeQualityUseCase` into `main.py` — deferred to the caller-switchover slice.

## Capabilities

### New Capabilities

- `analyze-quality`: LLM-backed document quality scoring across 4 dimensions, exposed via a
  use case returning `QualityResultDTO`. First capability in the migration with an external
  dependency, mediated through `LlmGeneratorPort`.

### Modified Capabilities

None.

## Approach — Why This Slice Differs Architecturally

Slices 2-4 (`validate-structure`, `validate-apa`, `validate-citations`) were pure functions:
input DTOs in, output DTOs out, zero I/O, zero coexistence risk beyond logic correctness. This
slice calls an external LLM, which the domain layer must never do directly (Import Invariants,
clean-architecture skill). The fix is the hexagonal port/adapter split:

1. **Port** (`src/domain/ports/llm_generator_port.py`) — the domain declares the *capability*
   it needs (`generate(prompt) -> str`), nothing about Ollama leaks into the interface.
2. **Adapter** (`src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py`) — the
   only place `ollama` is imported. Wraps the real call, extracts `.get('response', '').strip()`
   from Ollama's response dict (the domain never touches that shape), and is decorated with
   `@generic_error_handler` so unexpected exceptions are wrapped/logged consistently.
3. **Domain service** (`src/domain/quality/quality_analyzer.py`) — receives the port via
   constructor injection, contains all current business logic verbatim (text sampling, prompt
   templates, parsing, merge, scoring) with zero knowledge of Ollama.
4. **Use case** — thin orchestration, same shape as Slices 2-4.
5. **Wiring** — now assembles a real adapter instance, not just domain objects; this is the
   precedent for every future slice needing external I/O.

### Error Mapping (two distinct exceptions, both reused from Slice 1)

| Failure point | Exception raised | Rationale |
|---|---|---|
| `OllamaGeneratorAdapter.generate()` — Ollama unreachable, connection error, backend timeout | `LanguageModelUnavailable` (`language_model_errors.py`) | Infrastructure-level failure: the LLM backend itself could not be reached or errored. Raised at the adapter boundary via `@generic_error_handler`. |
| `quality_analyzer.py` — LLM responded but parsing/scoring cannot produce a usable result (e.g. response is empty or unparseable) | `QualityAnalysisFailed` (`quality_errors.py`) | Domain-level failure: the call succeeded but the *quality analysis itself* could not be completed. Raised by the domain service, not the adapter. |

This replaces the legacy's silent `except Exception: return fake 7.0/ACCEPTABLE`. Both
exceptions already exist from Slice 1 (`domain-exceptions`) with correct base classes
(`LanguageModelError(BaseSrcError)`, `QualityError(BaseSrcError)`) — no new exception classes
are created in this slice.

## Affected Areas

| Area | Impact | Description |
|------|--------|--------------|
| `src/domain/ports/llm_generator_port.py` | New | `LlmGeneratorPort` — first port in the migration |
| `src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py` | New | `OllamaGeneratorAdapter` — first adapter |
| `src/domain/enums/quality_dimension.py` | New | 4 actual dimensions (Claridad/Coherencia/Argumentación/Conclusiones) |
| `src/domain/quality/quality_analyzer.py` | New | Stateless domain service, depends on the port |
| `src/application/analyze_quality_use_case.py` | New | `AnalyzeQualityUseCase` |
| `src/infrastructure/wirings/analyze_quality_use_case_wiring.py` | New | Wiring, now assembling an adapter too |
| `src/domain/tests/quality/test_quality_analyzer.py` | New | Domain tests with fake port |
| `src/infrastructure/tests/test_ollama_generator_adapter.py` | New | Adapter tests, mocked `ollama` client |
| `business_logic/quality_analyzer.py` | Unchanged | Legacy stays alive during coexistence |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Port name `LlmGeneratorPort` sets precedent for all future LLM-calling slices | High (by design) | Documented explicitly here as the convention; future slices (e.g. `ArticleClassifier`) reuse it rather than inventing a new port |
| Raising exceptions instead of silent fallback is a behavior change `main.py` doesn't yet handle | Med | Deferred: caller switchover slice must add `try/except BaseSrcError` handling before wiring this use case live |
| Text-sampling/prompt-template logic is intricate (intro+middle+conclusion regex heuristics) | Med | Copy verbatim into the domain service; cover with tests mirroring legacy fixtures |
| New `QualityDimension` enum values (string keys) must match exactly what the parser regexes expect | Low | Single source of truth enum, parser tested against all 4 values |

## Rollback Plan

All new files are additive. Legacy `business_logic/quality_analyzer.py` is untouched. To roll
back: delete the new port, adapter, enum, domain service, use case, wiring, and test files.
`main.py` continues importing from `business_logic/`. No migration state to undo.

## Dependencies

- Slices 2-4 — establish the enum -> DTO -> domain service -> use case -> wiring pattern.
- Slice 1 (`domain-exceptions`) — `LanguageModelUnavailable`, `QualityAnalysisFailed` already
  exist and are reused as-is.
- `QualityResultDTO`, `QualityLevel` — already exist, reused as-is.
- `generic_error_handler` decorator — already exists, applied to the new adapter.

## Success Criteria

- [ ] `LlmGeneratorPort` exists with exactly one method, `generate(prompt: str) -> str`, and no
      Ollama-specific types in its signature
- [ ] `OllamaGeneratorAdapter` is the only file in the slice importing `ollama`
- [ ] Adapter raises `LanguageModelUnavailable` on backend failure; domain service raises
      `QualityAnalysisFailed` on post-response parsing/scoring failure — no silent fallback
- [ ] `quality_analyzer.py` domain service has zero imports from `infrastructure/` or `ollama`
- [ ] New `QualityDimension` enum has exactly 4 values matching Claridad/Coherencia/
      Argumentación/Conclusiones; `AnalysisDimension` is untouched
- [ ] No `print()` calls anywhere in the new domain/application code
- [ ] `article_type` parameter is present in `AnalyzeQualityUseCase.execute()` but documented
      as unused, with a registry entry in `migration/dead-code-registry`
- [ ] Legacy `business_logic/quality_analyzer.py` is unmodified; `main.py` still imports from
      `business_logic/`

## Open Questions

1. **Exact prompt-merge semantics** — legacy merges call-1/call-2 results per dimension,
   preferring call-1 unless its feedback is "No disponible", falling back to call-2. Since
   Claridad/Coherencia only ever appear in call-1's response and Argumentación/Conclusiones
   only in call-2's, the cross-check is currently a no-op in practice. Design phase must decide:
   preserve the (redundant) merge logic verbatim for behavioral fidelity, or simplify to direct
   per-dimension lookup since the cross-check never actually triggers. Either is safe; this is
   a design-level implementation detail, not a scope question.
2. **Fallback score/feedback constants** (`7.0`, "No disponible", "Análisis no disponible") —
   confirm in design whether these become named constants in the domain service or are removed
   entirely now that exceptions replace the silent fallback path.
