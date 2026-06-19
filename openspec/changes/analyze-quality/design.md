# Design: analyze-quality (Slice 5)

## Module Layout

First slice to introduce `src/domain/ports/` and `src/infrastructure/adapters/` — both
new top-level folders in the migration.

```
src/
├── domain/
│   ├── ports/                                   # NEW — first port folder in the migration
│   │   └── llm_generator_port.py                # LlmGeneratorPort (Protocol)
│   ├── enums/
│   │   ├── quality_dimension.py                 # NEW — CLARIDAD/COHERENCIA/ARGUMENTACION/CONCLUSIONES
│   │   └── quality_level.py                     # MODIFIED — add get_quality_level_from_score()
│   ├── quality/                                  # NEW — entity folder (mirrors structure/citation)
│   │   └── quality_analyzer.py                  # QualityAnalyzer domain service
│   ├── dtos/
│   │   └── quality_result_dto.py                # UNCHANGED — reused as-is
│   └── tests/
│       └── quality/
│           └── test_quality_analyzer.py         # NEW — fake LlmGeneratorPort double
├── application/
│   └── analyze_quality_use_case.py              # NEW — AnalyzeQualityUseCase
└── infrastructure/
    ├── adapters/                                 # NEW — first adapter folder in the migration
    │   └── llm_generator/
    │       └── ollama_generator_adapter.py      # OllamaGeneratorAdapter(LlmGeneratorPort)
    ├── wirings/
    │   └── analyze_quality_use_case_wiring.py   # NEW — assembles adapter + domain service
    └── tests/
        └── test_ollama_generator_adapter.py     # NEW — mocked ollama client
```

`business_logic/quality_analyzer.py` stays untouched (coexistence).

---

## ADR-1: Port Abstraction — `Protocol` over `abc.ABC`

**Decision**: `LlmGeneratorPort` is a `typing.Protocol`, not an `abc.ABC` subclass.

**Rationale**: Grepped `src/domain/` and `src/infrastructure/` for `abc`, `Protocol`,
`abstractmethod` — zero hits. No existing precedent either way. Per the proposal's own
guidance, default to `Protocol` for a pure single-method interface: no inheritance coupling
required from the adapter side (`OllamaGeneratorAdapter` satisfies the protocol structurally),
no `ABCMeta` metaclass machinery needed, and it keeps the port file dependency-free
(`typing` is stdlib, already used everywhere in this codebase per the PEP 604 rule).

**Rejected alternative**: `abc.ABC` + `@abstractmethod`. Would work identically at the call
site but adds an explicit `class OllamaGeneratorAdapter(LlmGeneratorPort, ABC)` inheritance
requirement. `Protocol` is lighter and equally clear for a one-method interface; this becomes
the project's convention for ports going forward (worth revisiting only if a future port needs
shared default-method behavior, which favors ABC instead).

```python
# src/domain/ports/llm_generator_port.py
from typing import Protocol


class LlmGeneratorPort(Protocol):
    """Capability to generate text from a prompt via a language model backend."""

    def generate(self, prompt: str) -> str:
        """Return the generated text for the given prompt."""
        ...
```

---

## ADR-2: Adapter Wraps `ollama.generate()` Directly, No `Client` Field

**Decision**: `OllamaGeneratorAdapter` calls module-level `ollama.generate(...)` exactly like
legacy's `self.ollama.generate(...)`. Does **not** carry over the unused `self.client =
ollama.Client(host=...)` field — confirmed dead in the proposal's Out-of-Scope section.

**Rejected alternative**: Porting the `Client` field "for completeness." Rejected — it was
never called in 247 lines of legacy code; carrying dead fields into a fresh port/adapter
defeats the purpose of cleanup-via-migration. Tracked in `migration/dead-code-registry`
already; no need to resurrect it here.

```python
# src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py
import ollama

from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.exceptions.language_model_errors import LanguageModelUnavailable
from src.domain.ports.llm_generator_port import LlmGeneratorPort

_MODEL_NAME = "llama3-gradient:8b-instruct-1048k-q4_K_M"
_GENERATION_OPTIONS = {
    "temperature": 0.2,
    "num_predict": 1000,
    "num_ctx": 4096,
    "repeat_penalty": 1.1,
    "timeout": 120,
}


class OllamaGeneratorAdapter(LlmGeneratorPort):
    """Adapter that generates text via a local Ollama backend."""

    def __init__(self, model_name: str = _MODEL_NAME) -> None:
        self._model_name = model_name

    @generic_error_handler
    def generate(self, prompt: str) -> str:
        """Return the stripped response text from Ollama for the given prompt."""
        try:
            response = ollama.generate(
                model=self._model_name,
                prompt=prompt,
                options=_GENERATION_OPTIONS,
            )
        except Exception as exc:
            raise LanguageModelUnavailable() from exc
        return response.get("response", "").strip()
```

Note: `@generic_error_handler` wraps *unexpected* exceptions into `SrcGenericError`; the
explicit `try/except Exception → LanguageModelUnavailable` inside `generate()` runs first and
takes priority, satisfying the spec's exact requirement ("backend failure raises
`LanguageModelUnavailable`", not a generic wrapper). The decorator still adds logging/no-op
re-raise behavior for `LanguageModelUnavailable` itself (it's a `BaseSrcError` subclass, hits
the `except BaseSrcError` branch, gets logged once, re-raised unchanged).

`OllamaGeneratorAdapter(LlmGeneratorPort)` — explicit inheritance from the `Protocol` is
optional in Python (structural typing satisfies it either way), but it is written explicitly
here for readability and to make the implemented-port relationship grep-able, matching the
existing codebase pattern of explicit relationships (e.g. exception hierarchies).

---

## ADR-3: Fallback Granularity — Single Named Constant, Not Three Duplicated Literals

**Decision**: One module-level constant in `quality_analyzer.py`:

```python
_UNSCORED_DIMENSION_SCORE = 7.0
_UNSCORED_DIMENSION_FEEDBACK = "No disponible"
```

used both for (a) a single dimension missing from an otherwise-parseable response, and (b) as
the seed/default values the parser starts from before overwriting matched blocks. This
replaces legacy's 3 duplicated `{"score": 7.0, "feedback": "No disponible"}` literals
(initial `result` dict, `s1` get-default, `s2` get-default) with one source of truth.

**Rejected alternative**: legacy's separate "Análisis no disponible" feedback string used only
in the `except Exception` fallback path. That entire fallback path is removed — the spec
mandates raising `QualityAnalysisFailed` instead of returning a fabricated result, so that
string has no caller left and is dropped, not ported.

---

## ADR-4: Direct Per-Call Assignment, Cross-Call Merge Removed

**Decision**: Per proposal Open Question 1 — simplify to direct per-dimension lookup.
Claridad/Coherencia are read directly from `parse_response(text_1)`; Argumentacion/Conclusiones
directly from `parse_response(text_2)`. The legacy `s1 if s1["feedback"] != "No disponible"
else s2` merge loop is deleted entirely.

**Rationale**: confirmed in the proposal as behaviorally a no-op (Claridad/Coherencia headers
never appear in Call 2's prompt template, and vice versa for Argumentacion/Conclusiones — the
two prompts ask for disjoint dimension pairs). Removing it loses no behavior and removes a
4-line loop that could never branch the way its name implies.

---

## ADR-5: Full-Call Parse Failure Detection

**Decision**: `_parse_response(text: str) -> dict[QualityDimension, DimensionScore]` always
returns all 4 dimension keys (2 real + 2 untouched defaults, structurally), but
`QualityAnalyzer.analyze()` only inspects the **2 dimensions relevant to that call** when
deciding whether to raise. A call's response "fully fails" when *neither* of its two relevant
headers matched, i.e. both relevant entries are still the sentinel default after parsing.

Implementation approach: `_parse_response` returns a dict keyed by all 4 `QualityDimension`
members PLUS a `set[QualityDimension]` of which dimensions were genuinely matched (not just
left as the default). `QualityAnalyzer.analyze()` checks, per call, whether at least one of its
2 relevant dimensions is in the matched set; if neither is, raise `QualityAnalysisFailed`.

```python
@dataclass(frozen=True)
class _ParsedResponse:
    scores: dict[QualityDimension, "_DimensionScore"]
    matched_dimensions: frozenset[QualityDimension]
```

This is an internal/private structure inside `quality_analyzer.py` — not exported, not a DTO
(no `BaseDTO` needed for an implementation-internal value object).

---

## `QualityAnalyzer` Domain Service — Full Design

```python
# src/domain/quality/quality_analyzer.py
import re
from dataclasses import dataclass

from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.enums.quality_dimension import QualityDimension
from src.domain.enums.quality_level import get_quality_level_from_score
from src.domain.exceptions.quality_errors import QualityAnalysisFailed
from src.domain.ports.llm_generator_port import LlmGeneratorPort

_UNSCORED_DIMENSION_SCORE = 7.0
_UNSCORED_DIMENSION_FEEDBACK = "No disponible"
_MINIMUM_SAMPLE_WORD_COUNT = 400
_TEXT_SAMPLE_CHARACTER_LIMIT = 8000
_REFERENCE_LINE_MARKERS = ("http", "doi.org", "https", "ISBN")
_DIMENSION_HEADER_PATTERN = re.compile(
    r"(?=\*\*(?:\d+\.\s*)?(?:Claridad|Coherencia|Argumentaci[oó]n|Conclusiones))",
    re.IGNORECASE,
)
_EXPLICIT_SCORE_PATTERN = re.compile(
    r"\[Puntuaci[oó]n:\s*(\d+(?:\.\d+)?)(?:/10)?\]|(\d+(?:\.\d+)?)\s*/\s*10",
    re.IGNORECASE,
)
_RECOMMENDATION_TAIL_PATTERN = re.compile(r"\*\*RECOMENDACIÓN.*", re.DOTALL | re.IGNORECASE)
_NARRATIVE_SCORE_KEYWORDS = (
    (("excelente", "sobresaliente", "muy bueno"), 8.5),
    (("bueno", "adecuado", "correcto"), 7.5),
    (("aceptable", "suficiente", "regular"), 6.0),
    (("deficiente", "débil", "pobre", "insuficiente"), 4.0),
)


@dataclass(frozen=True)
class _DimensionScore:
    score: float
    feedback: str


@dataclass(frozen=True)
class _ParsedResponse:
    scores: dict[QualityDimension, _DimensionScore]
    matched_dimensions: frozenset[QualityDimension]


class QualityAnalyzer:
    """Domain service that scores document quality across 4 dimensions via an LLM."""

    def __init__(self, llm_generator: LlmGeneratorPort) -> None:
        self._llm_generator = llm_generator

    def analyze(
        self, document_content: DocumentContentDTO, article_type
    ) -> QualityResultDTO:
        """Score document quality across Claridad, Coherencia, Argumentación and Conclusiones."""
        text_sample = self._build_text_sample(document_content)

        prompt_1 = self._build_prompt_one(text_sample)
        prompt_2 = self._build_prompt_two(text_sample)

        response_1 = self._llm_generator.generate(prompt_1)
        response_2 = self._llm_generator.generate(prompt_2)

        parsed_1 = self._parse_response(response_1)
        self._ensure_call_produced_usable_content(
            parsed_1, relevant_dimensions=(QualityDimension.CLARIDAD, QualityDimension.COHERENCIA)
        )

        parsed_2 = self._parse_response(response_2)
        self._ensure_call_produced_usable_content(
            parsed_2,
            relevant_dimensions=(QualityDimension.ARGUMENTACION, QualityDimension.CONCLUSIONES),
        )

        dimension_scores = {
            QualityDimension.CLARIDAD: parsed_1.scores[QualityDimension.CLARIDAD],
            QualityDimension.COHERENCIA: parsed_1.scores[QualityDimension.COHERENCIA],
            QualityDimension.ARGUMENTACION: parsed_2.scores[QualityDimension.ARGUMENTACION],
            QualityDimension.CONCLUSIONES: parsed_2.scores[QualityDimension.CONCLUSIONES],
        }

        overall_score = sum(d.score for d in dimension_scores.values()) / len(dimension_scores)
        quality_level = get_quality_level_from_score(overall_score)

        return QualityResultDTO(
            overall_score=overall_score,
            quality_level=quality_level,
            dimension_scores={
                dimension.value: {"score": value.score, "feedback": value.feedback}
                for dimension, value in dimension_scores.items()
            },
        )

    def _ensure_call_produced_usable_content(
        self,
        parsed_response: _ParsedResponse,
        relevant_dimensions: tuple[QualityDimension, QualityDimension],
    ) -> None:
        if not any(dimension in parsed_response.matched_dimensions for dimension in relevant_dimensions):
            raise QualityAnalysisFailed()

    def _build_text_sample(self, document_content: DocumentContentDTO) -> str:
        parts = [document_content.title or ""]
        parts.extend(document_content.paragraphs[:3])
        middle_index = len(document_content.paragraphs) // 2
        parts.extend(document_content.paragraphs[middle_index : middle_index + 2])
        parts.extend(self._collect_conclusion_or_tail_paragraphs(document_content.paragraphs))

        text_sample = " ".join(parts)[:_TEXT_SAMPLE_CHARACTER_LIMIT]
        if len(text_sample.split()) < _MINIMUM_SAMPLE_WORD_COUNT:
            return " ".join(document_content.paragraphs)[:_TEXT_SAMPLE_CHARACTER_LIMIT]
        return text_sample

    def _collect_conclusion_or_tail_paragraphs(self, paragraphs: list[str]) -> list[str]:
        conclusion_paragraphs = []
        in_conclusion = False
        for paragraph in paragraphs:
            if re.search(r"conclusi", paragraph, re.IGNORECASE):
                in_conclusion = True
            if in_conclusion and not self._is_reference_like(paragraph):
                conclusion_paragraphs.append(paragraph)

        if conclusion_paragraphs:
            return conclusion_paragraphs[:3]

        non_reference_paragraphs = [p for p in paragraphs if not self._is_reference_like(p)]
        return non_reference_paragraphs[-2:]

    def _is_reference_like(self, paragraph: str) -> bool:
        return any(marker in paragraph[:80] for marker in _REFERENCE_LINE_MARKERS)

    def _build_prompt_one(self, text_sample: str) -> str:
        return f"""Eres un revisor editorial académico experto. Analiza este fragmento en DOS dimensiones.

TEXTO A ANALIZAR:
{text_sample}

INSTRUCCIONES:
1. Evalúa SOLO lo que está presente en el texto
2. Sé específico: menciona qué funciona bien y qué necesita mejorar
3. La ortografía y gramática ya fueron verificadas - enfócate en el CONTENIDO

FORMATO DE RESPUESTA (OBLIGATORIO):

**1. Claridad del argumento** [Puntuación: X/10]
[Analiza si el argumento central es claro. ¿El lector entiende fácilmente el mensaje principal?]

**2. Coherencia** [Puntuación: X/10]
[Analiza si las ideas se conectan lógicamente. ¿Hay transiciones claras entre secciones?]

CRITERIOS: 9-10 Excelente | 7-8 Bueno | 5-6 Aceptable | 3-4 Deficiente | 0-2 Inaceptable
"""

    def _build_prompt_two(self, text_sample: str) -> str:
        return f"""Eres un revisor editorial académico experto. Analiza este fragmento en DOS dimensiones.

TEXTO A ANALIZAR:
{text_sample}

INSTRUCCIONES:
1. Evalúa SOLO lo que está presente en el texto
2. Para Conclusiones: si no hay sección formal, infiere del contenido final del texto
3. La ortografía y gramática ya fueron verificadas - enfócate en el CONTENIDO

FORMATO DE RESPUESTA (OBLIGATORIO):

**1. Argumentación** [Puntuación: X/10]
[Si hay argumentos, evalúa su calidad. Si no los hay, indícalo claramente y asigna una puntuación baja.]

**2. Conclusiones** [Puntuación: X/10]
[OBLIGATORIO: Evalúa siempre. Si no hay sección formal, analiza el párrafo final del texto y asigna puntuación.]

CRITERIOS: 9-10 Excelente | 7-8 Bueno | 5-6 Aceptable | 3-4 Deficiente | 0-2 Inaceptable
"""

    def _parse_response(self, text: str) -> _ParsedResponse:
        scores = {
            dimension: _DimensionScore(_UNSCORED_DIMENSION_SCORE, _UNSCORED_DIMENSION_FEEDBACK)
            for dimension in QualityDimension
        }
        matched_dimensions: set[QualityDimension] = set()

        blocks = _DIMENSION_HEADER_PATTERN.split(text.strip())
        for block in blocks:
            if not block.strip():
                continue

            score = self._extract_score(block)
            feedback = self._extract_feedback(block)
            dimension = self._map_block_to_dimension(block)
            if dimension is None:
                continue

            scores[dimension] = _DimensionScore(score, feedback)
            matched_dimensions.add(dimension)

        return _ParsedResponse(scores=scores, matched_dimensions=frozenset(matched_dimensions))

    def _extract_score(self, block: str) -> float:
        match = _EXPLICIT_SCORE_PATTERN.search(block)
        if match is None:
            return self._infer_score_from_narrative(block)

        score_text = match.group(1) or match.group(2)
        try:
            return max(0.0, min(10.0, float(score_text)))
        except ValueError:
            return _UNSCORED_DIMENSION_SCORE

    def _infer_score_from_narrative(self, block: str) -> float:
        block_lower = block.lower()
        for keywords, score in _NARRATIVE_SCORE_KEYWORDS:
            if any(keyword in block_lower for keyword in keywords):
                return score
        return _UNSCORED_DIMENSION_SCORE

    def _extract_feedback(self, block: str) -> str:
        lines = block.strip().split("\n")
        feedback_lines = [line.strip() for line in lines[1:] if line.strip()]
        feedback = " ".join(feedback_lines)
        feedback = _RECOMMENDATION_TAIL_PATTERN.sub("", feedback).strip()
        feedback = " ".join(feedback.split())

        if len(feedback) < 10:
            return _UNSCORED_DIMENSION_FEEDBACK

        sentences = [s.strip() for s in feedback.split(".") if s.strip()]
        if len(sentences) > 3:
            return ". ".join(sentences[:3]) + "."
        return feedback

    def _map_block_to_dimension(self, block: str) -> QualityDimension | None:
        block_lower = block[:200].lower()
        if "argumentaci" in block_lower:
            return QualityDimension.ARGUMENTACION
        if "conclusi" in block_lower:
            return QualityDimension.CONCLUSIONES
        if "coherencia" in block_lower:
            return QualityDimension.COHERENCIA
        if "claridad" in block_lower or "argumento" in block_lower:
            return QualityDimension.CLARIDAD
        return None
```

### Notes on the port-call count requirement

The spec requires `generate()` to be called exactly twice per `analyze()` invocation. The
design above calls `self._llm_generator.generate(prompt_1)` then `generate(prompt_2)`
sequentially and unconditionally — satisfies the requirement directly, matches legacy's
sequential (not parallel/async) call order.

### `QualityDimension` enum and `dimension_scores` dict shape

`QualityResultDTO.dimension_scores` is `dict[str, dict[str, Any]]` (existing DTO, unchanged).
The domain service keys the final dict by `dimension.value` (the enum's string value), not by
the enum member itself, to match the DTO's `dict[str, ...]` contract and mirror legacy's
`dict[str, dict]` shape exactly (legacy used `"claridad"`, `"coherencia"`, etc. as literal
string keys).

```python
# src/domain/enums/quality_dimension.py
from enum import Enum


class QualityDimension(Enum):
    """The 4 semantic dimensions scored during quality analysis."""

    CLARIDAD = "claridad"
    COHERENCIA = "coherencia"
    ARGUMENTACION = "argumentacion"
    CONCLUSIONES = "conclusiones"
```

### Porting `get_quality_level_from_score`

Not yet in `src/domain/enums/quality_level.py` — confirmed by reading the file (only the enum
exists). Port the legacy function verbatim into the same module, as a plain function (no class
wrapper needed — it's a pure mapping function, consistent with "prefer classes" rule's
exception for non-domain-entity helpers... but to stay strictly within the skill's "prefer
classes" rule, it is added as a module-level function in `quality_level.py` since it is a
direct companion/factory function for the enum it lives next to, similar to how enums
sometimes expose classmethod-style lookups):

```python
# appended to src/domain/enums/quality_level.py
def get_quality_level_from_score(score: float) -> QualityLevel:
    """Map a numeric overall score to its corresponding QualityLevel."""
    if score >= 9.0:
        return QualityLevel.EXCELLENT
    if score >= 7.0:
        return QualityLevel.GOOD
    if score >= 5.0:
        return QualityLevel.ACCEPTABLE
    if score >= 3.0:
        return QualityLevel.NEEDS_IMPROVEMENT
    return QualityLevel.POOR
```

---

## `AnalyzeQualityUseCase` — Thin Pass-Through

```python
# src/application/analyze_quality_use_case.py
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.quality.quality_analyzer import QualityAnalyzer


class AnalyzeQualityUseCase:
    def __init__(self, quality_analyzer: QualityAnalyzer) -> None:
        self._quality_analyzer = quality_analyzer

    def execute(self, document_content: DocumentContentDTO, article_type) -> QualityResultDTO:
        return self._quality_analyzer.analyze(document_content, article_type)
```

`article_type` stays unused in the body — per proposal, tracked in the dead-code registry, not
cleaned up in this slice.

---

## `AnalyzeQualityUseCaseWiring` — First Wiring Assembling a Real Adapter

```python
# src/infrastructure/wirings/analyze_quality_use_case_wiring.py
from src.application.analyze_quality_use_case import AnalyzeQualityUseCase
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.domain.quality.quality_analyzer import QualityAnalyzer
from src.infrastructure.adapters.llm_generator.ollama_generator_adapter import (
    OllamaGeneratorAdapter,
)


class AnalyzeQualityUseCaseWiring:
    """Factory for building a ready-to-use AnalyzeQualityUseCase."""

    def create_use_case(self) -> AnalyzeQualityUseCase:
        return AnalyzeQualityUseCase(quality_analyzer=self._get_quality_analyzer())

    def _get_quality_analyzer(self) -> QualityAnalyzer:
        return QualityAnalyzer(llm_generator=self._get_llm_generator_port())

    def _get_llm_generator_port(self) -> LlmGeneratorPort:
        return OllamaGeneratorAdapter()
```

This follows the `create_use_case()` naming established by Slices 2-4's wirings
(`ValidateStructureWiring`, `ValidateApaWiring`, `MatchCitationsUseCaseWiring`), with
`_get_llm_generator_port` returning the `LlmGeneratorPort` type, not the concrete adapter —
the instance-based `_get_*` accessor pattern, now extended one level deeper to assemble a
concrete infrastructure adapter for the first time in the migration.

---

## Test Doubles

```python
# src/infrastructure/tests/test_doubles/analyze_quality_use_case_wiring_for_test.py
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.infrastructure.wirings.analyze_quality_use_case_wiring import (
    AnalyzeQualityUseCaseWiring,
)


class AnalyzeQualityUseCaseWiringForTest(AnalyzeQualityUseCaseWiring):
    def __init__(self, fake_llm_generator: LlmGeneratorPort) -> None:
        self._fake_llm_generator = fake_llm_generator

    def _get_llm_generator_port(self) -> LlmGeneratorPort:
        return self._fake_llm_generator
```

Domain tests for `QualityAnalyzer` use a minimal fake implementing `LlmGeneratorPort`
structurally (no inheritance required — `Protocol` satisfied by duck typing), returning
scripted responses per call to exercise: numbered/unnumbered headers, narrative-only scoring,
short-feedback fallback, sentence truncation, `argumentaci`-vs-`argumento` disambiguation,
partial-call success, and full-call failure raising `QualityAnalysisFailed`.

Adapter tests mock `ollama.generate` (patch at the `ollama` module level inside
`ollama_generator_adapter.py`) to verify: success path strips `response['response']`, and any
raised exception from `ollama.generate` becomes `LanguageModelUnavailable`.

---

## Risks / Open Items for Tasks Phase

- `get_quality_level_from_score` is a net-new function append to an existing file
  (`quality_level.py`) rather than a new file — slightly bends the "one function per file"
  rule's spirit, but it is a companion/factory function for the enum directly above it in the
  same file, consistent with how the file already groups the enum's domain concept. Tasks
  phase should add a corresponding test in `domain/tests/enums/test_quality_level.py` (new or
  extended) for the 5 threshold scenarios in the spec.
- `_ParsedResponse` and `_DimensionScore` are private dataclasses local to
  `quality_analyzer.py`, not promoted to `domain/dtos/`. They are implementation details that
  never cross the domain-service boundary (the public `analyze()` method only returns
  `QualityResultDTO`), so they correctly stay un-promoted per BaseDTO's purpose (transfer
  across layers).
