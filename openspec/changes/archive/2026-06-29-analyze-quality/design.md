# Design: analyze-quality (Slice 5)

> **Update note**: This revision supersedes the original PR-A design. PR #13 is open,
> not yet merged, with 273 passing tests against the original monolithic `QualityAnalyzer`
> (240 lines: sampling + prompt building + parsing + orchestration all in one class). The
> user's 6 follow-up decisions split that monolith into focused collaborators. ADR-1 through
> ADR-5 below (port/adapter shape, fallback constant, direct per-call assignment, full-call
> failure detection) are UNCHANGED from the original design and kept for reference. What
> changes is everything inside `quality_analyzer.py` and 2 brand-new files plus 2 prompt
> resource files — see "ADR-6 through ADR-9" below for the new decisions.

## Module Layout

First slice to introduce `src/domain/ports/`, `src/infrastructure/adapters/`, AND
`src/infrastructure/resources/` — three new top-level folders in the migration.

```
src/
├── domain/
│   ├── ports/
│   │   └── llm_generator_port.py                     # LlmGeneratorPort (Protocol) — unchanged
│   ├── enums/
│   │   ├── quality_dimension.py                       # member names renamed to English (CLARITY/COHERENCE/ARGUMENTATION/CONCLUSIONS) — .value strings unchanged (matched literally against LLM responses)
│   │   ├── quality_level.py                           # unchanged — get_quality_level_from_score() already present
│   │   └── reference_line_marker.py                   # NEW — HTTP/DOI/HTTPS/ISBN, replaces tuple constant
│   ├── dtos/
│   │   ├── quality_result_dto.py                      # unchanged — reused as-is
│   │   ├── dimension_score_dto.py                     # NEW — replaces private _DimensionScore
│   │   └── parsed_response_dto.py                     # NEW — replaces private _ParsedResponse
│   ├── quality/                                        # entity folder — now 3 files instead of 1
│   │   ├── quality_analyzer.py                        # REWRITTEN — thin orchestrator, ~70 lines
│   │   ├── quality_text_sampler.py                    # NEW — owns sampling heuristic
│   │   └── quality_response_parser.py                 # NEW — owns response-parsing logic
│   └── tests/
│       ├── dtos/
│       │   ├── test_dimension_score_dto.py            # NEW
│       │   └── test_parsed_response_dto.py             # NEW
│       ├── enums/
│       │   └── test_reference_line_marker.py          # NEW
│       └── quality/
│           ├── test_quality_analyzer.py               # rewritten — fake collaborators, not fake LLM + real parsing
│           ├── test_quality_text_sampler.py            # NEW — moves sampling scenarios out of analyzer tests
│           └── test_quality_response_parser.py         # NEW — moves parsing scenarios out of analyzer tests
├── application/
│   └── analyze_quality_use_case.py                    # unchanged — thin pass-through
└── infrastructure/
    ├── adapters/
    │   └── llm_generator/
    │       └── ollama_generator_adapter.py             # unchanged
    ├── resources/                                       # NEW — first resources folder in the migration
    │   └── prompts/
    │       └── quality/
    │           ├── clarity_coherence_prompt.txt      # NEW — Call 1 template, {text_sample} placeholder
    │           └── argumentation_conclusions_prompt.txt  # NEW — Call 2 template, {text_sample} placeholder
    ├── wirings/
    │   └── analyze_quality_use_case_wiring.py          # PR-B scope — will load 2 .txt files + 2 env vars
    └── tests/
        └── test_ollama_generator_adapter.py            # unchanged
```

`business_logic/quality_analyzer.py` stays untouched (coexistence). `requirements.txt` gets
`python-dotenv` added (PR-B). `.env.example` at repo root documents
`QUALITY_MIN_SAMPLE_WORD_COUNT=400` and `QUALITY_TEXT_SAMPLE_CHARACTER_LIMIT=8000` (PR-B).

---

## ADR-1 through ADR-5 (unchanged from original design)

Port-as-`Protocol`, adapter-wraps-`ollama.generate()`-directly, single fallback constant,
direct per-call assignment with no cross-call merge, and full-call parse-failure detection via
`matched_dimensions: frozenset[QualityDimension]` are all unchanged in substance. Only their
physical location moves: the `_ParsedResponse`/`_DimensionScore` private dataclasses described
in the original ADR-5 are now `ParsedResponseDTO`/`DimensionScoreDTO` — real `BaseDTO`
subclasses in `src/domain/dtos/`, because `QualityResponseParser.parse()` now returns
`ParsedResponseDTO` across a class boundary (parser → analyzer), which is exactly what
`BaseDTO` exists for — a private dataclass was acceptable on the original design only because
parsing and consuming happened inside the same class.

---

## ADR-6: Three-Way Split — Sampler / Parser / Orchestrator

**Decision**: Split the original 240-line `QualityAnalyzer` into 3 classes, each owning one
responsibility:

- `QualityTextSampler` — owns the strategic-excerpt heuristic (title + intro + middle +
  conclusion-or-tail, fallback to full text below a word threshold).
- `QualityResponseParser` — owns response parsing (header splitting, score extraction,
  narrative inference, feedback extraction/truncation, dimension mapping).
- `QualityAnalyzer` — owns orchestration only: sample once, render 2 prompts, call the port
  twice, parse twice, validate usability, assign dimensions directly, average, map to
  `QualityLevel`, return `QualityResultDTO`.

**Rationale**: the original design satisfied "one port call site" and "no Ollama leakage," but
left 3 unrelated responsibilities (sampling, parsing, orchestration) fused into one class —
each changes for a different reason (sampling heuristic tuning vs. LLM prompt-format changes
vs. orchestration flow), which is a Single Responsibility Principle violation the user
correctly flagged before merge. Splitting also makes each piece independently testable without
needing a fake `LlmGeneratorPort` to exercise sampling or parsing edge cases — the original
`test_quality_analyzer.py` was forced to drive sampling/parsing scenarios through a fake LLM
response just to reach the regex logic.

**Rejected alternative**: keep parsing/sampling as private methods but extract them to mixins
or module-level functions. Rejected — per the clean-architecture skill, "prefer classes (OOP)"
and "one class per file" for domain logic; mixins add inheritance coupling for no benefit here,
and module-level functions would require either passing constants as parameters on every call
or falling back to module constants (which blocks the constructor-parameter requirement below).

---

## ADR-7: Sampling Tunables Are Constructor Parameters, Not Module Constants

**Decision**: `QualityTextSampler.__init__(self, min_sample_word_count: int = 400,
text_sample_character_limit: int = 8000)`. No `os.getenv`, no `load_dotenv`, no environment
access anywhere in `src/domain/`.

**Rationale**: these 2 values are genuinely tunable per the user's decision — operators may
want a lower word-count threshold for short documents or a larger character budget for a
larger-context model. Reading env vars directly inside the domain class would violate the
import invariant (`src/domain/` must not depend on infrastructure/file/env concerns) and would
make the sampler untestable without monkeypatching `os.environ`. Constructor injection keeps
the domain class pure and lets PR-B's wiring resolve the env vars once, at startup, via
`python-dotenv` — consistent with how `AnalyzeQualityUseCaseWiring` already assembles
`OllamaGeneratorAdapter` in the original design.

**Rejected alternative**: keep them as module-level constants in `quality_text_sampler.py`
(original design's approach for `_MINIMUM_SAMPLE_WORD_COUNT` / `_TEXT_SAMPLE_CHARACTER_LIMIT`).
Rejected per the user's explicit decision — module constants can't be overridden without
editing source, whereas constructor parameters can be wired from configuration.

---

## ADR-8: `ReferenceLineMarker` Enum Replaces the Reference-Line Tuple

**Decision**: `_REFERENCE_LINE_MARKERS = ("http", "doi.org", "https", "ISBN")` becomes a real
`Enum` with members `HTTP`, `DOI`, `HTTPS`, `ISBN`. Membership check becomes
`any(marker.value in paragraph[:80] for marker in ReferenceLineMarker)`.

**Rationale**: this is a closed, named category of "what a reference line looks like" — exactly
what enums are for in this codebase (see `QualityDimension`, `QualityLevel`). A bare string
tuple loses the ability to refer to "the DOI marker" by name in tests or call sites, and risks
silent typos (`"htttp"`) that a tuple of strings can't catch at all whereas an enum member
reference at least fails loudly if misspelled as an attribute access.

**Rejected alternative**: leave it as a tuple but promote it to a shared module constant
importable from multiple files. Rejected — the user's decision specifically calls for an enum,
and a categorical "one of these 4 fixed string markers" concept is the textbook enum use case,
not a constant-sharing problem.

---

## ADR-9: Prompt Templates as External Files, Injected as Strings

**Decision**: the 2 prompt bodies move from Python f-string methods (`_build_prompt_one` /
`_build_prompt_two`) to plain `.txt` files under
`src/infrastructure/resources/prompts/quality/`, with a single `{text_sample}` placeholder
using Python's `.format()` syntax (not f-string, since the domain class only holds the loaded
string, not a live f-string context). `QualityAnalyzer` receives both template strings as
constructor parameters and renders each via one private `_render_prompt(template, text_sample)`
helper that calls `.format(text_sample=text_sample)`.

**Rationale**: the original `_build_prompt_one`/`_build_prompt_two` were ~20 lines each of
near-identical f-string boilerplate differing only in the embedded Spanish copy — a textbook
duplication smell, and one that couples prompt *wording* (a content/copy concern, likely to be
tuned by a non-engineer or A/B tested later) to the domain class's *source code* (requiring a
redeploy to change a single sentence). Externalizing to files: (a) collapses the duplicated
method into one generic renderer, (b) keeps `src/domain/` free of any file I/O — the domain
class never reads a file, it only receives the already-loaded string at construction time
(consistent with the constructor-injection pattern already used for `LlmGeneratorPort`,
`QualityTextSampler`, `QualityResponseParser`), and (c) lets future prompt-wording iteration
happen without touching Python code at all.

**Rejected alternative**: keep prompts as Python string constants (module-level instead of
method-level) inside `quality_analyzer.py`. Rejected — still mixes prompt copy with domain
source, and still requires a code change + redeploy for a wording tweak; only marginally better
than the original status quo. The user's decision explicitly calls for file-based templates
with wiring-time loading.

**Why `.format()` and not an f-string**: an f-string requires the variable to be in scope at
string-literal-definition time; the loaded template is a plain string read from disk, so
`.format(text_sample=...)` is the correct (and only) interpolation mechanism for a
runtime-loaded template string.

---

## `ReferenceLineMarker` Enum

```python
# src/domain/enums/reference_line_marker.py
from enum import Enum


class ReferenceLineMarker(Enum):
    """Substrings that mark a paragraph's opening characters as reference-like."""

    HTTP = "http"
    DOI = "doi.org"
    HTTPS = "https"
    ISBN = "ISBN"
```

---

## `DimensionScoreDTO` and `ParsedResponseDTO`

```python
# src/domain/dtos/dimension_score_dto.py
from dataclasses import dataclass

from src.domain.dtos.base_dto import BaseDTO


@dataclass(frozen=True)
class DimensionScoreDTO(BaseDTO):
    """A single dimension's parsed score and feedback text."""

    score: float
    feedback: str
```

```python
# src/domain/dtos/parsed_response_dto.py
from dataclasses import dataclass, field

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.dimension_score_dto import DimensionScoreDTO
from src.domain.enums.quality_dimension import QualityDimension


@dataclass(frozen=True)
class ParsedResponseDTO(BaseDTO):
    """The result of parsing one LLM response into per-dimension scores."""

    scores: dict[QualityDimension, DimensionScoreDTO] = field(default_factory=dict)
    matched_dimensions: frozenset[QualityDimension] = field(default_factory=frozenset)
```

Both follow the exact `BaseDTO` / `@dataclass(frozen=True)` pattern used by
`DocumentContentDTO` and `QualityResultDTO` — no `__str__` override needed (neither is rendered
directly to a user), `field(default_factory=...)` used for the mutable-typed defaults
(`dict`, `frozenset`) per dataclass rules, matching `DocumentContentDTO`'s use of
`field(default_factory=list)` for its own mutable-typed fields.

---

## `QualityTextSampler` — Full Design

```python
# src/domain/quality/quality_text_sampler.py
import re

from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.enums.reference_line_marker import ReferenceLineMarker

_CONCLUSION_HEADER_PATTERN = re.compile(r"conclusi", re.IGNORECASE)


class QualityTextSampler:
    """Builds a strategic text excerpt for LLM-based quality analysis."""

    def __init__(
        self,
        min_sample_word_count: int = 400,
        text_sample_character_limit: int = 8000,
        reference_line_prefix_length: int = 80,
        introduction_paragraph_count: int = 3,
        middle_paragraph_count: int = 2,
        conclusion_paragraph_limit: int = 3,
        fallback_tail_paragraph_count: int = 2,
    ) -> None:
        self._min_sample_word_count = min_sample_word_count
        self._text_sample_character_limit = text_sample_character_limit
        self._reference_line_prefix_length = reference_line_prefix_length
        self._introduction_paragraph_count = introduction_paragraph_count
        self._middle_paragraph_count = middle_paragraph_count
        self._conclusion_paragraph_limit = conclusion_paragraph_limit
        self._fallback_tail_paragraph_count = fallback_tail_paragraph_count

    def build_sample(self, document_content: DocumentContentDTO) -> str:
        """Return a strategic excerpt of the document, or its full text if too short."""
        parts = [document_content.title or ""]
        parts.extend(document_content.paragraphs[: self._introduction_paragraph_count])
        middle_index = len(document_content.paragraphs) // 2
        parts.extend(
            document_content.paragraphs[
                middle_index : middle_index + self._middle_paragraph_count
            ]
        )
        parts.extend(self._collect_conclusion_or_tail_paragraphs(document_content.paragraphs))

        text_sample = " ".join(parts)[: self._text_sample_character_limit]
        if len(text_sample.split()) < self._min_sample_word_count:
            return " ".join(document_content.paragraphs)[: self._text_sample_character_limit]
        return text_sample

    def _collect_conclusion_or_tail_paragraphs(self, paragraphs: list[str]) -> list[str]:
        conclusion_paragraphs = []
        in_conclusion = False
        for paragraph in paragraphs:
            if _CONCLUSION_HEADER_PATTERN.search(paragraph):
                in_conclusion = True
            if in_conclusion and not self._is_reference_like(paragraph):
                conclusion_paragraphs.append(paragraph)

        if conclusion_paragraphs:
            return conclusion_paragraphs[: self._conclusion_paragraph_limit]

        non_reference_paragraphs = [
            paragraph for paragraph in paragraphs if not self._is_reference_like(paragraph)
        ]
        return non_reference_paragraphs[-self._fallback_tail_paragraph_count :]

    def _is_reference_like(self, paragraph: str) -> bool:
        prefix = paragraph[: self._reference_line_prefix_length]
        return any(marker.value in prefix for marker in ReferenceLineMarker)
```

Logic is copied verbatim from legacy `_build_text_sample` /
`_collect_conclusion_or_tail_paragraphs` / `_is_reference_like` — the module constants
(`_MINIMUM_SAMPLE_WORD_COUNT`, `_TEXT_SAMPLE_CHARACTER_LIMIT`, `_REFERENCE_LINE_MARKERS`) and the
paragraph-slicing magic numbers (`3`, `2`, `3`, `2`) all become constructor parameters with
defaults identical to the legacy values — keeping every tunable value injectable (future `.env`
reads happen in infrastructure wiring, never in this domain class, which stays free of
`os`/`dotenv` imports) without changing behavior. `_CONCLUSION_HEADER_PATTERN` stays a module
constant — it is a Spanish-language domain heuristic (detects "conclusión"/"conclusiones"
headers), not a technical/tunable value.

---

## `QualityResponseParser` — Full Design

```python
# src/domain/quality/quality_response_parser.py
import re

from src.domain.dtos.dimension_score_dto import DimensionScoreDTO
from src.domain.dtos.parsed_response_dto import ParsedResponseDTO
from src.domain.enums.quality_dimension import QualityDimension

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
_DIMENSION_KEYWORDS: tuple[tuple[QualityDimension, tuple[str, ...]], ...] = (
    (QualityDimension.ARGUMENTATION, ("argumentaci",)),
    (QualityDimension.CONCLUSIONS, ("conclusi",)),
    (QualityDimension.COHERENCE, ("coherencia",)),
    (QualityDimension.CLARITY, ("claridad", "argumento")),
)


class QualityResponseParser:
    """Parses a single LLM response into per-dimension scores and feedback."""

    def __init__(
        self,
        unscored_dimension_score: float = 7.0,
        unscored_dimension_feedback: str = "No disponible",
    ) -> None:
        self._unscored_dimension_score = unscored_dimension_score
        self._unscored_dimension_feedback = unscored_dimension_feedback

    def parse(self, response_text: str) -> ParsedResponseDTO:
        """Return the parsed scores for every dimension found in the response text."""
        scores = {
            dimension: DimensionScoreDTO(
                self._unscored_dimension_score, self._unscored_dimension_feedback
            )
            for dimension in QualityDimension
        }
        matched_dimensions: set[QualityDimension] = set()

        blocks = _DIMENSION_HEADER_PATTERN.split(response_text.strip())
        for block in blocks:
            if not block.strip():
                continue

            dimension = self._map_block_to_dimension(block)
            if dimension is None:
                continue

            score = self._extract_score(block)
            feedback = self._extract_feedback(block)
            scores[dimension] = DimensionScoreDTO(score, feedback)
            matched_dimensions.add(dimension)

        return ParsedResponseDTO(scores=scores, matched_dimensions=frozenset(matched_dimensions))

    def _extract_score(self, block: str) -> float:
        match = _EXPLICIT_SCORE_PATTERN.search(block)
        if match is None:
            return self._infer_score_from_narrative(block)

        score_text = match.group(1) or match.group(2)
        try:
            return max(0.0, min(10.0, float(score_text)))
        except ValueError:
            return self._unscored_dimension_score

    def _infer_score_from_narrative(self, block: str) -> float:
        block_lower = block.lower()
        for keywords, score in _NARRATIVE_SCORE_KEYWORDS:
            if any(keyword in block_lower for keyword in keywords):
                return score
        return self._unscored_dimension_score

    def _extract_feedback(self, block: str) -> str:
        lines = block.strip().split("\n")
        feedback_lines = [line.strip() for line in lines[1:] if line.strip()]
        feedback = " ".join(feedback_lines)
        feedback = _RECOMMENDATION_TAIL_PATTERN.sub("", feedback).strip()
        feedback = " ".join(feedback.split())

        if len(feedback) < 10:
            return self._unscored_dimension_feedback

        sentences = [sentence.strip() for sentence in feedback.split(".") if sentence.strip()]
        if len(sentences) > 3:
            return ". ".join(sentences[:3]) + "."
        return feedback

    def _map_block_to_dimension(self, block: str) -> QualityDimension | None:
        block_lower = block[:200].lower()
        for dimension, keywords in _DIMENSION_KEYWORDS:
            if any(keyword in block_lower for keyword in keywords):
                return dimension
        return None
```

Logic copied verbatim from legacy `_parse_response` / `_extract_score` /
`_infer_score_from_narrative` / `_extract_feedback` / `_map_block_to_dimension`. The 3 regex
patterns and `_NARRATIVE_SCORE_KEYWORDS` constants relocate here as named module-level
constants — not enums, per the spec's explicit reasoning (a compiled regex isn't a categorical
value; a lone default isn't a category set), mirroring `_SECTION_ALIASES` in
`structure_validator.py`. `_UNSCORED_DIMENSION_SCORE`/`_UNSCORED_DIMENSION_FEEDBACK` become
constructor parameters with defaults identical to the legacy values (same rationale as
`QualityTextSampler`'s tunables: injectable without the domain reading `.env` directly).
`_map_block_to_dimension` uses the declarative `_DIMENSION_KEYWORDS` table instead of an
if/elif chain, for the same reason `_NARRATIVE_SCORE_KEYWORDS` is already declarative in this
file — same evaluation order, same substrings, no behavior change.

---

## `QualityAnalyzer` — Rewritten as Thin Orchestrator

```python
# src/domain/quality/quality_analyzer.py
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.parsed_response_dto import ParsedResponseDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.enums.quality_dimension import QualityDimension
from src.domain.enums.quality_level import get_quality_level_from_score
from src.domain.exceptions.quality_errors import QualityAnalysisFailed
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.domain.quality.quality_response_parser import QualityResponseParser
from src.domain.quality.quality_text_sampler import QualityTextSampler


class QualityAnalyzer:
    """Domain service that orchestrates LLM-backed quality scoring across 4 dimensions."""

    def __init__(
        self,
        llm_generator: LlmGeneratorPort,
        text_sampler: QualityTextSampler,
        response_parser: QualityResponseParser,
        clarity_coherence_prompt_template: str,
        argumentation_conclusions_prompt_template: str,
    ) -> None:
        self._llm_generator = llm_generator
        self._text_sampler = text_sampler
        self._response_parser = response_parser
        self._clarity_coherence_prompt_template = clarity_coherence_prompt_template
        self._argumentation_conclusions_prompt_template = (
            argumentation_conclusions_prompt_template
        )

    def analyze(self, document_content: DocumentContentDTO, article_type) -> QualityResultDTO:
        """Score document quality across Claridad, Coherencia, Argumentación and Conclusiones."""
        text_sample = self._text_sampler.build_sample(document_content)

        clarity_coherence_prompt = self._render_prompt(
            self._clarity_coherence_prompt_template, text_sample
        )
        argumentation_conclusions_prompt = self._render_prompt(
            self._argumentation_conclusions_prompt_template, text_sample
        )

        clarity_coherence_response = self._llm_generator.generate(clarity_coherence_prompt)
        argumentation_conclusions_response = self._llm_generator.generate(
            argumentation_conclusions_prompt
        )

        clarity_coherence_parsed = self._response_parser.parse(clarity_coherence_response)
        self._ensure_call_produced_usable_content(
            clarity_coherence_parsed,
            relevant_dimensions=(QualityDimension.CLARITY, QualityDimension.COHERENCE),
        )

        argumentation_conclusions_parsed = self._response_parser.parse(
            argumentation_conclusions_response
        )
        self._ensure_call_produced_usable_content(
            argumentation_conclusions_parsed,
            relevant_dimensions=(QualityDimension.ARGUMENTATION, QualityDimension.CONCLUSIONS),
        )

        dimension_scores = {
            QualityDimension.CLARITY: clarity_coherence_parsed.scores[QualityDimension.CLARITY],
            QualityDimension.COHERENCE: clarity_coherence_parsed.scores[
                QualityDimension.COHERENCE
            ],
            QualityDimension.ARGUMENTATION: argumentation_conclusions_parsed.scores[
                QualityDimension.ARGUMENTATION
            ],
            QualityDimension.CONCLUSIONS: argumentation_conclusions_parsed.scores[
                QualityDimension.CONCLUSIONS
            ],
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

    def _render_prompt(self, template: str, text_sample: str) -> str:
        return template.format(text_sample=text_sample)

    def _ensure_call_produced_usable_content(
        self,
        parsed_response: ParsedResponseDTO,
        relevant_dimensions: tuple[QualityDimension, QualityDimension],
    ) -> None:
        if not any(
            dimension in parsed_response.matched_dimensions for dimension in relevant_dimensions
        ):
            raise QualityAnalysisFailed()
```

~75 lines, down from 240. Zero `import re`, zero regex, zero file I/O, zero `os`/`dotenv` —
every line is either delegation to an injected collaborator or pure orchestration arithmetic
(`average`, `get_quality_level_from_score`, dict assembly). Satisfies the spec's "exactly one
class defined in this file" requirement — `DimensionScoreDTO`/`ParsedResponseDTO` no longer
live here.

---

## Prompt Template Files

```text
# src/infrastructure/resources/prompts/quality/clarity_coherence_prompt.txt
Eres un revisor editorial académico experto. Analiza este fragmento en DOS dimensiones.

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
```

```text
# src/infrastructure/resources/prompts/quality/argumentation_conclusions_prompt.txt
Eres un revisor editorial académico experto. Analiza este fragmento en DOS dimensiones.

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
```

Each file's body is verbatim from legacy's `_build_prompt_one`/`_build_prompt_two` f-strings,
with `{text_sample}` retained as a literal placeholder — identical syntax works for both
f-string (original) and `.format()` (new), so the rendered output is byte-for-byte unchanged.

---

## PR-B: `OllamaGeneratorAdapter` — Full Design

```python
# src/infrastructure/adapters/llm_generator/ollama_generator_adapter.py
import ollama

from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler
from src.domain.exceptions.language_model_errors import LanguageModelUnavailable
from src.domain.ports.llm_generator_port import LlmGeneratorPort


class OllamaGeneratorAdapter(LlmGeneratorPort):
    """Generates text via a local Ollama backend."""

    def __init__(self, model_name: str, base_url: str) -> None:
        self._model_name = model_name
        self._base_url = base_url

    @generic_error_handler
    def generate(self, prompt: str) -> str:
        """Return Ollama's generated text for the given prompt."""
        try:
            response = ollama.generate(model=self._model_name, prompt=prompt)
        except Exception as exc:
            raise LanguageModelUnavailable() from exc
        return response.get("response", "").strip()
```

**Decision**: constructor takes `model_name`/`base_url` as required parameters, no defaults —
the single source of truth for their fallback values is `AnalyzeQualityUseCaseWiring._get_llm_generator()`,
where `getenv("OLLAMA_MODEL_NAME", "llama3-gradient:8b-instruct-1048k-q4_K_M")` and the
equivalent for `base_url` live (matching legacy's values exactly). Duplicating the same default
literals in both the adapter's `__init__` and the wiring's `getenv()` call would mean two places
to keep in sync if the legacy default ever changes; the adapter stays a pure constructor-injected
class with no implicit behavior of its own. Calls the module-level
`ollama.generate(model=..., prompt=...)` function
directly — no `ollama.Client(host=...)` instantiation. The inner `try/except Exception` catches
any failure from the `ollama` library itself (connection error, timeout, backend error) and
re-raises as `LanguageModelUnavailable`; `@generic_error_handler` wraps the whole method so that
`LanguageModelUnavailable` (a `LanguageModelError` → `BaseSrcError` subclass, confirmed by
reading `src/domain/exceptions/language_model_errors.py` — NOT a `SrcBaseWarning`) passes through
unmodified per the decorator's `except BaseSrcError as exc: ... raise exc` branch (logging it
once via `was_error_logged`, then re-raising), while any exception the inner `try` didn't
anticipate (e.g. a bug in this adapter itself, not an Ollama failure) still gets logged and
wrapped into `SrcGenericError` by the decorator's outer `except Exception` branch — this is
exactly the layering `generic_error_handler` is built for, confirmed by reading its source:
`BaseSrcError` subtypes (including this one) re-raise untouched after logging, anything else
gets wrapped.

**Rejected alternative**: let the decorator alone translate the raw `ollama` exception into
`LanguageModelUnavailable`. Rejected — `generic_error_handler` has no per-adapter exception-
mapping mechanism; it only distinguishes "already a `BaseSrcError`" from "anything else," so the
adapter itself must perform the Ollama-specific translation before the decorator sees it.

**Rejected alternative**: keep the unused `self.client = ollama.Client(host=...)` field from
legacy. Rejected per the proposal's explicit "Out of Scope" — confirmed dead, never called.

`ollama_generator_adapter.py` is confirmed the only file in the slice importing `ollama`
(spec requirement, scenario "Adapter is the sole Ollama import site").

---

## PR-B: `AnalyzeQualityUseCase` — Full Design

```python
# src/application/analyze_quality_use_case.py
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.quality_result_dto import QualityResultDTO
from src.domain.quality.quality_analyzer import QualityAnalyzer


class AnalyzeQualityUseCase:
    def __init__(self, analyzer: QualityAnalyzer) -> None:
        self._analyzer = analyzer

    def execute(self, document_content: DocumentContentDTO, article_type) -> QualityResultDTO:
        return self._analyzer.analyze(document_content, article_type)
```

Identical shape to `MatchCitationsUseCase` (`src/application/match_citations_use_case.py`) —
constructor takes the single domain collaborator, `execute()` is one-line delegation, no
business logic, no type annotation on `article_type` (kept undocumented/unused per the
proposal's explicit dead-parameter decision — not cleaned up this slice).

---

## PR-B: `.env` Exposure Decision — Which Tunables Get a Variable

9 tunable constructor parameters exist across the two PR-A collaborators:

| Class | Parameter | Exposed via `.env`? | Reason |
|---|---|---|---|
| `QualityTextSampler` | `min_sample_word_count` | **Yes** — `QUALITY_MIN_SAMPLE_WORD_COUNT` | Spec's "Updated Dependencies for PR-B" names this exact var; legacy's only 2 originally-tunable values |
| `QualityTextSampler` | `text_sample_character_limit` | **Yes** — `QUALITY_TEXT_SAMPLE_CHARACTER_LIMIT` | Same as above |
| `QualityTextSampler` | `reference_line_prefix_length` | No — constructor default (`80`) | Structural heuristic (how many chars of a paragraph count as "the prefix" for reference detection), not an operational knob; no plausible deployment reason to change it without also reviewing the regex/marker logic itself |
| `QualityTextSampler` | `introduction_paragraph_count` | No — default (`3`) | Paragraph-slicing shape, tied to the prompt's expectations; changing it without re-tuning the prompt risks degrading LLM output, not a safe standalone tunable |
| `QualityTextSampler` | `middle_paragraph_count` | No — default (`2`) | Same reasoning as above |
| `QualityTextSampler` | `conclusion_paragraph_limit` | No — default (`3`) | Same reasoning as above |
| `QualityTextSampler` | `fallback_tail_paragraph_count` | No — default (`2`) | Same reasoning as above |
| `QualityResponseParser` | `unscored_dimension_score` | No — default (`7.0`) | Internal fallback value for unparseable LLM output; an operator changing this is tuning *failure-mode severity*, not a deployment concern — better reviewed/changed in code with tests, not via `.env` |
| `QualityResponseParser` | `unscored_dimension_feedback` | No — default (`"No disponible"`) | Same reasoning — a user-facing fallback string, changing it casually via `.env` risks an untested/untranslated value reaching end users |

**Rationale for the 2 exposed values**: both directly affect cost/quality tradeoffs an operator
plausibly tunes per deployment — a smaller `min_sample_word_count` lets short documents skip
full-text fallback sooner (faster, cheaper LLM calls), and `text_sample_character_limit` is a
direct token-budget control tied to the model's context window, which varies by which Ollama
model is configured. These are the same 2 values the *original* (pre-split) design already
flagged as the only operationally tunable parameters before the user's ADR-7 decision turned
all 7 sampler params into constructor parameters for testability — PR-B's `.env` exposure
narrows back down to those original 2, while the other 5 (plus the parser's 2) stay
constructor-injectable-but-not-`.env`-sourced: still overridable in code/tests, just not exposed
as a deployment knob.

**Rejected alternative**: expose all 9 via `.env` since they're all already constructor
parameters. Rejected — `.env` exposure is a UX/operational surface, not a structural
constraint; exposing heuristic/failure-mode values as environment variables invites
misconfiguration with no corresponding test coverage at deploy time, for parameters whose
"correct" values are coupled to the prompt wording and parsing regexes, not to deployment
environment.

---

## PR-B: `AnalyzeQualityUseCaseWiring` — Full Design

```python
# src/infrastructure/resources/prompts/quality/__init__.py
from os import path

PROMPTS_DIR = path.dirname(__file__)
```

```python
# src/infrastructure/wirings/analyze_quality_use_case_wiring.py
from os.path import join
from pathlib import Path

from dotenv import load_dotenv
from os import getenv

from src.application.analyze_quality_use_case import AnalyzeQualityUseCase
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.domain.quality.quality_analyzer import QualityAnalyzer
from src.domain.quality.quality_response_parser import QualityResponseParser
from src.domain.quality.quality_text_sampler import QualityTextSampler
from src.infrastructure.adapters.llm_generator.ollama_generator_adapter import (
    OllamaGeneratorAdapter,
)
from src.infrastructure.resources.prompts.quality import PROMPTS_DIR

load_dotenv()


class AnalyzeQualityUseCaseWiring:
    """Factory for building a ready-to-use AnalyzeQualityUseCase."""

    def create_use_case(self) -> AnalyzeQualityUseCase:
        return AnalyzeQualityUseCase(analyzer=self._get_quality_analyzer())

    def _get_quality_analyzer(self) -> QualityAnalyzer:
        return QualityAnalyzer(
            llm_generator=self._get_llm_generator(),
            text_sampler=self._get_text_sampler(),
            response_parser=QualityResponseParser(),
            clarity_coherence_prompt_template=self._read_prompt_template(
                "clarity_coherence_prompt.txt"
            ),
            argumentation_conclusions_prompt_template=self._read_prompt_template(
                "argumentation_conclusions_prompt.txt"
            ),
        )

    def _get_llm_generator(self) -> LlmGeneratorPort:
        model_name = getenv("OLLAMA_MODEL_NAME", "llama3-gradient:8b-instruct-1048k-q4_K_M")
        base_url = getenv("OLLAMA_BASE_URL", "http://localhost:11434")
        return OllamaGeneratorAdapter(model_name=model_name, base_url=base_url)

    def _get_text_sampler(self) -> QualityTextSampler:
        return QualityTextSampler(
            min_sample_word_count=int(getenv("QUALITY_MIN_SAMPLE_WORD_COUNT", "400")),
            text_sample_character_limit=int(
                getenv("QUALITY_TEXT_SAMPLE_CHARACTER_LIMIT", "8000")
            ),
        )

    def _read_prompt_template(self, filename: str) -> str:
        file_path = Path(join(PROMPTS_DIR, filename))
        return file_path.read_text(encoding="utf-8")
```

**Decision**: `PROMPTS_DIR` is exported from the resources package's own `__init__.py`
(`src/infrastructure/resources/prompts/quality/__init__.py`) as `path.dirname(__file__)`,
not computed inline in the wiring via `Path(__file__).resolve().parents[N]`. This is the
**reusable convention going forward for any future slice that needs to read its own resource
files**: a resources package exports its own directory constant, so wiring/adapter code that
consumes those resources never has to know or maintain a relative-path depth (`parents[1]`,
`parents[2]`, ...) from its own file location to the resources folder — that coupling lives
once, inside the resources package itself, and survives the wiring file moving or being
restructured. `_read_prompt_template` joins `PROMPTS_DIR` and the filename via `os.path.join`
(consistent string-based path building, no `pathlib`-operator (`/`) mixing), then wraps the
joined string in `Path(...)` before calling `.read_text(encoding="utf-8")` — `os.path.join`
itself returns a plain `str`, which has no `.read_text()` method, so the `Path(...)` wrap is
required for the call to resolve.

`load_dotenv()` is called once at module load time (matching the common `python-dotenv` usage
pattern: load the `.env` file into `os.environ` before any `getenv()` call reads from it), not
inside a method — wiring classes in this codebase are already instantiated once per process
(see `ValidateApaWiring`, `MatchCitationsUseCaseWiring`), so module-level `load_dotenv()` runs
exactly once per process too, with no risk of being called multiple times per use-case
construction.

**Rejected alternative**: call `load_dotenv()` inside `_get_text_sampler()` or
`_get_llm_generator()`. Rejected — calling it per-accessor is redundant (the `.env` file doesn't
change between accessor calls within one process) and `python-dotenv`'s own convention is to
call `load_dotenv()` once near process start.

**Rejected alternative**: read `getenv()` results directly without `int()` conversion, relying
on `QualityTextSampler`'s constructor type hints. Rejected — `os.getenv()` always returns `str |
None`; without an explicit `int()` cast the wiring would pass a string into a parameter typed
`int`, silently breaking arithmetic comparisons (`len(text_sample.split()) <
self._min_sample_word_count`) the first time a value is read from `.env` instead of the Python
default.

---

## PR-B: `.env.example` Content

```dotenv
# .env.example
# Ollama backend connection (OllamaGeneratorAdapter)
OLLAMA_MODEL_NAME=llama3-gradient:8b-instruct-1048k-q4_K_M
OLLAMA_BASE_URL=http://localhost:11434

# Quality analysis text sampling (QualityTextSampler)
QUALITY_MIN_SAMPLE_WORD_COUNT=400
QUALITY_TEXT_SAMPLE_CHARACTER_LIMIT=8000
```

First `.env.example` in the repository. Values shown are the same defaults already baked into
the constructors, so an operator copying this file unmodified to `.env` reproduces current
behavior exactly — `.env.example` documents the override surface, it does not change any
default.

---

## PR-B: `requirements.txt` Addition

```diff
 python-docx
 gradio
 ollama
 language-tool-python
 Pillow
 pywin32; sys_platform == 'win32'
 pytest
+python-dotenv
```

Confirmed absent from `requirements.txt` (read before this design was written) and not already
a transitive dependency of any listed package — first new third-party dependency added since
the migration began, per the spec's "Risks / Open Items" note.

---

## PR-B: Adapter Test Strategy

```python
# src/infrastructure/tests/test_ollama_generator_adapter.py
@patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama")
def test_generate_returns_stripped_response_text(self, mock_ollama):
    mock_ollama.generate.return_value = {"response": "  some text  "}
    adapter = OllamaGeneratorAdapter()
    assert adapter.generate("prompt") == "some text"

@patch("src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama")
def test_generate_raises_language_model_unavailable_on_backend_failure(self, mock_ollama):
    mock_ollama.generate.side_effect = ConnectionError("backend unreachable")
    adapter = OllamaGeneratorAdapter()
    with self.assertRaises(LanguageModelUnavailable):
        adapter.generate("prompt")
```

Mocks the `ollama` module imported inside the adapter file (not the third-party package
globally) — same `@patch("...module_under_test.ollama")` pattern any adapter test in this
codebase would use to isolate the real network call. `ConnectionError` stands in for "any
backend exception" per the spec scenario; the test only needs to prove the catch-and-reraise
path, not enumerate every possible `ollama` failure mode.

---

## Test Doubles (updated)

`QualityAnalyzer`'s test double construction changes shape — tests now inject real
`QualityTextSampler()` / `QualityResponseParser()` instances (no fakes needed, they're pure and
fast) alongside a fake `LlmGeneratorPort`, plus literal prompt template strings containing
`{text_sample}`. This isolates `test_quality_analyzer.py` to orchestration concerns only
(port called twice, dimensions assigned from the correct call, failure raised on full-call
parse failure) — sampling and parsing edge cases move to their own dedicated test files
(`test_quality_text_sampler.py`, `test_quality_response_parser.py`), each driving the
respective class directly with no LLM fake required.

```python
fake_llm_generator = FakeLlmGeneratorPort(responses=[response_1_text, response_2_text])
analyzer = QualityAnalyzer(
    llm_generator=fake_llm_generator,
    text_sampler=QualityTextSampler(),
    response_parser=QualityResponseParser(),
    clarity_coherence_prompt_template="...{text_sample}...",
    argumentation_conclusions_prompt_template="...{text_sample}...",
)
```

Adapter tests (`OllamaGeneratorAdapter`) are unchanged from the original design.

---

## Risks / Open Items for Tasks Phase

- All 273 existing PR-A tests for the old monolithic `quality_analyzer.py` must be redistributed:
  sampling-heuristic tests move to `test_quality_text_sampler.py`, parsing tests move to
  `test_quality_response_parser.py`, and only orchestration tests remain in
  `test_quality_analyzer.py`. Tasks phase must enumerate this redistribution explicitly so no
  scenario is silently dropped during the restructuring.
- `python-dotenv` is a new third-party dependency — first one added since the migration began.
  Tasks phase should confirm it's not already transitively available before adding to
  `requirements.txt`.
- `src/infrastructure/resources/` is a new top-level folder not previously listed in the
  clean-architecture skill's standard skeleton — acceptable as a deliberate, documented
  extension (skill doc explicitly allows non-skeleton folders for adapter-specific needs);
  worth a one-line addition to the skill file in a future housekeeping pass, not blocking this
  slice.
- `QualityResponseParser` and `QualityTextSampler` are both stateless after construction
  (`QualityResponseParser` takes no constructor args at all) — confirms they are domain
  *services* per the clean-architecture naming convention (no `Service` suffix, descriptive
  name), not entities.
