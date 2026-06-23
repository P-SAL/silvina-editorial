# Design: classify-article (Slice 6)

## Module Layout

Second slice to touch `src/domain/ports/`, `src/infrastructure/adapters/llm_generator/`, AND
`src/infrastructure/resources/prompts/` — all three already exist from Slice 5
(`analyze-quality`); this slice extends/reuses them rather than introducing new top-level
folders.

```
src/
├── domain/
│   ├── ports/
│   │   └── llm_generator_port.py                       # MODIFIED — generate() gains options param
│   ├── enums/
│   │   ├── article_size.py                              # MODIFIED — adds classify_article_size()
│   │   └── classification_confidence.py                 # NEW — ClassificationConfidence(float, Enum), 5 members
│   ├── classification/                                   # NEW entity folder — 4 files
│   │   ├── imryd_signal_detector.py                      # NEW — renamed StructureAnalyzer, verbatim port
│   │   ├── article_classification_text_sampler.py        # NEW — owns _build_text_sample() heuristic
│   │   ├── article_classification_response_parser.py     # NEW — owns S4/S5/S6 response parsing
│   │   └── article_classifier.py                         # NEW — thin orchestrator, 19-case rule table
│   └── tests/
│       ├── enums/
│       │   ├── test_classification_confidence.py         # NEW
│       │   └── test_classify_article_size.py             # NEW
│       └── classification/
│           ├── __init__.py                               # NEW
│           ├── fake_llm_generator_port.py                 # NEW — test double, mirrors quality/fake_llm_generator_port.py
│           ├── test_imryd_signal_detector.py               # NEW
│           ├── test_article_classification_text_sampler.py # NEW
│           ├── test_article_classification_response_parser.py # NEW
│           ├── test_article_classifier_imryd_override.py   # NEW — case 1
│           ├── test_article_classifier_cientifico.py        # NEW — cases 2-5
│           ├── test_article_classifier_divulgacion_near_miss.py # NEW — cases 6-9
│           ├── test_article_classifier_divulgacion_standard.py  # NEW — cases 10-18
│           └── test_article_classifier_opinion.py           # NEW — case 19
├── application/
│   ├── classify_article_use_case.py                       # NEW — thin pass-through
│   └── tests/
│       └── test_classify_article_use_case.py               # NEW
└── infrastructure/
    ├── adapters/
    │   └── llm_generator/
    │       └── ollama_generator_adapter.py                 # MODIFIED — generate() forwards options
    ├── resources/
    │   ├── text_resource_loader.py                          # NEW — shared read_text_resource() helper (ADR-8)
    │   └── prompts/
    │       └── classification/                             # NEW
    │           ├── __init__.py                              # NEW — PROMPTS_DIR, same pattern as quality/__init__.py
    │           └── s4_s5_s6_signal_prompt.txt                # NEW — externalized Spanish prompt
    ├── wirings/
    │   ├── classify_article_use_case_wiring.py               # NEW — uses read_text_resource()
    │   └── analyze_quality_use_case_wiring.py                # MODIFIED — retrofitted to use read_text_resource() (ADR-8)
    └── tests/
        ├── test_classify_article_use_case_wiring.py          # NEW
        ├── test_text_resource_loader.py                      # NEW — shared helper unit test
        └── test_ollama_generator_adapter.py                  # MODIFIED — adds options-forwarding cases

tests/
└── smoke/
    └── test_classify_article_parity.py                       # NEW — legacy/new parity, both LLM calls mocked (ADR-9)
```

`business_logic/structure_analyzer.py` and `business_logic/article_classifier.py` stay
untouched (coexistence). `.env.example` gains `ARTICLE_CLASSIFIER_TEMPERATURE=0.1` and
`ARTICLE_CLASSIFIER_NUM_PREDICT=300`. No new third-party dependency — `ollama`,
`python-dotenv` already present from Slice 5.

---

## ADR-1: Renamed Signal Detector — `ImrydSignalDetector`

**Decision**: `business_logic/structure_analyzer.py`'s `StructureAnalyzer` is renamed to
`ImrydSignalDetector`, living at `src/domain/classification/imryd_signal_detector.py`. Single
public method `detect(document_content: DocumentContentDTO) -> dict[str, bool]`, ported
verbatim — same `IMRYD_KEYWORDS` table (bilingual), same ≤5-word header-candidate filter, same
6-key signal dict (`has_introduction`/`has_methods`/`has_results`/`has_discussion`/
`has_conclusion`/`imryd_complete`), same `imryd_complete` semantics (requires intro + methods +
results + discussion; conclusion is detected but NOT required for completeness).

**Rationale**: the proposal fixed the constraint that the migrated class must NOT be named
`StructureAnalyzer` (collision with `src/domain/structure/structure_validator.py`'s
`StructureValidator` — different problem, same misleading "Structure"-prefix). `ImrydSignalDetector`
names the class by what it actually detects (the IMRyD acronym — Introduction, Methods, Results,
and Discussion section presence) rather than by a generic "structure" label, which simultaneously
resolves the collision and is more precise than the legacy name ever was. The method name
`detect()` (not `analyze()`) follows the same precision logic — this class detects boolean
signals, it does not perform the broader "analysis" `StructureAnalyzer`'s name implied but its
actual scope (5 keyword lookups + 1 derived flag) never delivered.

**Rejected alternative**: `ImrydDetector` (proposal's example name). Rejected — "Imryd" alone
reads as a noun needing a qualifier; `ImrydSignalDetector` makes explicit that this class's
output (`dict[str, bool]`) is consumed as *signals* feeding `ArticleClassifier`'s rule table,
consistent with how the legacy classifier's own comments label `S2a`/`S2b`/`S3`/etc. as
"signals" — the new name reuses that vocabulary instead of introducing a parallel one.

**Rejected alternative**: `SectionPresenceAnalyzer`. Rejected — "Analyzer" suffix is exactly the
generic-sounding pattern this rename is trying to move away from (and is also the suffix
`QualityAnalyzer` already uses for a meaningfully heavier orchestration role — reusing it here
for a 6-keyword-lookup class would flatten an important scope difference).

---

## ADR-2: Four-Way Split Mirroring `QualityAnalyzer`'s ADR-6

**Decision**: split the legacy 280-line `ArticleClassifier` into 4 collaborating classes, same
shape as Slice 5's sampler/parser/orchestrator split, with one extra because classification has
an independent deterministic-signal detector that quality analysis does not:

- `ImrydSignalDetector` — pure, deterministic, zero LLM dependency (ADR-1 above).
- `ArticleClassificationTextSampler` — owns `_build_text_sample()`'s heuristic (first 3500 +
  last 2500 chars, bibliography-skip via short-paragraph marker detection).
- `ArticleClassificationResponseParser` — owns parsing the combined S4/S5/S6 response (3
  regex-extracted yes/no answers from one free-text block).
- `ArticleClassifier` — orchestration only: compute `article_size`, run the IMRyD override
  check, compute signals S2a/S2b/S3 deterministically, sample text, call the LLM port once for
  S4/S5/S6, parse, then apply the verbatim 19-case rule table.

**Rationale**: identical SRP reasoning as `analyze-quality`'s ADR-6 — sampling heuristic tuning,
LLM response parsing, and rule-table orchestration each change for unrelated reasons and need
independent test coverage without forcing every scenario through a fake LLM response. The
4th component (`ImrydSignalDetector`) is not a refactor split of legacy `ArticleClassifier` —
it was already a separate legacy class (`StructureAnalyzer`); this slice only renames and
relocates it, consistent with ADR-1.

**Rejected alternative**: fold `ImrydSignalDetector`'s logic directly into `ArticleClassifier`
as a private method (un-doing the legacy author's own separation). Rejected — the legacy
codebase already chose to keep this as an independently testable, dependency-free class; folding
it in would be a step backward in testability with no compensating benefit, and would contradict
the proposal's explicit scope item ("ported as its own injectable unit").

**Naming note — why `ArticleClassificationTextSampler`/`ArticleClassificationResponseParser`,
not `ClassificationTextSampler`/`ClassificationResponseParser`**: the longer names disambiguate
against a hypothetical future `StructureClassificationTextSampler` or similar if classification
ever grows additional LLM-backed sub-signals; `Quality*` and `ArticleClassification*` as prefixes
both name "what domain capability this sampler/parser serves," consistent with each other even
though their lengths differ slightly. Kept short would have been `ClassificationTextSampler` —
rejected only because `Classification` alone is ambiguous (classification of what — articles,
citations, sections?) in a codebase that already has `ClassificationCategory`,
`ClassificationResultDTO`, and `ClassificationConfidence` all under the generic "Classification"
umbrella; prefixing with `Article` ties the sampler/parser unambiguously to `ArticleClassifier`.

---

## `ImrydSignalDetector` — Full Design

```python
# src/domain/classification/imryd_signal_detector.py
from src.domain.dtos.document_content_dto import DocumentContentDTO

_IMRYD_KEYWORDS: dict[str, tuple[str, ...]] = {
    "introduction": (
        "introduction", "background", "context",
        "introducción", "introduccion", "intro",
    ),
    "methods": (
        "method", "methodology", "materials", "procedures",
        "método", "metodo", "metodología", "metodologia",
        "métodos", "metodos", "materiales",
    ),
    "results": ("results", "findings", "resultados", "hallazgos"),
    "discussion": ("discussion", "discusión", "discusion"),
    "conclusion": (
        "conclusion", "conclusions", "concluding",
        "conclusión", "conclusiones",
    ),
}
_HEADER_CANDIDATE_MAX_WORD_COUNT = 5


class ImrydSignalDetector:
    """Domain service that detects IMRyD section-presence signals in a document.

    Scans only short paragraphs (<= 5 words) as section-header candidates, to avoid
    false positives from body prose containing section-name words (e.g. "resultados"
    appearing mid-sentence rather than as a heading).
    """

    def detect(self, document_content: DocumentContentDTO) -> dict[str, bool]:
        """Return IMRyD section-presence signals for the given document."""
        header_candidates = [
            paragraph.strip().lower()
            for paragraph in document_content.paragraphs
            if 1 <= len(paragraph.strip().split()) <= _HEADER_CANDIDATE_MAX_WORD_COUNT
        ]

        signals = {
            "has_introduction": False,
            "has_methods": False,
            "has_results": False,
            "has_discussion": False,
            "has_conclusion": False,
            "imryd_complete": False,
        }

        for section, keywords in _IMRYD_KEYWORDS.items():
            if any(
                keyword in header for keyword in keywords for header in header_candidates
            ):
                signals[f"has_{section}"] = True

        signals["imryd_complete"] = (
            signals["has_introduction"]
            and signals["has_methods"]
            and signals["has_results"]
            and signals["has_discussion"]
        )

        return signals
```

Ported verbatim from `StructureAnalyzer.analyze()` — same keyword lists (now an immutable
module-level `tuple`-valued dict instead of a class attribute holding `list`s, since this
service is stateless and has no constructor parameters worth injecting; nothing here is a
genuinely tunable value the way `QualityTextSampler`'s thresholds are — the keyword lists are
the algorithm itself, not an operational knob). `_HEADER_CANDIDATE_MAX_WORD_COUNT` stays a module
constant for the same reason `_CONCLUSION_HEADER_PATTERN` stays a module constant in
`quality_text_sampler.py`: it is part of the detection heuristic's definition, not a deployment
tunable. No constructor at all (matching `StructureValidator`'s own `__init__(self) -> None:
pass` precedent for a zero-state domain service) — omitted entirely here since it would be
empty boilerplate; instantiation is simply `ImrydSignalDetector()`.

**Rejected alternative**: keep `IMRYD_KEYWORDS` as a public class attribute (legacy shape).
Rejected — nothing outside this class reads the keyword table directly (confirmed: legacy
`ArticleClassifier` only ever calls `.analyze()`, never touches `.IMRYD_KEYWORDS`), so making it
a private module constant is a strict improvement in encapsulation with zero behavior change.

---

## ADR-3: `ClassificationConfidence(float, Enum)` — Member Names Tied to Trigger Context

**Decision**: 5 members, each named after the specific case in the legacy `_apply_rule` table
that produces that literal value — not generic tiers like `HIGH`/`MEDIUM`/`LOW`, because the
legacy table's 4 non-override confidence values (0.90/0.86/0.85/0.83) are each triggered by a
*different missing signal*, not by a generic "how confident" scale:

```python
# src/domain/enums/classification_confidence.py
from enum import Enum


class ClassificationConfidence(float, Enum):
    """Confidence levels assigned by ArticleClassifier's CIENTIFICO rule-table branches."""

    IMRYD_OVERRIDE = 0.95
    FULL_SIGNAL_MATCH = 0.90
    RECENT_BIBLIOGRAPHY_SUPPORT = 0.86
    COMPLETE_BIBLIOGRAPHY_SUPPORT = 0.85
    SUFFICIENT_REFERENCE_COUNT = 0.83
```

| Member | Value | Legacy case | Triggering condition |
|---|---|---|---|
| `IMRYD_OVERRIDE` | 0.95 | case 1 (deterministic override) | `imryd_complete` AND size != `FUERA_RANGO` — bypasses the LLM signals entirely |
| `FULL_SIGNAL_MATCH` | 0.90 | case 2 | S3+S4+S5+**S2a+S2b+S6** all present — every signal fires |
| `RECENT_BIBLIOGRAPHY_SUPPORT` | 0.86 | case 3 | S3+S4+S5+**S2b+S6**, missing S2a (reference *count* threshold) — bibliography is recent but not large enough |
| `COMPLETE_BIBLIOGRAPHY_SUPPORT` | 0.85 | case 4 | S3+S4+S5+**S2a+S2b**, missing S6 — bibliography is both large enough and recent, but no theoretical-framework justification detected |
| `SUFFICIENT_REFERENCE_COUNT` | 0.83 | case 5 | S3+S4+S5+**S2a+S6**, missing S2b — enough references, with theoretical justification, but bibliography isn't recent |

**Rationale for the naming approach**: each name describes *what bibliographic/structural
condition is present or absent* at that confidence level, which is exactly the information a
future reader needs when seeing `ClassificationConfidence.RECENT_BIBLIOGRAPHY_SUPPORT` in a test
assertion or log line — it tells them which signal combination produced this exact number
without needing to cross-reference the rule table. A generic tier name
(`VERY_HIGH`/`HIGH`/`MEDIUM_HIGH`/`MEDIUM`) would require that same cross-reference anyway, so it
buys nothing over the current bare float literals it's replacing. `IMRYD_OVERRIDE` is named for
the override *mechanism* (not a signal-absence description) because case 1 is structurally
different from the other 4 — it is the only confidence value produced by a deterministic
short-circuit that skips signal computation altogether, not a graded outcome of the same 6-signal
evaluation the other 4 share.

**Verified against `_apply_rule` literally**: re-read `business_logic/article_classifier.py`
lines 320-370 to confirm exactly one missing signal differs between cases 2/3/4/5 relative to the
full S2a+S2b+S6 set, and that the docstring's own case labels ("case 2", "case 3", etc.) match
1:1 with the `if`/`elif` branches in source — no discrepancy found.

**Rejected alternative**: `EXACT_MATCH = 0.95` (proposal's illustrative example name). Rejected
on inspection — "exact match" does not describe what triggers 0.95 (the IMRyD structural
override), and would be actively misleading since cases 2-5 are themselves exact matches against
their respective signal subsets; the proposal's example was acknowledged there as illustrative
only ("verify each value's actual usage context… to name them meaningfully" was the explicit
design-phase instruction), not a fixed name.

**Rejected alternative**: name members by signal-bitmask shorthand, e.g. `S2A_S2B_S6 = 0.90`.
Rejected — couples the enum's public names to the legacy code's internal `S2a`/`S2b`/`S6`
shorthand, which is meaningless without reading the rule table's comments; the chosen names are
self-documenting without that cross-reference.

---

## ADR-4: `LlmGeneratorPort` / `OllamaGeneratorAdapter` — Additive `options` Parameter

**Exact diff against current code**:

```diff
 # src/domain/ports/llm_generator_port.py
 from typing import Protocol


 class LlmGeneratorPort(Protocol):
     """Capability to generate text from a prompt via a language model backend."""

-    def generate(self, prompt: str) -> str:
+    def generate(self, prompt: str, options: dict | None = None) -> str:
         """Return the generated text for the given prompt."""
         ...
```

```diff
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
-    def generate(self, prompt: str) -> str:
+    def generate(self, prompt: str, options: dict | None = None) -> str:
         """Return Ollama's generated text for the given prompt."""
         try:
-            response = ollama.generate(model=self._model_name, prompt=prompt)
+            response = ollama.generate(
+                model=self._model_name, prompt=prompt, options=options
+            )
         except (ollama.RequestError, ollama.ResponseError, ConnectionError) as exc:
             raise LanguageModelUnavailable() from exc
         return response.get("response", "").strip()
```

**Confirmed against actual current source** (not the `analyze-quality` design doc's snapshot,
which shows a broader `except Exception` — the merged code narrows this to
`(ollama.RequestError, ollama.ResponseError, ConnectionError)`): the `except` clause is
unchanged by this diff: `options=None` forwarded to `ollama.generate(options=None)` is
equivalent to omitting the kwarg entirely per the `ollama` library's own handling of `None`
options, so `analyze-quality`'s existing call site `self._llm_generator.generate(prompt)` (no
`options` argument) continues to bind `options=None` by default and produces byte-identical
behavior to pre-change code — confirmed by reading `quality_analyzer.py`'s two
`self._llm_generator.generate(...)` call sites, neither of which passes `options`.

**Rationale**: matches the proposal's fixed constraint exactly. No adapter-side translation
logic needed beyond the parameter forward — `ollama.generate()` already accepts `options` as a
native kwarg.

---

## `ArticleClassificationTextSampler` — Full Design

```python
# src/domain/classification/article_classification_text_sampler.py
from src.domain.dtos.document_content_dto import DocumentContentDTO

_INTRODUCTION_CHARACTER_LIMIT = 3500
_CONCLUSION_CHARACTER_LIMIT = 2500
_FALLBACK_CHARACTER_LIMIT = 6000
_BIBLIOGRAPHY_HEADER_MAX_LENGTH = 30
_BIBLIOGRAPHY_MARKERS = ("referencias", "bibliografía", "bibliography", "fuentes bibliográficas")


class ArticleClassificationTextSampler:
    """Builds a strategic text excerpt for LLM-based article-classification signals."""

    def build_sample(self, document_content: DocumentContentDTO) -> str:
        """Return the first 3500 + last 2500 chars of the document, skipping the bibliography."""
        full_text = " ".join(document_content.paragraphs)
        clean_text = self._strip_bibliography(document_content.paragraphs, full_text)

        introduction = clean_text[:_INTRODUCTION_CHARACTER_LIMIT]
        ending = (
            clean_text[-_CONCLUSION_CHARACTER_LIMIT:]
            if len(clean_text) > _INTRODUCTION_CHARACTER_LIMIT
            else ""
        )
        sample = (introduction + " " + ending).strip()
        return sample or full_text[:_FALLBACK_CHARACTER_LIMIT]

    def _strip_bibliography(self, paragraphs: list[str], full_text: str) -> str:
        bibliography_position = len(full_text)
        character_position = 0
        for paragraph in paragraphs:
            paragraph_lower = paragraph.strip().lower()
            is_bibliography_header = (
                len(paragraph_lower) <= _BIBLIOGRAPHY_HEADER_MAX_LENGTH
                and any(marker in paragraph_lower for marker in _BIBLIOGRAPHY_MARKERS)
            )
            if is_bibliography_header:
                bibliography_position = character_position
                break
            character_position += len(paragraph) + 1

        return full_text[:bibliography_position] if bibliography_position > 0 else full_text
```

Ported verbatim from `_build_text_sample()`. No constructor parameters — unlike
`QualityTextSampler`, the proposal does not flag these character limits as operator-tunable
(no `.env` exposure requested for classify-article's sampler in the proposal's scope or success
criteria), and they are tightly coupled to the externalized S4/S5/S6 prompt's own expectations
about sample shape (intro + conclusion split) the same way `analyze-quality`'s
non-`.env`-exposed paragraph-slicing constants are — changing them without re-tuning the prompt
risks degrading LLM output, so they stay as named module constants rather than constructor
parameters. This intentionally diverges from `QualityTextSampler`'s "everything is a constructor
parameter" ADR-7 — that decision was scoped to *that* class's specific tunables, not a blanket
rule that every sampler must expose every constant as a parameter; classify-article's proposal
contains no equivalent "make this tunable" instruction for these 2 constants.

**Rejected alternative**: mirror `QualityTextSampler` exactly and make
`introduction_character_limit`/`conclusion_character_limit` constructor parameters with
defaults. Rejected for this slice — would be speculative generality with no requirement driving
it (no `.env` var requested in the proposal's scope), and the proposal's own out-of-scope section
explicitly defers "accuracy improvements" and tuning to a future change; introducing tunability
not asked for risks silently becoming a quasi-feature. If a future slice needs this tunable, it
is a small, isolated follow-up.

---

## `ArticleClassificationResponseParser` — Full Design

```python
# src/domain/classification/article_classification_response_parser.py
import re

_S4_PATTERN = re.compile(r"S4\s*:\s*SI")
_S5_PATTERN = re.compile(r"S5\s*:\s*SI")
_S6_PATTERN = re.compile(r"S6\s*:\s*SI")


class ArticleClassificationResponseParser:
    """Parses the combined S4/S5/S6 LLM response into 3 boolean signals."""

    def parse(self, response_text: str) -> tuple[bool, bool, bool]:
        """Return (s4, s5, s6) extracted from the LLM's free-text yes/no response."""
        response_upper = response_text.strip().upper()
        s4 = bool(_S4_PATTERN.search(response_upper))
        s5 = bool(_S5_PATTERN.search(response_upper))
        s6 = bool(_S6_PATTERN.search(response_upper))
        return s4, s5, s6
```

Ported verbatim from `_signal_s4_s5_s6()`'s parsing half (the `try`/`except`-wrapped LLM call
itself moves to `ArticleClassifier`, since the parser's job is text-in/booleans-out with no
knowledge of the LLM port — same division of responsibility as
`QualityResponseParser.parse(response_text: str)` taking already-fetched text, never calling the
port itself). The 3 compiled patterns are pre-uppercased (`re.search` against an already-`.upper()`'d
string) rather than carrying `re.IGNORECASE` flags — matches the legacy's own `raw_upper =
raw.upper()` step exactly, kept as 3 separate module constants instead of inline `re.search(...)`
calls per call, for the same micro-optimization/clarity reason `QualityResponseParser`'s patterns
are module-level.

**No `ParsedResponseDTO`-equivalent needed here**: unlike `QualityResponseParser.parse()`, which
returns a structured `dict[QualityDimension, DimensionScoreDTO]` keyed by a 4-member enum (real
structural complexity justifying a DTO), this parser returns exactly 3 independent named
booleans with no natural grouping key — a bare `tuple[bool, bool, bool]` matches the legacy
return shape, is fully type-annotated, and introducing a 3-field DTO purely to wrap 3 unrelated
booleans would be the unjustified-abstraction smell call out in this project's own contributing
conventions ("3 lines repeated > 1 abstraction prematurely" applies equally to "3 plain values
> 1 wrapper DTO" here, since nothing downstream needs to pass this trio around as a unit beyond
immediate unpacking in `ArticleClassifier`).

**Rejected alternative**: return a `NamedTuple` (`Signal456Result` or similar) instead of a bare
tuple, for self-documenting field access (`.s4` instead of `[0]`). Considered, but the immediate
unpacking call site (`s4, s5, s6 = self._response_parser.parse(...)`) already names each value
at the point of use, making a NamedTuple's only benefit (named field access) redundant here.

---

## ADR-5: Exception Strategy

**Decision**: reuse `LanguageModelUnavailable` (already exists, `src/domain/exceptions/language_model_errors.py`)
for adapter-level Ollama failures — no change needed, `OllamaGeneratorAdapter` is shared as-is
between `analyze-quality` and `classify-article`. For unparseable/unusable LLM responses at the
domain level, reuse `ClassificationFailed` — **confirmed already present**,
`src/domain/exceptions/classification_errors.py`:

```python
# src/domain/exceptions/classification_errors.py (existing, unmodified)
from src.domain.exceptions.base_src_error import BaseSrcError


class ClassificationError(BaseSrcError):
    """Base class for all classification-related exceptions."""


class ClassificationFailed(ClassificationError):
    """Raised when article classification cannot be completed."""

    MESSAGE = "The article classification could not be completed."
```

`ArticleClassifier` raises `ClassificationFailed` if `document_content.paragraphs` is empty
(legacy's `classify_article()` raises a bare `ValueError("DocumentContent.paragraphs is empty")`
— migrated to the project's exception hierarchy instead of a built-in, consistent with how every
other migrated domain service in this codebase raises `BaseSrcError` subtypes, not built-ins).

**Rationale**: this mirrors `QualityAnalysisFailed`'s exact precedent (domain-level "the LLM
gave us something we can't use" signal) — discovered by reading `src/domain/exceptions/` before
deciding, which revealed `ClassificationFailed` was *already created* (likely scaffolded
speculatively in an earlier slice, or added defensively alongside `classification_errors.py`'s
sibling files during the exception-hierarchy slice). No new exception file needed; this design
only confirms reuse rather than introducing a duplicate
`ArticleClassificationFailed`/`UnparseableClassificationResponse` name.

**What actually triggers `ClassificationFailed` in `ArticleClassifier`**: unlike
`QualityAnalyzer`'s `_ensure_call_produced_usable_content` (which checks whether *any* of 2
expected dimensions appear in the parsed response), the S4/S5/S6 parser's `parse()` always
returns 3 well-defined booleans — there is no "unparseable" failure mode for this particular
response shape the way there is for `QualityResponseParser`'s richer per-dimension parsing
(a missing `S4:`/`S5:`/`S6:` line simply parses to `False`, exactly matching the legacy's
fail-safe `except Exception: return False, False, False` behavior). `ClassificationFailed` is
therefore reserved for the input-validation case (empty `paragraphs`), not a parse-failure case —
confirmed by re-reading legacy `_signal_s4_s5_s6()`'s own `except Exception as e: print(...);
return False, False, False` fallback: any LLM-call-level failure already degrades gracefully to
"all three signals absent" rather than raising, and this slice preserves that exact behavior
(parity is the goal) rather than upgrading it to raise — the only change is the *adapter's* own
LLM-unavailability path, where `LanguageModelUnavailable` propagates up through
`ArticleClassifier.classify()` uncaught, exactly as `QualityAnalyzer` already lets it propagate
uncaught from its own `generate()` calls.

**Rejected alternative**: catch `LanguageModelUnavailable` inside `ArticleClassifier` and
degrade to `(False, False, False)` the way the legacy's bare `except Exception` did. Rejected —
legacy's `except Exception` was catching *Ollama client-level* exceptions thrown synchronously
inside `_signal_s4_s5_s6`'s own `try` block, a responsibility this slice deliberately moves to
the adapter (`OllamaGeneratorAdapter` already raises `LanguageModelUnavailable` for backend
failures, confirmed in its existing source). Re-catching it in the domain service to silently
degrade would mean classify-article exhibits non-parity with `analyze-quality`'s own behavior
(`QualityAnalyzer` does not catch and degrade on `LanguageModelUnavailable` — it propagates),
introducing an inconsistency between the only two LLM-backed domain services in the codebase for
no documented reason; this is exactly the kind of "fixing a bug not in scope" the proposal's
"Out of Scope" section explicitly rules out (the legacy's print-and-degrade behavior on Ollama
failure is itself the kind of `print()`-based legacy scaffolding already excluded).

---

## `ArticleClassifier` — Full Design

```python
# src/domain/classification/article_classifier.py
from src.domain.classification.article_classification_response_parser import (
    ArticleClassificationResponseParser,
)
from src.domain.classification.article_classification_text_sampler import (
    ArticleClassificationTextSampler,
)
from src.domain.classification.imryd_signal_detector import ImrydSignalDetector
from src.domain.dtos.classification_result_dto import ClassificationResultDTO
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.enums.article_size import ArticleSize, classify_article_size
from src.domain.enums.article_type import ArticleType
from src.domain.enums.classification_confidence import ClassificationConfidence
from src.domain.exceptions.classification_errors import ClassificationFailed
from src.domain.ports.llm_generator_port import LlmGeneratorPort

_METHODOLOGICAL_VOCABULARY = (
    # Spanish — general methodology
    "metodología", "hipótesis", "variables", "variable dependiente",
    "variable independiente", "cuantitativo", "cualitativo", "mixto",
    "diseño de investigación", "diseño experimental", "cuasi-experimental",
    "correlación", "regresión", "análisis estadístico", "significancia",
    "validez", "confiabilidad", "encuesta", "entrevista",
    "observación sistemática", "triangulación", "marco metodológico",
    "población", "unidad de análisis", "categorías de análisis", "codificación",
    "datos primarios", "datos secundarios", "trabajo de campo",
    # ... (full bilingual list ported verbatim from legacy _METHODOLOGICAL_VOCAB,
    #      identical entries, omitted here for brevity — tasks phase copies verbatim)
)
_HARD_METHODOLOGICAL_TERMS = frozenset({
    "cuasi-experimental", "análisis estadístico", "triangulación",
    "marco metodológico", "unidad de análisis", "datos primarios",
    # ... (full set ported verbatim from legacy _HARD_TERMS)
})
_MINIMUM_REFERENCE_COUNT = 12
_RECENT_REFERENCE_YEAR_OFFSET = 4
_MINIMUM_RECENT_REFERENCE_RATIO = 0.5
_MINIMUM_VOCABULARY_TERM_COUNT = 4
_MINIMUM_HARD_TERM_COUNT = 1


class ArticleClassifier:
    """Domain service that classifies an academic article via a hybrid signal approach."""

    def __init__(
        self,
        llm_generator: LlmGeneratorPort,
        signal_detector: ImrydSignalDetector,
        text_sampler: ArticleClassificationTextSampler,
        response_parser: ArticleClassificationResponseParser,
        signal_prompt_template: str,
        temperature: float,
        num_predict: int,
    ) -> None:
        self._llm_generator = llm_generator
        self._signal_detector = signal_detector
        self._text_sampler = text_sampler
        self._response_parser = response_parser
        self._signal_prompt_template = signal_prompt_template
        self._temperature = temperature
        self._num_predict = num_predict

    def classify(self, document_content: DocumentContentDTO) -> ClassificationResultDTO:
        """Classify a document into an ArticleType with confidence and reasoning."""
        if not document_content.paragraphs:
            raise ClassificationFailed()

        article_size = classify_article_size(document_content.char_count)

        imryd_signals = self._signal_detector.detect(document_content)
        if imryd_signals["imryd_complete"] and article_size != ArticleSize.FUERA_RANGO:
            return ClassificationResultDTO.create(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=ClassificationConfidence.IMRYD_OVERRIDE,
                reasoning="Estructura IMRyD completa detectada (override determinístico).",
            )

        text_sample = self._text_sampler.build_sample(document_content)
        has_research_intent, has_evidence_based_contribution, has_theoretical_justification = (
            self._detect_research_intent_signals(text_sample, document_content.title)
        )
        signals = _ClassificationSignals(
            has_sufficient_reference_count=self._detect_sufficient_reference_count(document_content),
            has_recent_references=self._detect_recent_reference_majority(document_content),
            has_methodological_vocabulary=self._detect_methodological_vocabulary(document_content),
            has_research_intent=has_research_intent,
            has_evidence_based_contribution=has_evidence_based_contribution,
            has_theoretical_justification=has_theoretical_justification,
        )

        return self._apply_rule(signals, article_size)

    def _detect_sufficient_reference_count(self, document_content: DocumentContentDTO) -> bool:
        return len(document_content.references) >= _MINIMUM_REFERENCE_COUNT

    def _detect_recent_reference_majority(self, document_content: DocumentContentDTO) -> bool:
        # ported verbatim from _signal_reference_recency — regex year extraction,
        # >= 50% of references with max(year) >= current_year - 4
        ...

    def _detect_methodological_vocabulary(self, document_content: DocumentContentDTO) -> bool:
        # ported verbatim from _signal_methodological_vocab — NFD-normalize, count
        # matches against _METHODOLOGICAL_VOCABULARY, require >= 4 terms AND >= 1 hard term
        ...

    def _detect_research_intent_signals(
        self, text_sample: str, title: str | None
    ) -> tuple[bool, bool, bool]:
        """Return (has_research_intent, has_evidence_based_contribution, has_theoretical_justification)."""
        prompt = self._signal_prompt_template.format(title=title or "", text_sample=text_sample)
        response = self._llm_generator.generate(
            prompt, options={"temperature": self._temperature, "num_predict": self._num_predict}
        )
        return self._response_parser.parse(response)

    def _apply_rule(
        self, signals: "_ClassificationSignals", article_size: ArticleSize
    ) -> ClassificationResultDTO:
        # see ADR-6 (_RULE_TABLE dispatch) and ADR-7 (_ClassificationSignals) below for the
        # full table-driven implementation that replaces this section's earlier sketch
        ...
```

**Constructor signature rationale**: 7 parameters, none defaulted, mirroring the proposal's
explicit constraint ("no defaults in the domain service or adapter constructors") and
`QualityAnalyzer`'s own 5-parameter, all-required constructor shape. `temperature`/`num_predict`
are plain `float`/`int` (not wrapped in an `options` dict at the constructor boundary) — the dict
assembly happens only at the one `generate()` call site inside `_detect_research_intent_signals`,
keeping the constructor signature self-describing (`temperature: float, num_predict: int` reads
clearly; `options: dict` would hide what's actually required and lose static typing).

**Method naming**: legacy `_signal_reference_count`/`_signal_reference_recency`/
`_signal_methodological_vocab` become `_detect_sufficient_reference_count`/
`_detect_recent_reference_majority`/`_detect_methodological_vocabulary` — each renamed to
describe the boolean condition being tested (verb + condition), not just "which signal number
this is," consistent with `ImrydSignalDetector.detect()`'s naming and avoiding the legacy's
`s2a`/`s2b`/`s3` shorthand leaking into method names where it provides no information without
cross-referencing the rule table's comments. The `S2a`/`S2b`/`S3`/`S4`/`S5`/`S6` shorthand is
*kept* only inside `_apply_rule`'s local variable names and the ported Spanish reasoning strings
(matching the legacy's own reasoning text verbatim, which references "S2a", "S6" etc. by these
exact labels — changing them would alter user-visible output, which the proposal's parity
requirement forbids) and inside `ClassificationConfidence`'s doc comments/table above for
traceability back to the legacy case numbers.

**Why `classify()` not `classify_article()`**: the legacy method name `classify_article` is
redundant once this is a method on `ArticleClassifier` itself (`article_classifier.classify(...)`
reads identically to `article_classifier.classify_article(...)` but without repeating the class
name) — same naming-economy reasoning `QualityAnalyzer.analyze()` already applies (not
`QualityAnalyzer.analyze_quality()`).

**Why the full `_METHODOLOGICAL_VOCABULARY` tuple and rule-table `_apply_rule` body are elided
above with a verbatim-port comment, not transcribed in full here**: both are large
(`_METHODOLOGICAL_VOCAB` is ~70 string literals; `_apply_rule` is ~220 lines across 19 branches)
and pure verbatim data/logic ports with zero design decisions left to make — the proposal fixes
"ported verbatim, no reinterpretation" as a hard constraint, and the ADR-3 table above already
documents every behavior-relevant decision (the 4 confidence-value renames). Transcribing 290
combined lines of unchanged legacy text into this design document would not surface any
additional decision; the tasks phase copies both verbatim from
`business_logic/article_classifier.py` lines 15-55 (`_METHODOLOGICAL_VOCAB`) and 279-541
(`_apply_rule` body, signal variable renames only:
`s2a, s2b, s3, s4, s5, s6 = signals` keeps its exact legacy form since this destructuring,
unlike the signal-computation methods above, has no clearer name to give 6 already-well-labeled
local variables consumed entirely within one function body).

---

## `classify_article_size()` Placement

```python
# src/domain/enums/article_size.py
from enum import Enum


class ArticleSize(Enum):
    """Article size classification based on character count."""

    LARGO = "largo"  # 36,000 - 40,000 chars
    CORTO = "corto"  # 16,000 - 24,000 chars
    NO_DEFINIDO = "no_definido"  # 24,001 - 35,999 chars
    FUERA_RANGO = "fuera_rango"  # Outside all ranges


def classify_article_size(char_count: int) -> ArticleSize:
    """Classify article size based on character count with spaces."""
    if 36000 <= char_count <= 40000:
        return ArticleSize.LARGO
    if 16000 <= char_count <= 24000:
        return ArticleSize.CORTO
    if 24001 <= char_count <= 35999:
        return ArticleSize.NO_DEFINIDO
    return ArticleSize.FUERA_RANGO
```

Function appended to the existing `article_size.py`, exactly mirroring how
`get_quality_level_from_score()` is co-located with `QualityLevel` in `quality_level.py`. Logic
ported verbatim from legacy `domain/enums.py`'s `classify_article_size()` (confirmed identical
thresholds: 36000-40000 LARGO, 16000-24000 CORTO, 24001-35999 NO_DEFINIDO, else FUERA_RANGO) —
only the `if/elif/elif/else` chain is reformatted to early-return `if` statements (guard-clause
style, consistent with this project's "prefer early return over nested else" convention), which
is a pure control-flow restyling with identical branch outcomes for every input, not a logic
change.

---

## `ArticleClassifier` Bibliography Note — No `_build_text_sample` Duplication

The legacy `_build_text_sample` lives entirely inside `ArticleClassificationTextSampler` (see
above) — `ArticleClassifier` itself never touches paragraph-level bibliography-skipping logic,
calling only `self._text_sampler.build_sample(document_content)`. This is the same separation
`QualityAnalyzer` has with `QualityTextSampler.build_sample()`.

---

## Prompt Template File

```text
# src/infrastructure/resources/prompts/classification/s4_s5_s6_signal_prompt.txt
Analiza el siguiente fragmento de un artículo académico.

TÍTULO: {title}

TEXTO:
{text_sample}

TAREA: Responde TRES preguntas independientes sobre el texto.

PREGUNTA 1 — INTENCIÓN DE INVESTIGACIÓN (S4):
¿El artículo expresa explícitamente una intención de investigación mediante CUALQUIERA de estas formas?
- Verbos de intención: examinar, analizar, identificar, determinar, explorar, comprender, evaluar, investigar, estudiar, revisar, sintetizar
- Marcadores de alcance: "el presente estudio", "esta investigación", "la presente revisión", "el presente trabajo"
- Marcadores de problema: "el problema central", "el objetivo es", "la pregunta que guía", "se busca responder"
- Marcadores de propuesta experimental: "para fundamentar esta propuesta", "este trabajo combina", "este trabajo incluye", "a través de la simulación", "se propone demostrar", "se busca validar"
- Preguntas o hipótesis explícitas: una o múltiples, numeradas o no

PREGUNTA 2 — CONTRIBUCIÓN CONCLUSIVA (S5):
¿El artículo presenta conclusiones que exterioricen una contribución mediante CUALQUIERA de estas formas?
- Hallazgos de proceso sistemático: "los resultados demuestran", "la evidencia indica", "el análisis revela", "se identificaron"
- Propuesta de marco teórico, modelo, taxonomía o clasificación derivado del análisis
- Recomendaciones específicas derivadas de evidencia, no de opinión personal
- Identificación de brecha de conocimiento: "este estudio contribuye", "se propone", "se demuestra que"
- Síntesis que integra múltiples fuentes para arribar a una posición nueva
- Resultados experimentales cuantitativos: mejoras porcentuales, reducciones de tiempo, métricas de rendimiento
- Confirmación de hipótesis: "confirmando que", "lo que confirma", "los experimentos demostraron mejoras", "los resultados preliminares obtenidos fueron"

PREGUNTA 3 — JUSTIFICACIÓN TEÓRICA (S6):
¿El artículo justifica la selección de su marco teórico o identifica un vacío en el conocimiento existente que su investigación aborda mediante CUALQUIERA de estas formas?
- Referencia al estado del arte o literatura previa: "estudios previos han demostrado", "la literatura indica", "la literatura previa señala", "estudios previos muestran"
- Identificación de vacío: "sin embargo, no se ha explorado", "los estudios existentes no abordan", "existe un vacío en la literatura", "vacío en el conocimiento"
- Justificación del marco teórico: "se adopta el enfoque X porque", "este marco permite", "se seleccionó esta metodología porque"
- Anclaje en investigación previa: "a diferencia de estudios anteriores", "extendiendo el trabajo de", "en línea con"

FORMATO DE RESPUESTA — escribe ÚNICAMENTE estas tres líneas, sin encabezados, sin explicaciones:
S4: SI o NO
S5: SI o NO
S6: SI o NO
```

Verbatim from legacy's f-string, with `{title}` and `{text_sample}` retained as `.format()`
placeholders (2 placeholders here, vs. `analyze-quality`'s single `{text_sample}` — the
classification prompt embeds the title as legacy's f-string did, so both must survive the move
to a `.format()`-rendered template). `ArticleClassifier._detect_research_intent_signals` renders
via `self._signal_prompt_template.format(title=title or "", text_sample=text_sample)`,
defaulting `None` titles to an empty string the same way Python's f-string would have rendered
`{title}` as the literal text `"None"` only if `title` were the string `"None"` — actually: the
legacy f-string would render a `None` title as the literal text `"None"`, which is a latent
legacy quirk; the `title or ""` guard in this design is a deliberate, minimal divergence (not
parity-breaking for any realistic input, since every legacy call site always supplies a string
title) included only to avoid the new code crashing on `.format()` with a `None` value where an
f-string would have silently stringified it — `str.format()` and f-strings both call `str()`
on `None` to produce `"None"`, so this is actually NOT a divergence: `"{title}".format(title=None)`
produces `"None"`, identical to the f-string. Therefore **no `or ""` guard is needed** —
`_detect_research_intent_signals` renders with the title value as-is, preserving exact legacy
behavior for the (extremely unlikely, never-exercised) `None`-title case as well.

```python
# src/infrastructure/resources/prompts/classification/__init__.py
from os import path

PROMPTS_DIR = path.dirname(__file__)
```

Identical pattern to `src/infrastructure/resources/prompts/quality/__init__.py`.

---

## `ClassifyArticleUseCase` — Full Design

```python
# src/application/classify_article_use_case.py
from src.domain.classification.article_classifier import ArticleClassifier
from src.domain.dtos.classification_result_dto import ClassificationResultDTO
from src.domain.dtos.document_content_dto import DocumentContentDTO


class ClassifyArticleUseCase:
    def __init__(self, classifier: ArticleClassifier) -> None:
        self._classifier = classifier

    def execute(self, document_content: DocumentContentDTO) -> ClassificationResultDTO:
        return self._classifier.classify(document_content)
```

Identical shape to `AnalyzeQualityUseCase` — single constructor-injected domain collaborator,
one-line delegation in `execute()`. No `article_type` parameter (unlike
`AnalyzeQualityUseCase.execute()`, which carries the legacy's unused/undocumented
`article_type` parameter) — confirmed by reading legacy `classify_article(document_content)`'s
own signature, which never received an `article_type` argument; classification *produces*
`ArticleType`, it does not consume one, so there is no equivalent dead-parameter precedent to
carry forward here.

---

## `ClassifyArticleUseCaseWiring` — Full Design

```python
# src/infrastructure/wirings/classify_article_use_case_wiring.py
from os import getenv
from os.path import join
from pathlib import Path

from dotenv import load_dotenv

from src.application.classify_article_use_case import ClassifyArticleUseCase
from src.domain.classification.article_classification_response_parser import (
    ArticleClassificationResponseParser,
)
from src.domain.classification.article_classification_text_sampler import (
    ArticleClassificationTextSampler,
)
from src.domain.classification.article_classifier import ArticleClassifier
from src.domain.classification.imryd_signal_detector import ImrydSignalDetector
from src.domain.ports.llm_generator_port import LlmGeneratorPort
from src.infrastructure.adapters.llm_generator.ollama_generator_adapter import (
    OllamaGeneratorAdapter,
)
from src.infrastructure.resources.prompts.classification import PROMPTS_DIR

load_dotenv()


class ClassifyArticleUseCaseWiring:
    """Factory for building a ready-to-use ClassifyArticleUseCase."""

    def create_use_case(self) -> ClassifyArticleUseCase:
        return ClassifyArticleUseCase(classifier=self._get_article_classifier())

    def _get_article_classifier(self) -> ArticleClassifier:
        return ArticleClassifier(
            llm_generator=self._get_llm_generator(),
            signal_detector=ImrydSignalDetector(),
            text_sampler=ArticleClassificationTextSampler(),
            response_parser=ArticleClassificationResponseParser(),
            signal_prompt_template=self._read_prompt_template("s4_s5_s6_signal_prompt.txt"),
            temperature=float(getenv("ARTICLE_CLASSIFIER_TEMPERATURE", "0.1")),
            num_predict=int(getenv("ARTICLE_CLASSIFIER_NUM_PREDICT", "300")),
        )

    def _get_llm_generator(self) -> LlmGeneratorPort:
        model_name = getenv("OLLAMA_MODEL_NAME", "llama3-gradient:8b-instruct-1048k-q4_K_M")
        base_url = getenv("OLLAMA_BASE_URL", "http://localhost:11434")
        return OllamaGeneratorAdapter(model_name=model_name, base_url=base_url)

    def _read_prompt_template(self, filename: str) -> str:
        file_path = Path(join(PROMPTS_DIR, filename))
        return file_path.read_text(encoding="utf-8")
```

**Decision**: `_get_llm_generator()` is duplicated verbatim from
`AnalyzeQualityUseCaseWiring._get_llm_generator()` rather than extracted into a shared
`OllamaGeneratorAdapterFactory` or similar. **Rejected alternative**: extract a shared
factory/mixin used by both wirings. Rejected for this slice — 2 occurrences of a 3-line method
is below this project's stated duplication tolerance ("3 lines repeated > 1 abstraction
prematurely"), and the proposal's scope does not request a wiring-level shared abstraction;
revisit only if a 3rd LLM-backed wiring appears and the duplication becomes a real
maintenance burden (Rule of Three).

`temperature`/`num_predict` read from `ARTICLE_CLASSIFIER_TEMPERATURE`/
`ARTICLE_CLASSIFIER_NUM_PREDICT`, defaulting to the legacy's exact hardcoded
`{"temperature": 0.1, "num_predict": 300}` values — `float()`/`int()` casts applied for the same
reason `AnalyzeQualityUseCaseWiring` casts its own env values (`os.getenv()` always returns
`str | None`, an uncast string would silently break the dict passed to `ollama.generate()`,
which expects numeric option values).

---

## `.env.example` Addition

```diff
 # Ollama backend connection (OllamaGeneratorAdapter)
 OLLAMA_MODEL_NAME=llama3-gradient:8b-instruct-1048k-q4_K_M
 OLLAMA_BASE_URL=http://localhost:11434

 # Quality analysis text sampling (QualityTextSampler)
 QUALITY_MIN_SAMPLE_WORD_COUNT=400
 QUALITY_TEXT_SAMPLE_CHARACTER_LIMIT=8000
+
+# Article classification signal-extraction tuning (ArticleClassifier)
+ARTICLE_CLASSIFIER_TEMPERATURE=0.1
+ARTICLE_CLASSIFIER_NUM_PREDICT=300
```

Values match legacy's exact hardcoded tuning (`temperature=0.1` for low-variance yes/no
extraction, `num_predict=300` bounding response length) — operator copying this file unmodified
to `.env` reproduces current legacy behavior exactly, same guarantee `analyze-quality`'s
`.env.example` block already provides for its own 2 variables.

---

## Test File Layout — One `TestCase` Class Per File, No Exception for Fakes

Following this project's strict convention (confirmed via `src/domain/tests/structure/`'s
4-file-per-`StructureValidator` split and `src/domain/tests/quality/`'s
sampler/parser/analyzer 3-file split — no test file in either folder contains more than one
`TestCase` subclass):

| File | Scope |
|---|---|
| `src/domain/tests/classification/fake_llm_generator_port.py` | Test double — **not** a `TestCase`, mirrors `src/domain/tests/quality/fake_llm_generator_port.py`'s exact shape (`FakeLlmGeneratorPort(responses: list[str])`, tracks `call_count`/`received_prompts`); lives alongside the test files it serves, not under `src/application/tests/` or `src/infrastructure/tests/`, matching where the Slice-5 equivalent lives |
| `test_imryd_signal_detector.py` | All 5 keyword signals + `imryd_complete` true/false combinations (bilingual keyword matching, ≤5-word header filter) |
| `test_article_classification_text_sampler.py` | Bibliography-skip detection, intro+ending concatenation, fallback-to-full-text path |
| `test_article_classification_response_parser.py` | S4/S5/S6 regex extraction — present/absent/mixed-case combinations |
| `test_article_classifier_imryd_override.py` | Case 1 — IMRyD-complete + size != `FUERA_RANGO` short-circuits to `CIENTIFICO`/`IMRYD_OVERRIDE`; IMRyD-complete + `FUERA_RANGO` does NOT short-circuit |
| `test_article_classifier_cientifico.py` | Cases 2-5 — all 4 `CIENTIFICO` non-override branches, one test per confidence member |
| `test_article_classifier_divulgacion_near_miss.py` | Cases 6-9 — S3+S4+S5 present but below the 0.83 threshold |
| `test_article_classifier_divulgacion_standard.py` | Cases 10-18 — missing one or more of S3/S4/S5 |
| `test_article_classifier_opinion.py` | Case 19 — no signals detected |
| `src/domain/tests/enums/test_classification_confidence.py` | 5 members exist, each is a `float` instance, values match the table above |
| `src/domain/tests/enums/test_classify_article_size.py` | Boundary values for all 4 `ArticleSize` thresholds (mirrors `test_get_quality_level_from_score.py`'s boundary-testing style) |
| `src/application/tests/test_classify_article_use_case.py` | `execute()` delegates to `classifier.classify()`, same assertion shape as `test_analyze_quality_use_case.py` |
| `src/infrastructure/tests/test_classify_article_use_case_wiring.py` | `create_use_case()` returns correct type; `_get_llm_generator` return-type hint check — same 2-test shape as `test_analyze_quality_use_case_wiring.py` |
| `src/infrastructure/tests/test_ollama_generator_adapter.py` | **Modified, not new** — adds cases asserting `options` forwards to `ollama.generate(options=...)` verbatim, and that omitting `options` preserves prior (Slice 5) behavior unchanged |

**Why 5 separate `test_article_classifier_*.py` files instead of one** (the proposal's own
language anticipates this — "finer than `validate-structure`'s per-`ArticleType` test-file
split, given classification has 19 distinct cases vs. structure-validation's 4"): grouping by
outcome category (override / cientifico-confidence-tiers / divulgacion-near-miss /
divulgacion-standard / opinion) rather than one-file-per-case (19 files) balances the "one
`TestCase` class per file" convention against practicality — 19 separate single-test files would
fragment a single rule table's branches across more files than the table itself has meaningful
groupings; the 5 chosen groupings match the legacy code's own comment-delimited sections
(`# ── CIENTÍFICO paths`, `# ── Near-miss`, `# ── DIVULGACIÓN standard`, `# ── OPINIÓN`),
so each test file's scope is traceable to a documented section of the source it tests.

**`fake_llm_generator_port.py` exception note**: this file is explicitly *not* a `TestCase`
subclass and is the one deliberate, already-established exception to "one TestCase class per
file" — it contains zero test classes, only a plain test-double class, exactly matching
`src/domain/tests/quality/fake_llm_generator_port.py`'s precedent (the proposal/convention's
"strict... no exception for fakes" instruction is interpreted here as: fakes still get their own
dedicated file (not bundled into a test file), not that fakes must themselves be `TestCase`
subclasses — confirmed by reading the existing `quality/fake_llm_generator_port.py`, which the
codebase already ships as a non-`TestCase` file).

---

## ADR-6: Declarative Rule Table Replaces the `_apply_rule` If/Elif Tree

**Decision**: replace the legacy 19-case `if`/`elif` tree with an ordered tuple of `_RuleCase`
entries, matched top-to-bottom by a single dispatch loop — the same "ordered tuple + first-match
loop" shape already established by `quality_response_parser.py`'s `_DIMENSION_KEYWORDS` and
`_NARRATIVE_SCORE_KEYWORDS`.

```python
# src/domain/classification/article_classifier.py
from dataclasses import dataclass
from typing import Callable

@dataclass(frozen=True)
class _RuleCase:
    """One row of the 19-case classification rule table."""
    predicate: Callable[["_ClassificationSignals"], bool]
    article_type: ArticleType
    confidence: ClassificationConfidence | None
    reasoning: Callable[["_ClassificationSignals", str, str], str]  # (signals, active_text, inactive_text) -> str

_FULL_CORE = lambda sig: sig.has_methodological_vocabulary and sig.has_research_intent and sig.has_evidence_based_contribution

_RULE_TABLE: tuple[_RuleCase, ...] = (
    # ── CIENTÍFICO: full core (S3+S4+S5) + structural support, confidence >= 0.83 ──
    _RuleCase(lambda s: _FULL_CORE(s) and s.has_sufficient_reference_count and s.has_recent_references and s.has_theoretical_justification,
               ArticleType.CIENTIFICO, ClassificationConfidence.FULL_SIGNAL_MATCH, _reasoning_case_2),
    _RuleCase(lambda s: _FULL_CORE(s) and s.has_recent_references and s.has_theoretical_justification,
               ArticleType.CIENTIFICO, ClassificationConfidence.RECENT_BIBLIOGRAPHY_SUPPORT, _reasoning_case_3),
    _RuleCase(lambda s: _FULL_CORE(s) and s.has_sufficient_reference_count and s.has_recent_references,
               ArticleType.CIENTIFICO, ClassificationConfidence.COMPLETE_BIBLIOGRAPHY_SUPPORT, _reasoning_case_4),
    _RuleCase(lambda s: _FULL_CORE(s) and s.has_sufficient_reference_count and s.has_theoretical_justification,
               ArticleType.CIENTIFICO, ClassificationConfidence.SUFFICIENT_REFERENCE_COUNT, _reasoning_case_5),
    # ── DIVULGACIÓN near-miss: full core present, below 0.83 threshold (cases 6-9) ──
    _RuleCase(lambda s: _FULL_CORE(s) and s.has_theoretical_justification, ArticleType.DIVULGACION, None, _reasoning_case_6),
    _RuleCase(lambda s: _FULL_CORE(s) and s.has_recent_references, ArticleType.DIVULGACION, None, _reasoning_case_7),
    _RuleCase(lambda s: _FULL_CORE(s) and s.has_sufficient_reference_count, ArticleType.DIVULGACION, None, _reasoning_case_8),
    _RuleCase(_FULL_CORE, ArticleType.DIVULGACION, None, _reasoning_case_9),
    # ── DIVULGACIÓN standard (cases 10-18) ──
    _RuleCase(lambda s: s.has_methodological_vocabulary and s.has_research_intent, ArticleType.DIVULGACION, None, _reasoning_case_10),
    _RuleCase(lambda s: s.has_methodological_vocabulary and s.has_evidence_based_contribution, ArticleType.DIVULGACION, None, _reasoning_case_11),
    _RuleCase(lambda s: s.has_methodological_vocabulary and s.has_sufficient_reference_count and s.has_recent_references, ArticleType.DIVULGACION, None, _reasoning_case_12),
    _RuleCase(lambda s: s.has_methodological_vocabulary and s.has_sufficient_reference_count, ArticleType.DIVULGACION, None, _reasoning_case_13),
    _RuleCase(lambda s: s.has_methodological_vocabulary and s.has_recent_references, ArticleType.DIVULGACION, None, _reasoning_case_14),
    _RuleCase(lambda s: s.has_methodological_vocabulary, ArticleType.DIVULGACION, None, _reasoning_case_15),
    _RuleCase(lambda s: s.has_research_intent and s.has_evidence_based_contribution, ArticleType.DIVULGACION, None, _reasoning_case_16),
    _RuleCase(lambda s: s.has_research_intent, ArticleType.DIVULGACION, None, _reasoning_case_17),
    _RuleCase(lambda s: s.has_evidence_based_contribution, ArticleType.DIVULGACION, None, _reasoning_case_18),
    # case 19 has no row — it is the loop's fallback, not a table entry (see rationale)
)

def _apply_rule(self, signals: "_ClassificationSignals", article_size: ArticleSize) -> ClassificationResultDTO:
    active, inactive = self._describe_signals(signals)
    for case in _RULE_TABLE:
        if case.predicate(signals):
            return ClassificationResultDTO.create(
                article_type=case.article_type, article_size=article_size,
                confidence=case.confidence, reasoning=case.reasoning(signals, active, inactive),
            )
    return ClassificationResultDTO.create(
        article_type=ArticleType.OPINION, article_size=article_size, confidence=None,
        reasoning=_reasoning_case_19(signals, active, inactive),
    )
```

**How the gate (S3∧S4∧S5) plus sub-condition is represented**: each of cases 2-9's predicates is
a lambda that calls the shared `_FULL_CORE(s)` helper (itself `s3 and s4 and s5`) **and**-ed with
the case-specific extra condition — this mirrors the legacy's own structure exactly (one outer
`if s3 and s4 and s5:` gate, with cases 2-5 as nested `if`/`elif` and 6-9 as the `if`/`elif`/`else`
fallthrough inside that same gate). The table format doesn't need a special "gated group" concept;
repeating `_FULL_CORE(s) and ...` in 8 predicates is the literal translation of "nested under one
gate," and `_FULL_CORE` is itself a one-line named lambda so the repetition reads as "still under
the S3+S4+S5 umbrella" at a glance, not 8 independent unrelated conditions.

**How case 19 (the fallback) is represented**: it is **not a table row** — it is the loop's
`else`-equivalent, returned after the `for` loop exhausts the table without a match. This mirrors
`QualityResponseParser._map_block_to_dimension`'s own shape (loop returns `None` if nothing
matches; the caller decides what `None` means) more closely than forcing a synthetic
"always-True" row into the table, which would misrepresent case 19 as a *condition* when it is
actually "absence of any condition" — a structurally different thing from rows 1-18, each of
which IS a specific signal combination.

**Reasoning-string functions, not f-strings inline in the table**: each case's reasoning text is
a small module-level function (`_reasoning_case_2`, etc.) taking `(signals, active_text,
inactive_text)` and returning the **exact, unmodified** legacy Spanish string (including the
`_sig(active, inactive)` suffix, ported as the `_describe_signals` helper producing the same
`"Señales presentes: ..."`/`"Señales ausentes: ..."` text). Functions, not table-embedded string
templates, because several reasoning strings are not simple `.format()` substitutions — case 2's
text is fully static prose with only the trailing signal-summary appended, so a function avoids
inventing a templating mini-language for what is, in 18 of 19 cases, just "static string + common
suffix."

**Order preservation**: the tuple's row order is **exactly** the legacy's branch-check order
(case 2 before 3 before 4 ... before 18), so `for case in _RULE_TABLE: if case.predicate(signals):
return ...` reproduces first-match-wins semantics identically — case 10 is only reached if none of
rows 1-9 (cases 2-9) matched, exactly as legacy's `if s3 and s4 and s5:` block being skipped (or
falling through without an early return) lets control reach the `if s3 and s4:` line below it.

**Rejected alternative**: keep the if/elif tree. Rejected — explicit user instruction; "it doesn't
scale" was the stated reason, and the codebase already has a working precedent
(`quality_response_parser.py`) for replacing exactly this shape of branching logic.

**Rejected alternative**: a `dict[frozenset[str], _RuleCase]` keyed by the frozenset of
required-true signal names. Rejected — doesn't naturally express "match the LARGEST satisfied
subset first" (case 2's frozenset is a superset of case 5's; a plain dict lookup can't express
"prefer the most specific match," only exact-key lookup), and the near-miss cases (6-9) aren't
expressible as a clean frozenset at all since they're defined by exclusion ("S3∧S4∧S5 present,
AND none of cases 2-5 matched") rather than a fixed signal set. The ordered-tuple-of-predicates
shape handles "first satisfying row wins" natively, which is what legacy's branch order actually
encodes; a dict would need an additional explicit ranking mechanism on top to recover the same
behavior, adding complexity instead of removing it.

**Rejected alternative**: a `match`/`case` statement (Python 3.10+ structural pattern matching).
Rejected — `match` shines when branching on a value's *shape* (type, sequence pattern); here
every branch is a boolean conjunction over named flags, which `match` would need to express via
`case _ if predicate:` guards anyway — functionally identical to the chosen ordered-tuple loop but
without the benefit of the table being a first-class, separately inspectable/testable data
structure (e.g. `len(_RULE_TABLE) == 18` is a meaningful sanity-check assertion the tasks phase
can write; an equivalent assertion against a `match` block's branch count is not expressible the
same way).

---

## ADR-7: Self-Explanatory Internal Signal Naming via a `_ClassificationSignals` Dataclass

**Decision**: replace the legacy `list[bool]`/positional-tuple-unpacking shorthand (`s2a, s2b,
s3, s4, s5, s6`) with a frozen dataclass of named booleans, used internally by `ArticleClassifier`
and the rule table. Output DTO and reasoning text are unaffected — purely an internal
readability change.

```python
# src/domain/classification/article_classifier.py
@dataclass(frozen=True)
class _ClassificationSignals:
    """Named replacement for legacy's positional s2a/s2b/s3/s4/s5/s6 tuple."""
    has_sufficient_reference_count: bool   # legacy s2a — >= 12 references
    has_recent_references: bool             # legacy s2b — >= 50% references recent
    has_methodological_vocabulary: bool     # legacy s3  — >= 4 vocab terms incl. >= 1 hard term
    has_research_intent: bool               # legacy s4  — LLM: explicit research intent
    has_evidence_based_contribution: bool   # legacy s5  — LLM: evidence-based conclusive contribution
    has_theoretical_justification: bool     # legacy s6  — LLM: theoretical framework / knowledge-gap
```

**Final names** (consistency over individual preference, per the user's own instruction):
`has_sufficient_reference_count` (s2a), `has_recent_references` (s2b),
`has_methodological_vocabulary` (s3), `has_research_intent` (s4),
`has_evidence_based_contribution` (s5), `has_theoretical_justification` (s6). All 6 share the
`has_*` prefix (matching `ImrydSignalDetector`'s own `has_introduction`/`has_methods`/etc.
convention from ADR-1/2 — one consistent boolean-naming style across both signal sources in this
slice) and each name states the condition being tested, not a signal ID.

**Why a frozen dataclass, not a dict with these names as keys**: the rule table's predicates
access signals via attribute (`s.has_methodological_vocabulary`), which a dataclass gives for
free with static type-checking and typo-detection at lint/type-check time; a `dict[str, bool]`
would require `s["has_methodological_vocabulary"]` (or `.get()` with a default, masking a real
typo as a silent `False`). `frozen=True` documents that signals are computed once per
classification call and never mutated afterward — matching how the legacy's own `signals` list
is built once and only read by `_apply_rule`. `ImrydSignalDetector.detect()`'s return type stays a
plain `dict[str, bool]` (ADR-1, unchanged) because that dict is a small, already-flat detector
output consumed as `imryd_signals["imryd_complete"]` exactly once per call (no rule-table
predicate ever iterates or destructures it) — there's no parity reason to convert that one too.

**Rejected alternative**: a `NamedTuple` instead of a dataclass. Rejected — functionally
near-identical, but `frozen=True` dataclasses read more idiomatically as "immutable value object"
in this codebase, matching `DimensionScoreDTO`/other DTOs already defined as `@dataclass(frozen=True)`,
whereas no existing internal (non-DTO) value type in `src/domain/` currently uses `NamedTuple`.

**Rejected alternative**: keep `s2a`...`s6` as the dataclass's own field names (rename only the
*type*, not the fields). Rejected — defeats the purpose; the whole point raised in review is that
`s2a`/`s3`/etc. carry zero information without cross-referencing the rule table's comments. The
`# legacy s2a` comments above exist precisely so a reader can still map back to the rule-table
spec/legacy code when needed, without the field names themselves being cryptic.

---

## ADR-8: Shared `read_text_resource()` Helper — Infrastructure Utility, Not a Port

**Decision**: extract a tiny module-level function, used by both wirings, replacing each
wiring's private `_read_prompt_template` method:

```python
# src/infrastructure/resources/text_resource_loader.py
from os.path import join
from pathlib import Path


def read_text_resource(directory: str, filename: str) -> str:
    """Read a UTF-8 text resource file from the given directory."""
    return Path(join(directory, filename)).read_text(encoding="utf-8")
```

`ClassifyArticleUseCaseWiring` calls `read_text_resource(PROMPTS_DIR, "s4_s5_s6_signal_prompt.txt")`
directly (no `self._read_prompt_template` wrapper method) at its `_get_article_classifier()` call
site. **`AnalyzeQualityUseCaseWiring` is retrofitted in this same slice**: its private
`_read_prompt_template` method is deleted, its 2 call sites become
`read_text_resource(PROMPTS_DIR, "clarity_coherence_prompt.txt")` and the argumentation
equivalent.

**Why retrofit Slice 5's already-merged wiring rather than leave it alone**: confirmed via
`src/infrastructure/tests/test_analyze_quality_use_case_wiring.py` that `_read_prompt_template`
has **zero direct test coverage** today (only `create_use_case()` and `_get_llm_generator`'s type
hint are tested) — so deleting the private method and inlining the shared function call is a
behavior-invisible change with no test to update or break. Leaving Slice 5's copy in place while
adding a second copy for Slice 6 would mean the exact duplication the user flagged exists for as
long as nobody revisits Slice 5 — better to collapse it now, while both call sites are in the
same person's working memory, than to defer and risk a 3rd near-identical copy appearing in a
future slice before anyone notices the pattern.

**Why this is explicitly NOT a port** (re-confirmed from the earlier decision in this
conversation, restated here for the design doc's own completeness): no domain boundary is
crossed — the domain layer never calls this function or knows it exists; `ArticleClassifier` and
`QualityAnalyzer` both receive an already-loaded `str` (`signal_prompt_template`,
`clarity_coherence_prompt_template`, etc.) via constructor injection. There is no real driver for
swapping implementations at runtime (unlike `LlmGeneratorPort`, which has two meaningfully
different real-world backends to abstract over — Ollama today, potentially a cloud LLM
tomorrow — a text-file read has exactly one implementation anyone has ever needed: read a file
from disk). A `Protocol` + adapter pair here would be ceremony with no abstraction payoff,
contradicting this project's own "no abstractions, helpers, or features not asked for" convention
in spirit (the inverse failure mode: *over*-abstracting a non-varying file read into a fake port).

**File location rationale**: `src/infrastructure/resources/text_resource_loader.py` — sibling to
`src/infrastructure/resources/prompts/`, since both prompt packages (`quality/`, `classification/`)
are the resources this loader exists to serve; `wirings/` was rejected as the location because the
function has no wiring-specific logic (no env vars, no factory behavior) and living under
`resources/` keeps "what is a resource" and "how resources are loaded" co-located.

**Rejected alternative**: a shared base class (`BaseUseCaseWiring`) with `_read_prompt_template`
as a protected method, inherited by both wirings. Rejected — introduces an inheritance
relationship between two otherwise-unrelated wirings purely to share a 2-line method; a plain
function import achieves the same de-duplication with composition instead of inheritance, and
doesn't risk future unrelated wiring-specific state/behavior leaking across the new shared base
as both wirings evolve independently.

---

## ADR-9: `tests/smoke/test_classify_article_parity.py` — Both Sides' LLM Calls Mocked

**Decision**: follow `test_validate_structure_parity.py`'s exact `TestCase` + `setUpClass`
shape, against the same 3 real `.docx` files, but patch **both** sides' LLM call so the test never
touches a live Ollama instance.

```python
# tests/smoke/test_classify_article_parity.py
# ruff: noqa: E402
import sys
from pathlib import Path
from unittest import TestCase
from unittest.mock import patch

ROOT = Path(__file__).parent.parent.parent
sys.path.insert(0, str(ROOT))

from data_access.word_reader import WordReader
from business_logic.article_classifier import ArticleClassifier as LegacyClassifier
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.infrastructure.wirings.classify_article_use_case_wiring import (
    ClassifyArticleUseCaseWiring,
)

DOCS = ROOT / "docs" / "sample-documents"
_CANNED_RESPONSE = {"response": "S4: SI\nS5: SI\nS6: SI"}
_DOCUMENTS = ["1. test_Científico.docx", "2. test_divulgacion_v2.docx", "3. test_opinion_v2.docx"]


class TestClassifyArticleParity(TestCase):
    @classmethod
    def setUpClass(cls):
        cls.reader = WordReader()
        cls.legacy = LegacyClassifier()
        cls.use_case = ClassifyArticleUseCaseWiring().create_use_case()

    def _run(self, filename: str):
        paragraphs = self.reader.read_word_document(str(DOCS / filename))
        doc = DocumentContentDTO(
            word_count=sum(len(p.split()) for p in paragraphs),
            char_count=sum(len(p) for p in paragraphs),
            paragraphs=paragraphs,
        )
        with patch(
            "business_logic.article_classifier.ollama.Client.generate",
            return_value=_CANNED_RESPONSE,
        ):
            legacy_result = self.legacy.classify_article(doc)
        with patch(
            "src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.generate",
            return_value=_CANNED_RESPONSE,
        ):
            new_result = self.use_case.execute(doc)
        return legacy_result, new_result

    def test_cientifico_parity(self):
        legacy, new = self._run(_DOCUMENTS[0])
        self.assertEqual(new.article_type.value, legacy.article_type.value)
        self.assertEqual(new.confidence, legacy.confidence)

    def test_divulgacion_parity(self):
        legacy, new = self._run(_DOCUMENTS[1])
        self.assertEqual(new.article_type.value, legacy.article_type.value)
        self.assertEqual(new.confidence, legacy.confidence)

    def test_opinion_parity(self):
        legacy, new = self._run(_DOCUMENTS[2])
        self.assertEqual(new.article_type.value, legacy.article_type.value)
        self.assertEqual(new.confidence, legacy.confidence)
```

**Exact patch targets, confirmed against actual source (not assumed)**: legacy's `ArticleClassifier.__init__`
sets `self.client = ollama.Client(host=base_url)` and the S4/S5/S6 call is
`self.client.generate(model=..., prompt=..., options=...)` — an **instance method on
`ollama.Client`**, not the module-level `ollama.generate` function. So the legacy-side patch
target is `business_logic.article_classifier.ollama.Client.generate` (patching the `Client`
class's `generate` method affects the `self.client` instance constructed in `setUpClass`,
since `self.client` is an instance of that same patched class). The new-side patch target is
`src.infrastructure.adapters.llm_generator.ollama_generator_adapter.ollama.generate` — confirmed
from `OllamaGeneratorAdapter.generate()`'s actual call `ollama.generate(model=self._model_name,
prompt=prompt, options=options)`, the **module-level** function, a different call shape from
legacy's `Client`-instance method. These are NOT the same patch target — conflating them would
silently leave one side's classification driven by a real Ollama call.

**What gets asserted**: `article_type` (compared via `.value` since legacy's `ArticleType` and
the new `src.domain.enums.article_type.ArticleType` are separate enum classes with matching
string values, mirroring `test_validate_structure_parity.py`'s own type-boundary handling) and
`confidence` (compared directly — legacy returns a bare `float | None`; new returns
`ClassificationConfidence | None`, but `ClassificationConfidence(float, Enum)` compares equal to
its underlying float via `Enum`'s `__eq__`, so `assertEqual` passes without needing `.value`).
`reasoning` is deliberately NOT asserted in this smoke test — `ArticleClassificationParity`'s
domain-level unit tests (ADR-6's `_RULE_TABLE`, one per case) already assert exact reasoning-text
parity per case; re-asserting it here against real (canned-LLM) documents would duplicate that
coverage without adding confidence, and risks false failures if a sample `.docx`'s real deterministic
signals (s2a/s2b/s3) land on a case boundary sensitive to environment-specific text extraction
differences not worth chasing in a smoke test whose unique value is "real `.docx` parsing
end-to-end," not exhaustive case coverage (already owned by ADR-6's table-driven unit tests).

**Why real `.docx` parsing is still exercised despite mocking the LLM call**: `WordReader` and
`document_content` construction are untouched by either patch — only the S4/S5/S6 signal
extraction's network call is faked. The deterministic signals (IMRyD override, S2a/S2b/S3) still
run against the real, parsed document text on both sides, which is exactly the coverage this test
exists to provide beyond what fake-port domain unit tests already give.

**Rejected alternative**: patch at the `ArticleClassifier`/`OllamaGeneratorAdapter` method level
(`patch.object(LegacyClassifier, "_signal_s4_s5_s6", return_value=(True, True, True))` and
similarly for the new side). Rejected — patching one level higher (the LLM call itself, not the
method that calls it) keeps the test exercising each side's own response-parsing logic
(`_signal_s4_s5_s6`'s regex extraction; `ArticleClassificationResponseParser.parse()`), which is
real production code worth covering end-to-end in this smoke test rather than bypassing.

---

## Risks / Open Items for Tasks Phase

- `_METHODOLOGICAL_VOCAB`/`_HARD_TERMS` (~70 + ~30 string literals) must be copied
  character-for-character from `business_logic/article_classifier.py` lines 15-55 — tasks phase
  should treat this as a literal copy-paste step with a diff-against-source verification step,
  not a re-transcription, to eliminate transcription-error risk (the proposal's top risk).
- `OllamaGeneratorAdapter`'s actual current source (confirmed during this design read) already
  diverges from `analyze-quality`'s own design doc snapshot (narrower `except` clause) — tasks
  phase must diff against the **current real file**, not the Slice-5 design doc, when applying
  ADR-4's `options` parameter change.
- `ClassificationFailed` already exists in `src/domain/exceptions/classification_errors.py` —
  tasks phase must NOT create a duplicate; only the new `ArticleClassifier.classify()` call site
  that raises it is new.
- Confirm during tasks/apply that `document_content.references` defaults to `[]` (not `None`)
  per `DocumentContentDTO`'s `field(default_factory=list)` — legacy's
  `references = document_content.references or []` guard is therefore unnecessary on the
  DTO-backed version and should be simplified to a direct `len(document_content.references)`,
  a small, justified deviation from verbatim-port (eliminating dead defensive code against a
  `None` value the DTO's own dataclass field guarantees can't occur), consistent with how this
  migration has already trimmed other DTO-guaranteed null-checks in prior slices.
- ADR-6's `_RULE_TABLE` row order and each `_reasoning_case_N` function's string content must be
  diff-verified character-for-character against legacy `_apply_rule` (lines 279-541) during
  tasks/apply — the table restructuring is a pure refactor of *dispatch mechanism*, not logic;
  any drift in a reasoning string or a predicate's signal combination is a parity regression.
- ADR-8's retrofit touches `AnalyzeQualityUseCaseWiring` (already merged from Slice 5) — tasks
  phase must re-run Slice 5's existing wiring test
  (`test_analyze_quality_use_case_wiring.py`) after the retrofit to confirm no regression, even
  though `_read_prompt_template` itself has no dedicated test today.
- ADR-9's smoke test patch targets are call-shape-sensitive (`ollama.Client.generate` instance
  method on the legacy side vs. module-level `ollama.generate` on the new side) — tasks/apply
  must verify the patch actually intercepts the call (e.g. via the mock's `call_count`) rather
  than silently falling through to a real Ollama call if either source file's call shape changes
  before this slice lands.

---

## Amendment (PR-1 code review, post-implementation)

User code review of PR-1 (Phases 0-7) surfaced 3 design corrections, applied directly to the
implemented files (design snippets above are now stale on these points — kept for historical
record, not to be copied verbatim by future tasks):

**1. `classify_article_size()` violates the project's full-OOP convention** (decision recorded
2026-06-12: every standalone function becomes a class, one class per file). The
"`classify_article_size()` Placement" section above (appending a bare function to
`article_size.py`) is superseded: it is now `ArticleSizeClassifier.classify(char_count)` at
`src/domain/classification/article_size_classifier.py`, test at
`src/domain/tests/classification/test_article_size_classifier.py`. `article_size.py` reverts to
containing only the `ArticleSize` enum. The `ArticleClassifier` Full Design snippet's call
`classify_article_size(document_content.char_count)` must become
`self._article_size_classifier.classify(document_content.char_count)` with
`ArticleSizeClassifier` injected as an 8th constructor param when Phase 8 is implemented.

**2. `LlmGeneratorPort` is `ABC`, not `Protocol`** (project-wide port convention, stated
explicitly during this review — overrides ADR-4's `Protocol` shape). `generate()` is now an
`@abstractmethod`. `OllamaGeneratorAdapter` already declared explicit inheritance, so this was a
2-line change with zero call-site impact (confirmed: full `src` suite, 304 tests, green after the
change).

**3. Module-level "single-class-usage" constants move inside their class** (user preference:
constants belong to the class that uses them; a constant shared by 2+ classes signals a design
problem, not a reason to keep it at module scope). This **overrides ADR-2's stated rationale**
for `ImrydSignalDetector` ("keyword lists are the algorithm itself, not an operational knob, so
they stay module-level like `quality_text_sampler.py`'s `_CONCLUSION_HEADER_PATTERN`") — that
rationale is no longer the project's position. Applied in this slice:
- `imryd_signal_detector.py`: `_IMRYD_KEYWORDS`, `_HEADER_CANDIDATE_MAX_WORD_COUNT` → class attrs.
- `article_classification_text_sampler.py`: all 5 constants → class attrs.
- `article_classification_response_parser.py`: `_S4_PATTERN`/`_S5_PATTERN`/`_S6_PATTERN` → class
  attrs, **and renamed** to `_RESEARCH_INTENT_PATTERN`/`_EVIDENCE_BASED_CONTRIBUTION_PATTERN`/
  `_THEORETICAL_JUSTIFICATION_PATTERN` (aligns with ADR-7's `_ClassificationSignals` field names —
  the regex *content* still matches literal `"S4:"`/`"S5:"`/`"S6:"` tokens, since that's the LLM's
  actual wire-format per the prompt template; only the Python identifier changed).

**Known tech debt — NOT touched in this slice** (Slice 5, already merged into
`refactor/hexagonal-migration`, same violations, deferred to a dedicated refactor pass so it
doesn't bloat this PR):
- `src/domain/enums/quality_level.py`'s `get_quality_level_from_score()` → should become
  `QualityLevelResolver.resolve()` per the same full-OOP decision as point 1 above.
- `src/domain/quality/quality_text_sampler.py`'s module-level `_CONCLUSION_HEADER_PATTERN` →
  should move inside `QualityTextSampler` per point 3 above.
- `src/domain/tests/quality/test_quality_text_sampler.py`'s module-level `build_document_content()`
  helper → should become a private `TestCase` method (same fix applied in this slice to
  `test_article_classification_text_sampler.py` and `test_imryd_signal_detector.py`).
- `src/application/tests/fake_llm_generator_adapter.py`'s `FakeLlmGeneratorAdapterForTest.generate()`
  is missing the `options` parameter added to `LlmGeneratorPort` in this slice (ADR-4) — latent
  signature drift, not currently exercised by any test that passes `options=`, found incidentally
  while checking ABC-conversion impact, not part of this review's original scope.

---

## Amendment 2 (PR-2 code review, post-implementation)

PR-2's first draft named the 19 reasoning functions `_reasoning_case_2`...`_reasoning_case_19`
(legacy case numbers as identifiers). User review rejected this: numbers inside identifiers
aren't self-explanatory, and the project already has the right precedent —
`ClassificationConfidence`'s 4 members (`FULL_SIGNAL_MATCH`, `RECENT_BIBLIOGRAPHY_SUPPORT`,
`COMPLETE_BIBLIOGRAPHY_SUPPORT`, `SUFFICIENT_REFERENCE_COUNT`) describe the *condition*, not the
legacy case number, with the case-number cross-reference kept only in this doc's ADR-3 table.

**Renamed, all 19 now describe the signal combination that triggers them** (same vocabulary as
`_ClassificationSignals`'s field names):

| Old (rejected) | New | Legacy case |
|---|---|---|
| `_reasoning_case_2` | `_reasoning_full_signal_match` | 2 |
| `_reasoning_case_3` | `_reasoning_recent_bibliography_support` | 3 |
| `_reasoning_case_4` | `_reasoning_complete_bibliography_support` | 4 |
| `_reasoning_case_5` | `_reasoning_sufficient_reference_count` | 5 |
| `_reasoning_case_6` | `_reasoning_near_miss_theoretical_justification_only` | 6 |
| `_reasoning_case_7` | `_reasoning_near_miss_recent_bibliography_only` | 7 |
| `_reasoning_case_8` | `_reasoning_near_miss_sufficient_references_only` | 8 |
| `_reasoning_case_9` | `_reasoning_near_miss_no_bibliographic_support` | 9 |
| `_reasoning_case_10` | `_reasoning_vocabulary_and_research_intent` | 10 |
| `_reasoning_case_11` | `_reasoning_vocabulary_and_evidence_based_contribution` | 11 |
| `_reasoning_case_12` | `_reasoning_vocabulary_and_complete_bibliography` | 12 |
| `_reasoning_case_13` | `_reasoning_vocabulary_and_sufficient_references` | 13 |
| `_reasoning_case_14` | `_reasoning_vocabulary_and_recent_bibliography` | 14 |
| `_reasoning_case_15` | `_reasoning_vocabulary_only` | 15 |
| `_reasoning_case_16` | `_reasoning_research_intent_and_evidence_based_contribution` | 16 |
| `_reasoning_case_17` | `_reasoning_research_intent_only` | 17 |
| `_reasoning_case_18` | `_reasoning_evidence_based_contribution_only` | 18 |
| `_reasoning_case_19` | `_reasoning_no_signals_detected` | 19 (OPINION fallback) |

With no numbers in any name, alphabetical order (the project's method-ordering convention) needs
no special-case rule here — the 18 `@staticmethod` functions sort cleanly by string comparison.
Re-verified after the rename: legacy vs. new `_apply_rule()` produce byte-identical
`article_type`/`confidence`/`reasoning` for all 19 cases (programmatic diff, not just visual
inspection — see PR-2's review notes).

**Separate doc-only correction**: ADR-6's text above says the rule table has "18" rows; the
table itself (and the legacy source) has 17 rows (cases 2-18) plus the case-19 fallback outside
the table. `len(_RULE_TABLE) == 17` is correct; "18" in the prose was a miscount, not a behavior
change.

**New general rule, for any future case where a number must stay inside a method name** (recorded
in `docs/plan-migracion-hexagonal.md` §9): numeric order takes precedence over lexicographic
string order — `_2`/`_02` sorts before `_19`/`_019`, not after, the way plain string comparison
would put it. Not exercised in this slice since the rename above eliminated the only case that
would have needed it, but the rule stands for future code.
