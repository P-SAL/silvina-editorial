# Design: validate-citations (Slice 4)

## Approach

Follow the Slice 2/3 pattern exactly: enum -> DTO -> stateless domain service -> thin use case
-> instance-based wiring. No ports/adapters — `CitationMatcher` is pure computation (regex and
dict lookups only), identical in nature to `ApaValidator` (Slice 3).

The key structural change versus the legacy class is **statelessness**: the legacy
`CitationMatcher.__init__` builds `self.citation_keys` / `self.reference_keys` lookup dicts from
constructor-injected `citations`/`references`. The migrated service takes `citations` and
`references` as method parameters on every call, with the lookup dicts built as **local
variables inside each method** (or via a shared private helper), not instance state. This
mirrors Slice 3 dropping `self.violations`.

## Module Layout

```
src/domain/citation/citation_matcher.py          # CitationMatcher domain service (new)
src/application/match_citations_use_case.py      # MatchCitationsUseCase (new)
src/infrastructure/wirings/match_citations_use_case_wiring.py   # MatchCitationsUseCaseWiring (new)
src/domain/tests/citation/test_citation_matcher.py              # domain tests (new)
src/infrastructure/tests/test_match_citations_use_case_wiring.py # wiring test (new)
```

`src/domain/citation/` and `src/domain/tests/citation/` already exist (Slice 3). No new
top-level folders needed.

## Domain Service: `CitationMatcher`

```python
# src/domain/citation/citation_matcher.py
import re

from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResultDTO
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.enums.citation_type import CitationType
from src.domain.enums.section_name import SectionName


class CitationMatcher:
    """Matches in-text citations with reference list entries via normalized author keys."""

    _NON_AUTHOR_PATTERNS = [
        r"^[A-Z]{2,}\s+\d",
        r"^arXiv:",
        r"^doi:",
        r"^repositorio",
        r"^no hay",
        r"^\w.*\d{4}.*\d{4}",
    ]

    def find_orphaned_citations(
        self, citations: list[CitationDTO], references: list[ReferenceDTO]
    ) -> list[CitationDTO]:
        """Return citations whose normalized author has no matching reference entry."""
        reference_keys = self._build_reference_keys(references)
        return [
            citation
            for citation in self._citable(citations)
            if self._is_orphaned_citation(citation, reference_keys)
        ]

    def find_orphaned_references(
        self, citations: list[CitationDTO], references: list[ReferenceDTO]
    ) -> list[ReferenceDTO]:
        """Return references never cited by any in-text citation."""
        citation_keys = self._build_citation_keys(citations)
        return [
            reference
            for reference in references
            if self._normalize_author(reference.text) not in citation_keys
        ]

    def match_citations_to_references(
        self,
        citations: list[CitationDTO],
        references: list[ReferenceDTO],
        section_type: SectionName = SectionName.REFERENCES,
    ) -> CitationAnalysisResultDTO:
        """Compute aggregate citation-reference match statistics for a section."""
        orphaned_citations = self.find_orphaned_citations(citations, references)
        valid_citations = self._citable(citations)
        return CitationAnalysisResultDTO(
            total_citations=len(valid_citations),
            total_references=len(references),
            matched_count=max(0, len(valid_citations) - len(orphaned_citations)),
            unmatched_count=len(orphaned_citations),
            citations_by_type={},
            unmatched_citations=[citation.text for citation in orphaned_citations],
        )

    def _citable(self, citations: list[CitationDTO]) -> list[CitationDTO]:
        return [
            citation
            for citation in citations
            if citation.citation_type != CitationType.FOOTNOTE and citation.author
        ]

    def _build_citation_keys(self, citations: list[CitationDTO]) -> dict[str, CitationDTO]:
        return {
            self._normalize_author(citation.author): citation
            for citation in self._citable(citations)
        }

    def _build_reference_keys(self, references: list[ReferenceDTO]) -> dict[str, ReferenceDTO]:
        return {self._normalize_author(reference.text): reference for reference in references}

    def _is_orphaned_citation(
        self, citation: CitationDTO, reference_keys: dict[str, ReferenceDTO]
    ) -> bool:
        key = self._normalize_author(citation.author)
        return key != "__non_author__" and key not in reference_keys

    def _normalize_author(self, text: str) -> str:
        """Extract and normalize the first author surname; non-author text yields a sentinel."""
        text_stripped = text.strip().lstrip("(").rstrip(")")
        is_non_author = any(
            re.search(pattern, text_stripped, re.IGNORECASE)
            for pattern in self._NON_AUTHOR_PATTERNS
        )
        if is_non_author:
            return "__non_author__"

        year_match = re.search(r"\((?:\d{1,2}\s+de\s+\w+\s+de\s+)?\d{4}[a-z]?\)", text)
        if year_match:
            text = text[: year_match.start()].strip()

        text = re.sub(r"\s+et\s+al\.?.*", "", text, flags=re.IGNORECASE).strip()
        text = re.sub(r"\s+y\s+.*", "", text, flags=re.IGNORECASE).strip()
        text = re.sub(r"\b[A-ZÁÉÍÓÚÑ]\.\s*", "", text).strip()
        text = re.sub(r"[,&.()\[\]]", "", text).strip()

        words = text.split()
        return words[0].lower() if words else ""
```

Notes:
- `_NON_AUTHOR_PATTERNS` regex list and the full body of `_normalize_author` are copied verbatim
  from `business_logic/citation_matcher.py` lines 61-97 — same patterns, same order, same
  transformations. No behavior changes.
- `_citable()` is a new private helper extracting the repeated `footnote or not author` filter
  that appeared three times in the legacy class (`__init__`, `find_orphaned_citations`,
  `match_citations_to_references`). This is a pure refactor for the stateless design, not a
  behavior change — same predicate, same result set.
- `CitationType.FOOTNOTE` comparison: legacy compares `cit.citation_type.value == 'footnote'`
  because legacy's enum is untyped Python (`domain.enums.SeverityLevel`-style). The migrated
  `CitationType` enum (`src/domain/enums/citation_type.py`) is a plain `Enum`, so the idiomatic
  comparison is `citation.citation_type != CitationType.FOOTNOTE` — equivalent result, no
  external behavior change since this is an internal implementation detail.

## Use Case: `MatchCitationsUseCase`

```python
# src/application/match_citations_use_case.py
from src.domain.citation.citation_matcher import CitationMatcher
from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResultDTO
from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.enums.section_name import SectionName


class MatchCitationsUseCase:
    def __init__(self, matcher: CitationMatcher) -> None:
        self._matcher = matcher

    def execute(
        self,
        citations: list[CitationDTO],
        references: list[ReferenceDTO],
        section_type: SectionName = SectionName.REFERENCES,
    ) -> CitationAnalysisResultDTO:
        return self._matcher.match_citations_to_references(
            citations=citations, references=references, section_type=section_type
        )
```

Thin pass-through, matching `ValidateApaUseCase.execute()`'s shape. No empty-list short-circuit
is needed here (unlike `ValidateApaUseCase`'s `if not citations` guard) because
`match_citations_to_references` already handles empty `citations`/`references` correctly —
`len([]) == 0`, `max(0, 0 - 0) == 0` — producing a valid zeroed `CitationAnalysisResultDTO`
without a special case.

## Wiring: `MatchCitationsUseCaseWiring`

Following the exact pattern in `src/infrastructure/wirings/validate_apa_wiring.py`
(`create_use_case()` + one private `_get_*`-style accessor per dependency):

```python
# src/infrastructure/wirings/match_citations_use_case_wiring.py
from src.application.match_citations_use_case import MatchCitationsUseCase
from src.domain.citation.citation_matcher import CitationMatcher


class MatchCitationsUseCaseWiring:
    """Factory for building a ready-to-use MatchCitationsUseCase."""

    def create_use_case(self) -> MatchCitationsUseCase:
        return MatchCitationsUseCase(matcher=self._get_citation_matcher())

    def _get_citation_matcher(self) -> CitationMatcher:
        return CitationMatcher()
```

This replicates `ValidateApaWiring` literally: public `create_use_case()` method, one private
`_get_<dependency>()` accessor returning the concrete domain service (no port/adapter split
needed since there is no I/O).

## Non-Goals (explicit)

- **`generate_report()`** — not migrated. Mixes Spanish presentation strings and emoji into
  business logic; same exclusion rationale as `ApaValidator.generate_report()` in Slice 3.
- **`extract_all_citations()`** — not migrated. Confirmed dead code (zero call sites, including
  internally within `citation_matcher.py`).
- **`ArticleAnalyzer` (`business_logic/article_analyzer.py`)** — entire module out of scope.
  Confirmed zero import/instantiation sites; was the only caller of `generate_report()`.
- **Orphaned-reference severity-by-section rule** (legacy `generate_report()` lines 145-153:
  `WARNING` if `section_type == "Referencias"` else `INFO`) is left in legacy code, untouched.
  Not extracted as a domain policy in this slice — no current consumer of the migrated service
  needs it, and `generate_report()` (its only context) has zero live callers. If a future slice
  needs this rule, it should become a small pure function under `src/domain/citation/` (for
  example `severity_for_orphaned_references(section: SectionName) -> SeverityLevel`), never a
  DTO, since DTOs in this codebase are plain dataclasses with no behavior.
- Legacy `business_logic/citation_matcher.py` stays unmodified; `main.py` keeps importing from
  `business_logic/`. Caller switchover is deferred to a future slice.

## Testing Plan

`src/domain/tests/citation/test_citation_matcher.py` (`unittest.TestCase`), mirroring the
`test_apa_validator_skip_patterns.py` / `test_apa_validator_*.py` split style used in Slice 3:

- One test per `_NON_AUTHOR_PATTERNS` entry (institutional acronym, arXiv, doi, repositorio, "no
  hay", date range) verifying `_normalize_author` returns `"__non_author__"`.
- Author normalization: et-al stripping, "y"-conjunction stripping, initials stripping,
  punctuation stripping, case-folding to first-author surname.
- `find_orphaned_citations`: citation with no matching reference is returned; footnote citations
  and citations without an author are excluded; non-author citations are excluded.
- `find_orphaned_references`: reference with no matching citation is returned.
- `match_citations_to_references`: returns a populated `CitationAnalysisResultDTO` with correct
  `total_citations`, `total_references`, `matched_count`, `unmatched_count`, matching the fields
  `main.py:241-244` currently consumes; verify the `SectionName` parameter accepts the enum, not
  a raw string.

`src/infrastructure/tests/test_match_citations_use_case_wiring.py`, mirroring
`test_validate_apa_wiring.py`: `create_use_case()` returns a `MatchCitationsUseCase`; calling
`execute()` with empty lists returns a `CitationAnalysisResultDTO`.

## Rejected Alternatives

- **Keep constructor-injected state (legacy shape)** — rejected per proposal; stateless-per-call
  matches the Slice 3 precedent (`ApaValidator` dropped `self.violations`) and avoids hidden
  mutable state that would force callers to re-instantiate the matcher for every
  citations/references pair instead of reusing one instance.
- **Extract the severity-by-section rule into a DTO method** — rejected; DTOs in this codebase
  are plain dataclasses with no behavior (`BaseDTO` convention). If ever needed, it becomes a
  free function in `src/domain/citation/`, not a DTO method.
- **Migrate `generate_report()` as a formatter use case now** — rejected; defers to a future
  formatter-adapter slice, same boundary Slice 3 drew for `ApaValidator.generate_report()`.
