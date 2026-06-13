# Delta Spec: Domain Foundations (Slice 0)

> Normative guide: `.agent/skills/clean-architecture/SKILL.md`
> Parent proposal: `openspec/changes/domain-foundations/proposal.md`
> Migration plan reference: `docs/plan-migracion-hexagonal.md` §4.1, §8

---

## Purpose

This spec defines WHAT must be true after the domain-foundations migration is
applied. It does not describe HOW to implement it. Every requirement is
expressed as a testable assertion; each scenario translates directly to a
failing-first `unittest.TestCase`.

---

## Entity vs DTO Classification (resolved)

Resolved case-by-case using the SKILL criteria:
- **Entity** (`BaseEntity` + `@dataclass` mutable): has behavior methods or
  mutating `__post_init__` logic (enforces invariants, recalculates fields,
  validates on construction).
- **DTO** (`BaseDTO`, `@dataclass(frozen=True)`): immutable record that crosses
  a boundary (enters/exits a use case or adapter); carries only data.

| Legacy type | Classification | Reason |
|---|---|---|
| `Citation` | **Entity** | Has `__str__` (behavior); mutable; consumed inside domain logic, not a boundary record |
| `Reference` | **Entity** | Has `__str__` (behavior); mutable; consumed inside domain logic |
| `Section` | **Entity** | Has behavior: `get_word_count()`, `is_empty()`; `__post_init__` validates title and auto-detects section type — **cannot be frozen** |
| `DocumentContent` | **Entity** | `__post_init__` mutates `word_count` from paragraphs — **cannot be frozen**; aggregates mutable entities |
| `ClassificationResult` | **DTO** | Immutable output record; crosses use-case boundary; no mutation after creation |
| `QualityResult` | **DTO** | Immutable output record; `QualityAnalysisResult` duplicate eliminated here |
| `StructureValidationResult` | **DTO** | Immutable output record; boundary-crossing output of validation use case |
| `CitationAnalysisResult` | **DTO** | Immutable output record; boundary-crossing output of citation use case |
| `AnalysisResult` | **DTO** | Top-level immutable aggregate; crosses the boundary to presenters/formatters |

**Critical constraint**: `BaseDTO` is `@dataclass(frozen=True)`. Any type
with mutating `__post_init__` or behavior methods MUST be an Entity. Placing
`Section` or `DocumentContent` as a DTO would cause a `FrozenInstanceError`
at runtime and violate the SKILL invariant.

---

## Requirement: Enums

### REQ-ENUM-1 — Each enum is its own module

After migration, every enum from `domain/enums.py` SHALL live in its own
dedicated file under `src/domain/enums/`, one enum class per file, following
the SKILL one-class-per-file rule.

**Files required:**

| Class | File path |
|---|---|
| `ArticleType` | `src/domain/enums/article_type.py` |
| `ArticleSize` | `src/domain/enums/article_size.py` |
| `CitationType` | `src/domain/enums/citation_type.py` |
| `ClassificationCategory` | `src/domain/enums/classification_category.py` |
| `QualityLevel` | `src/domain/enums/quality_level.py` |
| `SectionType` | `src/domain/enums/section_type.py` |
| `AnalysisDimension` | `src/domain/enums/analysis_dimension.py` |
| `ValidationStatus` | `src/domain/enums/validation_status.py` |
| `RecommendationPriority` | `src/domain/enums/recommendation_priority.py` |
| `SeverityLevel` | `src/domain/enums/severity_level.py` |

**Bug #1 fix baked in**: `SeverityLevel` was listed in `__all__` before its
definition (line 277 before line 284). By placing each enum in its own module,
no shared `__all__` exists and the ordering bug disappears by construction.

#### Scenario: ArticleType enum is importable from src

- GIVEN the `src/domain/enums/article_type.py` module exists
- WHEN `from src.domain.enums.article_type import ArticleType` is executed
- THEN the import succeeds and `ArticleType.CIENTIFICO.value == "científico"`

#### Scenario: SeverityLevel is importable before any other enum

- GIVEN the `src/domain/enums/severity_level.py` module exists
- WHEN `from src.domain.enums.severity_level import SeverityLevel` is imported
  independently without importing any other enum
- THEN `SeverityLevel.CRITICAL.value == "critical"` is accessible (bug #1 fixed)

#### Scenario: All ten enums are importable independently

- GIVEN all ten enum modules exist under `src/domain/enums/`
- WHEN each module is imported individually in isolation
- THEN each import succeeds and exposes the correct members with expected values

### REQ-ENUM-2 — Enum member values are preserved

Each migrated enum MUST expose the same member names and values as the legacy
original.

#### Scenario: ArticleType preserves all values

- GIVEN `ArticleType` migrated to `src/domain/enums/article_type.py`
- WHEN the enum members are inspected
- THEN members are `CIENTIFICO`, `DIVULGACION`, `OPINION`, `UNKNOWN` with values
  `"científico"`, `"divulgación"`, `"opinión"`, `"unknown"` respectively

#### Scenario: SectionType preserves all 22 members

- GIVEN `SectionType` migrated to `src/domain/enums/section_type.py`
- WHEN the enum members are counted
- THEN `len(SectionType)` equals 22 and all bilingual section names are present

### REQ-ENUM-3 — Enum imports use only stdlib

Each enum module MUST import only from `enum` (stdlib). No cross-domain or
external imports inside an enum file.

#### Scenario: Enum module has no domain imports

- GIVEN any single enum file under `src/domain/enums/`
- WHEN the file's module-level imports are inspected
- THEN the only import is `from enum import Enum` (and `IntEnum` / `Flag` if
  applicable); no `src.domain` imports appear

---

## Requirement: Citation Entity

### REQ-ENTITY-CITATION-1 — Citation is a BaseEntity subclass

`Citation` SHALL be migrated to `src/domain/citation/citation.py` as a
`@dataclass` subclass of `BaseEntity`.

#### Scenario: Citation is a BaseEntity

- GIVEN `from src.domain.citation.citation import Citation`
- WHEN `issubclass(Citation, BaseEntity)` is checked
- THEN the result is `True`

### REQ-ENTITY-CITATION-2 — Citation fields and types

`Citation` MUST expose fields: `text: str`, `citation_type: CitationType`,
`location: int`, `author: str | None = None`, `year: str | None = None`.

No `Optional` type hint; Python 3.10+ union syntax only.

#### Scenario: Citation instantiation with required fields only

- GIVEN `Citation(text="(Smith, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=5)`
- WHEN the instance is created
- THEN `citation.text == "(Smith, 2020)"`, `citation.location == 5`,
  `citation.author is None`, `citation.year is None`

#### Scenario: Citation as_dict returns expected keys

- GIVEN a fully-populated `Citation` instance
- WHEN `citation.as_dict()` is called
- THEN the returned dict contains keys `text`, `citation_type`, `location`,
  `author`, `year`

### REQ-ENTITY-CITATION-3 — Citation has __str__ behavior

`Citation` MUST retain `__str__` returning a preview of the text.

#### Scenario: Citation __str__ truncates at 50 chars

- GIVEN a `Citation` with `text` longer than 50 characters
- WHEN `str(citation)` is called
- THEN the result starts with `"Citation("` and contains `"..."`

---

## Requirement: Reference Entity

### REQ-ENTITY-REFERENCE-1 — Reference is a BaseEntity subclass

`Reference` SHALL be migrated to `src/domain/reference/reference.py` as a
`@dataclass` subclass of `BaseEntity`.

#### Scenario: Reference is a BaseEntity

- GIVEN `from src.domain.reference.reference import Reference`
- WHEN `issubclass(Reference, BaseEntity)` is checked
- THEN the result is `True`

### REQ-ENTITY-REFERENCE-2 — Reference fields

`Reference` MUST expose: `text: str`, `authors: str | None = None`,
`year: str | None = None`, `title: str | None = None`,
`source: str | None = None`. No `Optional`.

#### Scenario: Reference instantiation with only required field

- GIVEN `Reference(text="Smith, J. (2020). Title. Journal.")`
- WHEN the instance is created
- THEN `reference.authors is None` and `reference.year is None`

#### Scenario: Reference __str__ returns formatted string

- GIVEN `Reference(text="...", authors="Smith", year="2020")`
- WHEN `str(reference)` is called
- THEN the result is `"Reference(Smith, 2020)"`

---

## Requirement: Section Entity

### REQ-ENTITY-SECTION-1 — Section is a BaseEntity subclass

`Section` SHALL be migrated to `src/domain/section/section.py` as a
`@dataclass` subclass of `BaseEntity`. It MUST NOT be a DTO: `get_word_count()`
and `is_empty()` are behavior methods incompatible with `frozen=True`.

#### Scenario: Section is a BaseEntity

- GIVEN `from src.domain.section.section import Section`
- WHEN `issubclass(Section, BaseEntity)` is checked
- THEN the result is `True`

### REQ-ENTITY-SECTION-2 — Section fields

`Section` MUST expose: `title: str`, `content: str`,
`section_type: SectionType | None = None`, `start_position: int = 0`,
`end_position: int = 0`, `level: int = 1`. No `Optional`.

#### Scenario: Section with missing title raises ValueError

- GIVEN an attempt to instantiate `Section(title="", content="Some text")`
- WHEN the dataclass `__post_init__` runs
- THEN a `ValueError` is raised with a message about title not being empty

### REQ-ENTITY-SECTION-3 — Section __post_init__ does NOT call domain helpers

The legacy `Section.__post_init__` used a local import of
`classify_section_by_name` from `domain.enums` — a local (in-function) import
violating SKILL §3 and a cross-module coupling to a helper out of this slice's
scope.

In the migrated `Section`, auto-detection via `classify_section_by_name` MUST
be removed. `section_type` defaults to `None` when not provided; callers supply
it explicitly. No local imports; no cross-module function calls in `__post_init__`.

#### Scenario: Section without section_type has section_type None

- GIVEN `Section(title="Introduction", content="...")`  (no `section_type` arg)
- WHEN the instance is created
- THEN `section.section_type is None` (no auto-detection; bug #7 local-import
  removed)

#### Scenario: Section with explicit section_type preserves it

- GIVEN `Section(title="Introduction", content="...", section_type=SectionType.INTRODUCTION)`
- WHEN the instance is created
- THEN `section.section_type == SectionType.INTRODUCTION`

### REQ-ENTITY-SECTION-4 — Section behavior methods

`Section` MUST expose `get_word_count() -> int` and `is_empty() -> bool`.

#### Scenario: get_word_count returns word count of content

- GIVEN `Section(title="T", content="Hello world foo")`
- WHEN `section.get_word_count()` is called
- THEN the result is `3`

#### Scenario: is_empty returns True for blank content

- GIVEN `Section(title="T", content="   ")`
- WHEN `section.is_empty()` is called
- THEN the result is `True`

#### Scenario: is_empty returns False for non-blank content

- GIVEN `Section(title="T", content="Some text")`
- WHEN `section.is_empty()` is called
- THEN the result is `False`

---

## Requirement: DocumentContent Entity

### REQ-ENTITY-DOCCONTENT-1 — DocumentContent is a BaseEntity subclass

`DocumentContent` SHALL be migrated to
`src/domain/document/document_content.py` as a `@dataclass` subclass of
`BaseEntity`. It MUST NOT be a DTO: its `__post_init__` mutates `word_count`
when paragraphs are present — frozen instances cannot be mutated after `__init__`.

#### Scenario: DocumentContent is a BaseEntity

- GIVEN `from src.domain.document.document_content import DocumentContent`
- WHEN `issubclass(DocumentContent, BaseEntity)` is checked
- THEN the result is `True`

### REQ-ENTITY-DOCCONTENT-2 — DocumentContent fields

`DocumentContent` MUST expose (with legacy field names preserved):
`word_count: int`, `char_count: int`, `paragraph_count: int = 0`,
`title: str | None = None`, `authors: str | None = None`,
`abstract: str | None = None`, `keywords: list[str]` (default empty),
`references: list[Reference]` (default empty), `paragraphs: list[str]`
(default empty), `sections: dict[str, str]` (default empty).

Types MUST use Python 3.10+ syntax: `list[T]`, `dict[K, V]`, `T | None`.
No `List`, `Dict`, `Optional`.

#### Scenario: DocumentContent field types use modern Python syntax

- GIVEN `DocumentContent` class definition in `src/domain/document/document_content.py`
- WHEN `get_type_hints(DocumentContent)` is inspected
- THEN no hint contains `typing.List`, `typing.Dict`, or `typing.Optional`

### REQ-ENTITY-DOCCONTENT-3 — __post_init__ auto-calculates word_count

When `word_count` is `0` and `paragraphs` is non-empty, `__post_init__` MUST
calculate `word_count` from the paragraphs.

#### Scenario: word_count is auto-calculated from paragraphs

- GIVEN `DocumentContent(word_count=0, char_count=100, paragraphs=["hello world", "foo"])`
- WHEN the instance is created
- THEN `document.word_count == 3` (sum of words across paragraphs)

#### Scenario: explicit word_count is not overwritten

- GIVEN `DocumentContent(word_count=42, char_count=100, paragraphs=["hello"])`
- WHEN the instance is created
- THEN `document.word_count == 42` (non-zero value preserved)

---

## Requirement: ClassificationResult DTO

### REQ-DTO-CLASSIFICATION-1 — ClassificationResult is a BaseDTO subclass

`ClassificationResult` SHALL be migrated to
`src/domain/dtos/classification_result_dto.py` as a `@dataclass(frozen=True)`
subclass of `BaseDTO`.

#### Scenario: ClassificationResult is a BaseDTO

- GIVEN `from src.domain.dtos.classification_result_dto import ClassificationResult`
- WHEN `issubclass(ClassificationResult, BaseDTO)` is checked
- THEN the result is `True`

### REQ-DTO-CLASSIFICATION-2 — ClassificationResult fields

`ClassificationResult` MUST expose:
`article_type: ArticleType`, `article_size: ArticleSize`,
`confidence: float | None`, `reasoning: str`,
`timestamp: datetime` (default `datetime.now()`).

#### Scenario: ClassificationResult instantiation with correct fields

- GIVEN `ClassificationResult(article_type=ArticleType.CIENTIFICO, article_size=ArticleSize.LARGO, confidence=0.9, reasoning="...")`
- WHEN the instance is created
- THEN `result.article_type == ArticleType.CIENTIFICO` and `result.confidence == 0.9`

#### Scenario: ClassificationResult is immutable

- GIVEN a valid `ClassificationResult` instance
- WHEN an attempt is made to assign `result.article_type = ArticleType.OPINION`
- THEN a `FrozenInstanceError` (or `dataclasses.FrozenInstanceError`) is raised

### REQ-DTO-CLASSIFICATION-3 — ClassificationResult.create() factory method

The legacy `create_classification_result()` module-level helper is broken: it
passes `category=` but the field name is `article_type`, and it omits
`article_size` (required field) — a `TypeError` at runtime.

The migrated DTO MUST provide a corrected `@classmethod` factory
`ClassificationResult.create(article_type, article_size, confidence, reasoning)`
that constructs the DTO with the correct field names. The broken legacy helper
is NOT reproduced in `src/`; it remains in legacy code untouched.

#### Scenario: create() factory builds a valid ClassificationResult

- GIVEN `ClassificationResult.create(article_type=ArticleType.DIVULGACION, article_size=ArticleSize.CORTO, confidence=0.75, reasoning="Evidence-based classification")`
- WHEN the factory is called
- THEN the returned instance has `article_type == ArticleType.DIVULGACION`,
  `article_size == ArticleSize.CORTO`, `confidence == 0.75`, and a `timestamp`
  is set automatically

#### Scenario: create() without confidence produces None confidence

- GIVEN `ClassificationResult.create(article_type=ArticleType.UNKNOWN, article_size=ArticleSize.FUERA_RANGO, confidence=None, reasoning="Could not determine")`
- WHEN the factory is called
- THEN `result.confidence is None`

#### Scenario: create() returns a frozen instance

- GIVEN a `ClassificationResult` built via `create()`
- WHEN an attempt is made to mutate any field
- THEN a `FrozenInstanceError` is raised

#### Scenario: ClassificationResult __str__ produces human-readable output

- GIVEN `ClassificationResult.create(article_type=ArticleType.CIENTIFICO, article_size=ArticleSize.LARGO, confidence=0.9, reasoning="...")`
- WHEN `str(result)` is called
- THEN the output contains `"científico"` and `"largo"` (enum values) and
  the confidence rendered as a percentage

---

## Requirement: QualityResult DTO (consolidation of QualityResult + QualityAnalysisResult)

### REQ-DTO-QUALITY-1 — QualityResult consolidates duplicate types

Legacy `models.py` defines both `QualityResult` and `QualityAnalysisResult`
(overlapping dataclasses for the same concept — bug #2). In `src/`, a **single**
`QualityResult` SHALL exist. `QualityAnalysisResult` is eliminated.

The consolidated type is `QualityResult` at
`src/domain/dtos/quality_result_dto.py`, subclass of `BaseDTO`.

Fields retained from `QualityResult` (the richer shape):
`overall_score: float`, `quality_level: QualityLevel`,
`dimension_scores: dict[str, dict[str, Any]]` (default empty dict),
`timestamp: datetime` (default `datetime.now()`).

`Any` MUST be imported from `typing`. `dict` uses lowercase generics. No
`Dict`, `Optional`, or use of builtin `any` as a type (bug #4 fixed).

#### Scenario: QualityResult is a BaseDTO

- GIVEN `from src.domain.dtos.quality_result_dto import QualityResult`
- WHEN `issubclass(QualityResult, BaseDTO)` is checked
- THEN the result is `True`

#### Scenario: Only one quality result type exists in src/

- GIVEN the `src/domain/dtos/` directory
- WHEN listing all DTO modules
- THEN there is no `quality_analysis_result_dto.py`; only `quality_result_dto.py`

#### Scenario: dimension_scores uses Any from typing not builtin any

- GIVEN the `quality_result_dto.py` source file
- WHEN the type annotation of `dimension_scores` is inspected
- THEN it is `dict[str, dict[str, Any]]` with `Any` from `typing` (bug #4 fixed)

#### Scenario: QualityResult __str__ returns score and level

- GIVEN `QualityResult(overall_score=8.5, quality_level=QualityLevel.GOOD)`
- WHEN `str(result)` is called
- THEN the output is `"Quality: 8.5/10 (Bueno)"`

### REQ-DTO-QUALITY-2 — QualityResult is immutable

#### Scenario: QualityResult mutation raises FrozenInstanceError

- GIVEN a valid `QualityResult` instance
- WHEN an attempt is made to assign `result.overall_score = 9.0`
- THEN a `FrozenInstanceError` is raised

---

## Requirement: StructureValidationResult DTO

### REQ-DTO-STRUCTURE-1 — StructureValidationResult is a BaseDTO subclass

`StructureValidationResult` SHALL be migrated to
`src/domain/dtos/structure_validation_result_dto.py` as a `@dataclass(frozen=True)`
subclass of `BaseDTO`.

Fields: `is_valid: bool`, `missing_sections: list[str]` (default empty),
`section_details: dict[str, dict[str, Any]]` (default empty; uses `Any` from
`typing`, not builtin `any` — bug #4 fixed), `timestamp: datetime` (default now).

#### Scenario: StructureValidationResult is a BaseDTO

- GIVEN `from src.domain.dtos.structure_validation_result_dto import StructureValidationResult`
- WHEN `issubclass(StructureValidationResult, BaseDTO)` is checked
- THEN the result is `True`

#### Scenario: StructureValidationResult __str__ for valid structure

- GIVEN `StructureValidationResult(is_valid=True, missing_sections=[])`
- WHEN `str(result)` is called
- THEN the output is `"Structure: Valid"`

#### Scenario: StructureValidationResult __str__ for invalid structure

- GIVEN `StructureValidationResult(is_valid=False, missing_sections=["abstract", "methodology"])`
- WHEN `str(result)` is called
- THEN the output is `"Structure: Invalid (2 missing)"`

#### Scenario: StructureValidationResult is immutable

- GIVEN a valid `StructureValidationResult` instance
- WHEN an attempt is made to assign `result.is_valid = False`
- THEN a `FrozenInstanceError` is raised

---

## Requirement: CitationAnalysisResult DTO

### REQ-DTO-CITATION-ANALYSIS-1 — CitationAnalysisResult is a BaseDTO subclass

`CitationAnalysisResult` SHALL be migrated to
`src/domain/dtos/citation_analysis_result_dto.py` as a `@dataclass(frozen=True)`
subclass of `BaseDTO`.

Fields: `total_citations: int`, `total_references: int`, `matched_count: int`,
`unmatched_count: int`, `citations_by_type: dict[str, int]` (default empty),
`unmatched_citations: list[str]` (default empty),
`timestamp: datetime` (default now).

#### Scenario: CitationAnalysisResult is a BaseDTO

- GIVEN `from src.domain.dtos.citation_analysis_result_dto import CitationAnalysisResult`
- WHEN `issubclass(CitationAnalysisResult, BaseDTO)` is checked
- THEN the result is `True`

#### Scenario: CitationAnalysisResult __str__ with citations

- GIVEN `CitationAnalysisResult(total_citations=10, total_references=8, matched_count=8, unmatched_count=2)`
- WHEN `str(result)` is called
- THEN the output is `"Citations: 10 (80.0% matched)"`

#### Scenario: CitationAnalysisResult __str__ with zero citations

- GIVEN `CitationAnalysisResult(total_citations=0, total_references=0, matched_count=0, unmatched_count=0)`
- WHEN `str(result)` is called
- THEN the output is `"Citations: 0 (0.0% matched)"` (no division by zero)

#### Scenario: CitationAnalysisResult is immutable

- GIVEN a valid `CitationAnalysisResult` instance
- WHEN an attempt is made to mutate any field
- THEN a `FrozenInstanceError` is raised

---

## Requirement: AnalysisResult DTO

### REQ-DTO-ANALYSIS-1 — AnalysisResult is a BaseDTO subclass

`AnalysisResult` SHALL be migrated to
`src/domain/dtos/analysis_result_dto.py` as a `@dataclass(frozen=True)`
subclass of `BaseDTO`. It aggregates the other result DTOs; crossing the
use-case boundary to presenters/formatters makes it the primary boundary-crossing
record.

Fields: `filename: str`,
`document_content: DocumentContent` (Entity, not frozen but contained by the frozen DTO),
`classification: ClassificationResult`, `quality: QualityResult`,
`structure: StructureValidationResult`, `citations: CitationAnalysisResult`,
`timestamp: datetime` (default now).

> Note: `document_content` is a `DocumentContent` Entity (mutable). Python's
> `frozen=True` freezes the DTO's own attribute references, not the objects
> they point to. `DocumentContent` instances are not mutated after they are
> set in `AnalysisResult`, so this composition is safe.

#### Scenario: AnalysisResult is a BaseDTO

- GIVEN `from src.domain.dtos.analysis_result_dto import AnalysisResult`
- WHEN `issubclass(AnalysisResult, BaseDTO)` is checked
- THEN the result is `True`

### REQ-DTO-ANALYSIS-2 — AnalysisResult.to_dict() serialization contract

`AnalysisResult` MUST expose a `to_dict() -> dict[str, Any]` method that
produces a plain dictionary suitable for serialization. The contract below is
normative; downstream formatters and future adapters depend on these keys.

Required top-level keys: `filename`, `timestamp`, `classification`, `quality`,
`structure`, `citations`.

Sub-dictionaries:

**`classification`**: `article_type` (enum `.value`), `article_size` (enum
`.value`), `confidence`, `reasoning`.
> Note: `article_type` is the corrected key (fixing the legacy bug #3 mapping
> where the field was named `article_type` but the broken helper passed `category=`).

**`quality`**: `overall_score`, `quality_level` (enum `.value`),
`dimension_scores`.

**`structure`**: `is_valid`, `missing_sections`, `section_details`.

**`citations`**: `total_citations`, `total_references`, `matched_count`,
`unmatched_count`, `citations_by_type`, `unmatched_citations`.

#### Scenario: to_dict returns all required top-level keys

- GIVEN a fully-populated `AnalysisResult` instance
- WHEN `result.to_dict()` is called
- THEN the returned dict contains exactly the keys: `filename`, `timestamp`,
  `classification`, `quality`, `structure`, `citations`

#### Scenario: to_dict classification sub-dict uses article_type key

- GIVEN a fully-populated `AnalysisResult` instance
- WHEN `result.to_dict()["classification"]` is inspected
- THEN it contains key `"article_type"` (not `"category"`) with the enum's
  string value (bug #3 corrected representation in the new type)

#### Scenario: to_dict timestamp is an ISO-8601 string

- GIVEN a fully-populated `AnalysisResult` instance
- WHEN `result.to_dict()["timestamp"]` is inspected
- THEN the value is a string matching ISO-8601 format (`.isoformat()` output)

#### Scenario: AnalysisResult is immutable

- GIVEN a valid `AnalysisResult` instance
- WHEN an attempt is made to assign `result.filename = "other.docx"`
- THEN a `FrozenInstanceError` is raised

---

## Requirement: Import Conventions

### REQ-IMPORTS-1 — No Optional, List, Dict type hints in migrated files

All migrated modules MUST use Python 3.10+ type syntax exclusively.

#### Scenario: No legacy typing generics in any migrated src file

- GIVEN all Python files under `src/domain/` created by this slice
- WHEN the files are statically scanned for the patterns `Optional[`, `List[`,
  `Dict[`
- THEN no such patterns appear in any migrated file

### REQ-IMPORTS-2 — No local (in-function) imports

SKILL §3 forbids imports inside functions or methods.

#### Scenario: No local imports in any migrated src file

- GIVEN all Python files under `src/domain/` created by this slice
- WHEN the files are statically scanned for import statements inside functions
  or methods (indented `import` / `from ... import`)
- THEN no such patterns appear (bug #7 fixed)

### REQ-IMPORTS-3 — No wildcard imports

#### Scenario: No wildcard imports in any migrated src file

- GIVEN all Python files under `src/domain/` created by this slice
- WHEN the files are statically scanned for `import *`
- THEN no such patterns appear

### REQ-IMPORTS-4 — Absolute imports from src.domain

All cross-domain imports within `src/` MUST use absolute paths starting with
`src.domain.`, never relative (`.enums`) or legacy (`domain.enums`).

#### Scenario: Section imports SectionType via absolute path

- GIVEN `src/domain/section/section.py`
- WHEN the import statements are inspected
- THEN `SectionType` is imported as
  `from src.domain.enums.section_type import SectionType`; no relative import
  or `from domain.` import is present (bug #7 fixed)

---

## Requirement: Test Coverage

### REQ-TEST-1 — Each migrated type has a unittest.TestCase

Every migrated enum, entity, and DTO MUST have a corresponding test file under
`src/domain/tests/`.

**Test file locations:**

| Migrated type | Test file |
|---|---|
| `ArticleType` | `src/domain/tests/enums/test_article_type.py` |
| `ArticleSize` | `src/domain/tests/enums/test_article_size.py` |
| `CitationType` | `src/domain/tests/enums/test_citation_type.py` |
| `ClassificationCategory` | `src/domain/tests/enums/test_classification_category.py` |
| `QualityLevel` | `src/domain/tests/enums/test_quality_level.py` |
| `SectionType` | `src/domain/tests/enums/test_section_type.py` |
| `AnalysisDimension` | `src/domain/tests/enums/test_analysis_dimension.py` |
| `ValidationStatus` | `src/domain/tests/enums/test_validation_status.py` |
| `RecommendationPriority` | `src/domain/tests/enums/test_recommendation_priority.py` |
| `SeverityLevel` | `src/domain/tests/enums/test_severity_level.py` |
| `Citation` | `src/domain/tests/entities/test_citation.py` |
| `Reference` | `src/domain/tests/entities/test_reference.py` |
| `Section` | `src/domain/tests/entities/test_section.py` |
| `DocumentContent` | `src/domain/tests/entities/test_document_content.py` |
| `ClassificationResult` | `src/domain/tests/dtos/test_classification_result.py` |
| `QualityResult` | `src/domain/tests/dtos/test_quality_result.py` |
| `StructureValidationResult` | `src/domain/tests/dtos/test_structure_validation_result.py` |
| `CitationAnalysisResult` | `src/domain/tests/dtos/test_citation_analysis_result.py` |
| `AnalysisResult` | `src/domain/tests/dtos/test_analysis_result.py` |

#### Scenario: Test suite passes after all types are migrated

- GIVEN all 19 migrated types and their test files exist
- WHEN `python -m pytest src/` is executed
- THEN all tests pass with zero failures or errors

### REQ-TEST-2 — Tests use unittest.TestCase exclusively

No pytest-only fixtures or marks. Tests MUST extend `unittest.TestCase` and use
`self.assert*` methods.

#### Scenario: Test files use TestCase

- GIVEN any test file under `src/domain/tests/` created by this slice
- WHEN the file is inspected for the base class
- THEN the test class extends `unittest.TestCase` imported as
  `from unittest import TestCase`

### REQ-TEST-3 — Tests validate bug fixes with explicit assertions

Each documented bug (#1, #2, #4, #7) MUST have at least one test assertion that
would fail if the bug were reintroduced.

#### Scenario: Bug #1 — SeverityLevel importable independently

See REQ-ENUM-1 scenarios above.

#### Scenario: Bug #2 — QualityAnalysisResult does not exist in src/

- GIVEN the `src/domain/dtos/` directory
- WHEN scanning for a `QualityAnalysisResult` class definition in `src/`
- THEN no such class exists; only `QualityResult` is present

#### Scenario: Bug #4 — dimension_scores annotation uses typing.Any

- GIVEN a `QualityResult` with `dimension_scores={"dim": {"score": 8.5}}`
- WHEN the instance is created and `result.dimension_scores` is accessed
- THEN no `NameError` is raised (no accidental use of builtin `any` as a type)

#### Scenario: Bug #7 — Section __post_init__ has no local import

- GIVEN `Section(title="Introduction", content="Some content")`
- WHEN instantiated
- THEN no `ImportError` or `ModuleNotFoundError` is raised and
  `section.section_type is None`

---

## Requirement: Coexistence

### REQ-COEXISTENCE-1 — Legacy code remains unmodified

No file under `domain/`, `business_logic/`, `data_access/`, `presentation/`,
`main.py`, `gradio_app.py`, or `tests/` is modified by this slice.

#### Scenario: Legacy domain still importable after migration

- GIVEN the migration is complete
- WHEN `from domain.models import Citation` is executed (legacy import)
- THEN the import succeeds and the legacy `Citation` dataclass is accessible

#### Scenario: Legacy pytest suite still passes

- GIVEN the migration is complete
- WHEN `python -m pytest tests/` is executed (legacy test runner)
- THEN all legacy tests pass (no regressions introduced)

---

## Requirement: File Structure

### REQ-STRUCTURE-1 — One class per file, PascalCase class = snake_case filename

SKILL §4 naming rule applied to all migrated types.

#### Scenario: Entity folder structure matches convention

- GIVEN all entity source files created by this slice
- WHEN the directory tree is inspected
- THEN each entity lives in its own folder: `src/domain/citation/citation.py`,
  `src/domain/reference/reference.py`, `src/domain/section/section.py`,
  `src/domain/document/document_content.py`

#### Scenario: DTO files follow naming convention

- GIVEN all DTO source files created by this slice
- WHEN the directory tree is inspected
- THEN DTOs are at `src/domain/dtos/<class_snake_case>_dto.py` with no extra suffix
  overlap (e.g., `classification_result_dto.py`, not `classification_result_dto_dto.py`)

### REQ-STRUCTURE-2 — Each file ends with exactly one blank line

#### Scenario: All migrated files end with a single newline

- GIVEN all `.py` files created by this slice
- WHEN the last two bytes of each file are inspected
- THEN each file ends with exactly one `\n` character

---

## Out of Scope (explicitly excluded from this spec)

The following are NOT requirements for this slice and MUST NOT be implemented:

- Port interfaces (`*Port`, `*Repository`) — Slice 1+
- Adapters — Slice 2+
- Use cases — Slice 3+
- Wirings — follows use cases
- Exception subclasses beyond `base_src_error.py` — Slice 1
- Helper functions as service classes (`classify_article_size`,
  `get_quality_level_from_score`, `classify_section_by_name`,
  `get_required_sections_for_category`, `get_citation_type_from_pattern`) — Slice 4+
- `create_empty_document`, `create_quality_result` legacy factory functions — out of scope
- Deleting or modifying `domain/enums.py` or `domain/models.py` — Slice 16
- `QualityAnalysisResult` in `src/` (eliminated; was a duplicate)

---

## Bug Fix Summary (normative)

| Bug | Legacy location | Status in src/ | Verified by |
|---|---|---|---|
| #1 `SeverityLevel` in `__all__` before definition | `enums.py:277` | Fixed by construction (no shared `__all__`, one file per enum) | REQ-ENUM-1 scenario |
| #2 `QualityAnalysisResult` duplicate | `models.py:79` | Eliminated; single `QualityResult` | REQ-DTO-QUALITY-1 scenario |
| #4 `any` (builtin) used as type annotation | `models.py:90,102,201` | Fixed: `dict[str, dict[str, Any]]` with `from typing import Any` | REQ-DTO-QUALITY-1, REQ-DTO-STRUCTURE-1 scenarios |
| #7 Local/mixed imports | `models.py:10-11, Section.__post_init__` | Fixed: absolute top-level imports only; local import in `__post_init__` removed | REQ-IMPORTS-2, REQ-ENTITY-SECTION-3 scenarios |

---

## ClassificationResult Factory Fix (normative — decision resolved)

The proposal flagged the spec executor's task: spec a corrected factory ON the
DTO. Resolution:

- Legacy `create_classification_result()` helper stays **untouched** in `domain/models.py`.
- The migrated `ClassificationResult` DTO at
  `src/domain/dtos/classification_result_dto.py` exposes:

  ```
  @classmethod
  def create(
      cls,
      article_type: ArticleType,
      article_size: ArticleSize,
      confidence: float | None,
      reasoning: str,
  ) -> "ClassificationResult":
  ```

  This factory is the ONLY recommended construction path for new code. It uses
  the correct field names and populates `timestamp` automatically.

- No `category=` parameter; no `article_type=` confusion; no missing required
  field. The legacy bug is documented (proposal bug table) but not carried
  forward.

---
