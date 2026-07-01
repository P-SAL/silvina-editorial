# Delta Spec: Report Export (Slice 13 + Refactor Delta)

> Parent spec: `openspec/specs/export-report/spec.md`
> Parent proposal: `openspec/changes/analyze-document-orchestrator/proposal.md`
> Normative guide: `.agent/skills/clean-architecture/SKILL.md`
> **Refactor note**: Slice 14 integration observations updated the verdict rendering model. This document reflects the final implemented state.

---

## Purpose

This delta specification defines the modifications to the export-report capabilities. It updates `ReportInputDTO` and `DocxReportAdapter` to incorporate the strongly typed `RecommendationDTO` list, new `PublicationVerdictDTO`, new `EumicViolationDTO` list, and attribute-based access patterns.

---

## Requirements

### Requirement: ReportInputDTO Fields Update

`ReportInputDTO` (in `src/domain/dtos/report_input_dto.py`) MUST include the following fields as part of Slice 13:

- `recommendations: list[RecommendationDTO]` — specific editorial recommendations (HIGH/MEDIUM/LOW priority)
- `verdict: PublicationVerdictDTO` — the final publication verdict (CRITICAL/WARNING/APPROVED)
- `eumic_violations: list[EumicViolationDTO]` — EUMIC formatting compliance violations

> **Rationale**: `verdict` was split from `recommendations` because a publication verdict is semantically distinct from a specific editorial recommendation. Mixing CRITICAL/WARNING/APPROVED into `RecommendationPriority` conflated two different concepts.

#### Scenario: ReportInputDTO holds typed verdict and recommendation lists

- GIVEN a valid instance of `ReportInputDTO`
- WHEN `recommendations`, `verdict`, and `eumic_violations` are read
- THEN `recommendations` is `list[RecommendationDTO]`, `verdict` is `PublicationVerdictDTO`, and `eumic_violations` is `list[EumicViolationDTO]`

---

### Requirement: DocxReportAdapter Verdict-Based Rendering

`DocxReportAdapter` (in `src/infrastructure/adapters/report/docx_report_adapter.py`) MUST render the publication verdict using `report_input.verdict` (a `PublicationVerdictDTO`) rather than searching for a final recommendation inside `report_input.recommendations`.

The `_add_recommendations` method MUST:
1. Always render the verdict paragraph from `report_input.verdict.message`, styled bold with color from `verdict_colors[verdict.verdict]` where:
   - `PublicationVerdict.CRITICAL` → `settings.reject_color_rgb`
   - `PublicationVerdict.WARNING` → `settings.reject_color_rgb`
   - `PublicationVerdict.APPROVED` → `settings.publishable_color_rgb`
2. Then, if `report_input.recommendations` is non-empty, render the specific recommendations list using `RecommendationPriority` icons:
   - `HIGH` → `"🔴"`
   - `MEDIUM` → `"🟡"`
   - `LOW` → `"🟢"`
3. Use attribute access (`rec.priority`, `rec.message`) — NOT dictionary access.

> **Rationale**: The verdict is always present (even when there are no specific recommendations). Rendering it independently of the recommendations list makes the report structure predictable and eliminates the need to search recommendations for a special "final" entry.

#### Scenario: Adapter renders verdict from PublicationVerdictDTO

- GIVEN a `ReportInputDTO` with `verdict.verdict = PublicationVerdict.APPROVED`
- WHEN `DocxReportAdapter._add_recommendations` is called
- THEN the verdict message is rendered bold with `publishable_color_rgb`

#### Scenario: Adapter renders specific recommendations after verdict

- GIVEN a `ReportInputDTO` with a non-empty `recommendations` list
- WHEN `_add_recommendations` is called
- THEN each recommendation is rendered as a bullet with the corresponding priority icon

#### Scenario: Adapter renders verdict even when recommendations list is empty

- GIVEN a `ReportInputDTO` with an empty `recommendations` list
- WHEN `_add_recommendations` is called
- THEN the verdict paragraph is still rendered (no KeyError or AttributeError)
