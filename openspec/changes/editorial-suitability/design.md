# Design: Editorial Suitability

## Technical Approach
This change introduces qualitative editorial suitability analysis (value contribution and research lines alignment) for Visión Conjunta.

The pipeline is extended by creating a stateless domain service `EditorialSuitabilityAnalyzer` and a parser `EditorialSuitabilityParser`. The suitability result is embedded in `QualityResultDTO` and rendered in the Word report and Gradio UI.

## Architecture Decisions

| Decision | Options | Tradeoffs | Selected |
|---|---|---|---|
| **DTO Placement** | A) Return separate DTO from use case.<br>B) Nest inside `QualityResultDTO`. | A requires major use case and controller signature changes.<br>B keeps interfaces clean with backward compatibility. | **B**: Nest in `QualityResultDTO` |
| **Resolver Removal** | A) Retain `QualityLevelResolver` dependency.<br>B) Use direct enum function call. | A adds unnecessary collaborator.<br>B simplifies `QualityAnalyzer` constructor to exactly 5 collaborators. | **B**: Call `get_quality_level_from_score` directly |
| **Field Truncation** | A) Truncate raw string at fixed limit.<br>B) Extract first sentence, then truncate at word boundary. | A may cut mid-word/sentence.<br>B provides polished summaries matching EUMIC style. | **B**: Sentence extraction & word boundary truncation |
| **Research Lines Source** | A) Hardcoded Python module constant.<br>B) External text file read via `FileGatewayPort` at wiring time. | A requires a code change/redeploy whenever the Facultad updates its research lines.<br>B keeps `EditorialSuitabilityAnalyzer` free of I/O (receives `research_lines: str`, same pattern as prompt templates) and lets the file be edited without touching code. `FileGatewayPort`/`FileGatewayAdapter` already existed in the codebase, unused; `FileGatewayAdapter` had a pre-existing encoding bug (missing `encoding="utf-8"`) that must be fixed to safely read/write Spanish text on Windows. | **B**: Read `research_lines.txt` via `FileGatewayPort` in the wiring layer |

## Data Flow
```
DocumentContent ──► QualityTextSampler ──► text_sample
                                                 │
  ┌──────────────────────────────────────────────┴───────────────┐
  ▼                                                              ▼
QualityAnalyzer                                     EditorialSuitabilityAnalyzer
  ├─► LLM (Quality x2)                                ├─► LLM (Contribution)
  └─► QualityResponseParser                           ├─► LLM (Alignment)
                                                      └─► EditorialSuitabilityParser
                                                                 │
                                                                 ▼
                                                    EditorialSuitabilityDTO
                                                                 │
                                                                 ▼
QualityResultDTO (nested DTO) ◄──────────────────────────────────┘
```

## File Changes

| File | Action | Description |
|---|---|---|
| `src/domain/dtos/editorial_suitability_dto.py` | Create | `EditorialSuitabilityDTO` frozen dataclass |
| `src/domain/quality/editorial_suitability_parser.py` | Create | stateless regex parser and sentence truncator |
| `src/domain/quality/editorial_suitability_analyzer.py` | Create | stateless coordinator executing LLM calls |
| `src/domain/enums/quality_level.py` | Modify | Add `get_quality_level_from_score()` and constants |
| `src/domain/dtos/quality_result_dto.py` | Modify | Add optional `editorial_suitability` field |
| `src/domain/quality/quality_analyzer.py` | Modify | Update constructor to 5 collaborators; call resolver directly; delegate suitability analysis |
| `src/infrastructure/resources/prompts/quality/contribution_prompt.txt` | Create | Prompt template for contribution evaluation |
| `src/infrastructure/resources/prompts/quality/alignment_prompt.txt` | Create | Prompt template for research lines alignment |
| `src/infrastructure/resources/prompts/quality/research_lines.txt` | Create | External text file with the 7 FMC research lines (content pending editorial validation) |
| `src/infrastructure/adapters/gateway/file_gateway_adapter.py` | Modify | Fix missing `encoding="utf-8"` in `read()`/`write()` |
| `src/domain/quality/editorial_suitability_analyzer.py` | Modify | Replace `_RESEARCH_LINES` module constant with injected `research_lines: str` constructor param |
| `src/infrastructure/adapters/report/docx_report_adapter.py` | Modify | Add `_add_editorial_suitability` and call it in export pipeline |
| `src/infrastructure/wirings/analyze_document_use_case_wiring.py` | Modify | Wire `EditorialSuitabilityAnalyzer` and update `QualityAnalyzer` construction |
| `gradio_app.py` | Modify | Render editorial suitability section in results HTML |

## Interfaces / Contracts

```python
# src/domain/dtos/editorial_suitability_dto.py
from dataclasses import dataclass
from src.domain.dtos.base_dto import BaseDTO

@dataclass(frozen=True)
class EditorialSuitabilityDTO(BaseDTO):
    contribution_verdict: str  # SUSTENTADA, PARCIAL, NO SUSTENTADA
    contribution_phrase: str
    contribution_observation: str
    alignment_verdict: str  # ALINEADO, PARCIALMENTE ALINEADO, NO ALINEADO
    alignment_lines: str
    alignment_justification: str
```

## Testing Strategy

| Layer | What to Test | Approach |
|---|---|---|
| Unit | `EditorialSuitabilityParser` regex parsing and truncation | Parametrized tests covering case insensitivity, word boundaries, and fallback sentences |
| Unit | `EditorialSuitabilityAnalyzer` orchestration | Mock `LlmGeneratorPort` verifying exactly 2 calls with temperature 0.1, num_predict 300 |
| Unit | `QualityAnalyzer` integration | Verify nested suitability DTO in `QualityResultDTO` and collaborator limits |
| Integration | `DocxReportAdapter` rendering | Verify Word report generation formats suitability correctly |

## Threat Matrix
`N/A — no routing, shell, subprocess, VCS/PR automation, executable-file classification, or process-integration boundary.`

## Migration / Rollout
No migration required. No changes to database schemas or external systems.
