# Editorial Suitability Specification

## Purpose
Qualitative evaluation of an academic article's suitability for a military research journal across two dimensions: value contribution and alignment with the military research lines.

## Requirements

### Requirement: EditorialSuitabilityDTO

`EditorialSuitabilityDTO` (`src/domain/dtos/editorial_suitability_dto.py`) is a frozen `BaseDTO` containing:
- `contribution_verdict: str` (SUSTENTADA, PARCIAL, or NO SUSTENTADA)
- `contribution_phrase: str` (Extracted contribution summary)
- `contribution_observation: str` (Verdict-consistent justification)
- `alignment_verdict: str` (ALINEADO, PARCIALMENTE ALINEADO, or NO ALINEADO)
- `alignment_lines: str` (Matched research line numbers and themes)
- `alignment_justification: str` (Relationship justification)

#### Scenario: DTO structure is valid

- GIVEN `EditorialSuitabilityDTO` instantiation
- WHEN fields are inspected
- THEN all six string fields are present and immutable

### Requirement: EditorialSuitabilityParser

`EditorialSuitabilityParser` (`src/domain/quality/editorial_suitability_parser.py`) is a stateless parser service.
- It parses raw text outputs into verdicts and justifications using regex case-insensitively:
  - `VEREDICTO: <verdict_value>`
  - `CONTRIBUCION: <contribution_value>`
  - `OBSERVACION: <observation_value>`
  - `LINEAS: <lines_value>`
  - `JUSTIFICACION: <justification_value>`
- It truncates extracted fields at the first sentence's word boundary:
  - `contribution_phrase`, `contribution_observation`, `alignment_justification` to max 120 characters.
  - `alignment_lines` to max 80 characters.
  - If a sentence is truncated, it appends a single trailing `…`.
- Enforces consistency between the contribution verdict and observation:
  - `NO SUSTENTADA` -> observation is `"Sin contribución observada o declarada."`
  - `PARCIAL` -> observation is `"Contribución declarada pero no suficientemente sustentada."`
  - `SUSTENTADA` -> observation is `"Contribución sustentada — {contribution_phrase}"` (or `"Contribución sustentada."` if no phrase was successfully parsed).

#### Scenario: Parse contribution response with NO SUSTENTADA verdict

- GIVEN a raw text containing "VEREDICTO: NO SUSTENTADA"
- WHEN `parse_contribution` is called
- THEN verdict is "NO SUSTENTADA" and observation is "Sin contribución observada o declarada."

#### Scenario: Truncate long justification at word boundary

- GIVEN a justification longer than 120 characters
- WHEN `parse_alignment` is called
- THEN the justification is truncated to under 120 characters at a word boundary and ends with "…"

### Requirement: EditorialSuitabilityAnalyzer

`EditorialSuitabilityAnalyzer` (`src/domain/quality/editorial_suitability_analyzer.py`) is a stateless domain service that coordinates the calls.
- Injected with `LlmGeneratorPort`, `EditorialSuitabilityParser`, contribution prompt, alignment prompt, and `research_lines: str`.
- Calls `LlmGeneratorPort.generate` exactly twice: once for contribution and once for alignment (using temperature 0.1, num_predict 300).
- `research_lines` is not hardcoded in the domain service: the wiring layer reads it from `src/infrastructure/resources/prompts/quality/research_lines.txt` via `FileGatewayPort`/`FileGatewayAdapter` and injects it as a constructor argument, so the Facultad's research lines can be updated without a code change.

#### Scenario: LLM port called exactly twice

- GIVEN a mock `LlmGeneratorPort`
- WHEN `analyze` is called with a document sample
- THEN the mock port's `generate` is called exactly 2 times
