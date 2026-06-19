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

    def analyze(self, document_content: DocumentContentDTO, article_type) -> QualityResultDTO:
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
        if not any(
            dimension in parsed_response.matched_dimensions for dimension in relevant_dimensions
        ):
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
