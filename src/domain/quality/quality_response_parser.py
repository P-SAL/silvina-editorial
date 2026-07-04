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
_LIST_MARKER_PATTERN = re.compile(r"^[*+-]\s+")
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
    """Parses one LLM response into per-dimension scores and feedback."""

    def __init__(
        self,
        unscored_dimension_score: float = 7.0,
        unscored_dimension_feedback: str = "No disponible",
    ) -> None:
        self._unscored_dimension_score = unscored_dimension_score
        self._unscored_dimension_feedback = unscored_dimension_feedback

    def parse(self, text: str) -> ParsedResponseDTO:
        """Parse an LLM response into a ParsedResponseDTO of per-dimension scores."""
        scores = {
            dimension: DimensionScoreDTO(
                self._unscored_dimension_score, self._unscored_dimension_feedback
            )
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

            scores[dimension] = DimensionScoreDTO(score, feedback)
            matched_dimensions.add(dimension)

        return ParsedResponseDTO(scores=scores, matched_dimensions=frozenset(matched_dimensions))

    def _extract_feedback(self, block: str) -> str:
        lines = block.strip().split("\n")
        feedback_lines = [
            _LIST_MARKER_PATTERN.sub("", line.strip()) for line in lines[1:] if line.strip()
        ]
        feedback = " ".join(feedback_lines)
        feedback = _RECOMMENDATION_TAIL_PATTERN.sub("", feedback).strip()
        feedback = " ".join(feedback.split())

        if len(feedback) < 10:
            return self._unscored_dimension_feedback

        sentences = [s.strip() for s in feedback.split(".") if s.strip()]
        if len(sentences) > 3:
            return ". ".join(sentences[:3]) + "."
        return feedback

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

    def _map_block_to_dimension(self, block: str) -> QualityDimension | None:
        block_lower = block[:200].lower()
        for dimension, keywords in _DIMENSION_KEYWORDS:
            if any(keyword in block_lower for keyword in keywords):
                return dimension
        return None
