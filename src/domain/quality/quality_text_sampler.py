import re

from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.enums.reference_line_marker import ReferenceLineMarker


class QualityTextSampler:
    """Builds a strategic text excerpt for LLM-based quality analysis."""

    _CONCLUSION_HEADER_PATTERN = re.compile(r"conclusi", re.IGNORECASE)

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
            document_content.paragraphs[middle_index : middle_index + self._middle_paragraph_count]
        )
        parts.extend(self._collect_conclusion_or_tail_paragraphs(document_content.paragraphs))

        text_sample = self._join_to_paragraph_boundary(parts)
        if len(text_sample.split()) < self._min_sample_word_count:
            return self._join_to_paragraph_boundary(document_content.paragraphs)
        return text_sample

    def _join_to_paragraph_boundary(self, paragraphs: list[str]) -> str:
        """Join paragraphs, completing in full the one that first reaches the character limit."""
        included: list[str] = []
        for paragraph in paragraphs:
            included.append(paragraph)
            if len(" ".join(included)) >= self._text_sample_character_limit:
                break
        return " ".join(included)

    def _collect_conclusion_or_tail_paragraphs(self, paragraphs: list[str]) -> list[str]:
        conclusion_paragraphs = []
        in_conclusion = False
        for paragraph in paragraphs:
            if self._CONCLUSION_HEADER_PATTERN.search(paragraph):
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
