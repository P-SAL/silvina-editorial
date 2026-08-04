from src.domain.dtos.document_content_dto import DocumentContentDTO


class ArticleClassificationTextSampler:
    """Builds a strategic text excerpt for LLM-based article-classification signals."""

    _INTRODUCTION_CHARACTER_LIMIT = 3500
    _CONCLUSION_CHARACTER_LIMIT = 2500
    _FALLBACK_CHARACTER_LIMIT = 6000
    _BIBLIOGRAPHY_HEADER_MAX_LENGTH = 30
    _BIBLIOGRAPHY_MARKERS = (
        "referencias",
        "bibliografía",
        "bibliography",
        "fuentes bibliográficas",
    )

    def build_sample(self, document_content: DocumentContentDTO) -> str:
        """Return the first 3500 + last 2500 chars of the document, skipping the bibliography."""
        full_text = " ".join(document_content.paragraphs)
        clean_text = self._strip_bibliography(document_content.paragraphs, full_text)

        introduction = clean_text[: self._INTRODUCTION_CHARACTER_LIMIT]
        ending = (
            clean_text[-self._CONCLUSION_CHARACTER_LIMIT :]
            if len(clean_text) > self._INTRODUCTION_CHARACTER_LIMIT
            else ""
        )
        sample = (introduction + " " + ending).strip()
        return sample or full_text[: self._FALLBACK_CHARACTER_LIMIT]

    def _strip_bibliography(self, paragraphs: list[str], full_text: str) -> str:
        bibliography_position = len(full_text)
        character_position = 0
        for paragraph in paragraphs:
            paragraph_lower = paragraph.strip().lower()
            is_bibliography_header = len(
                paragraph_lower
            ) <= self._BIBLIOGRAPHY_HEADER_MAX_LENGTH and any(
                marker in paragraph_lower for marker in self._BIBLIOGRAPHY_MARKERS
            )
            if is_bibliography_header:
                bibliography_position = character_position
                break
            character_position += len(paragraph) + 1

        return full_text[:bibliography_position] if bibliography_position > 0 else full_text
