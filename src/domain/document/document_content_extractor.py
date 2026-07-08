from dataclasses import replace

from src.domain.document.character_count_port import CharacterCountPort
from src.domain.document.content_extraction_port import ContentExtractionPort
from src.domain.document.document_text_port import DocumentTextPort
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.exceptions.count_errors import CharacterCountUnavailable


class DocumentContentExtractor:
    """Domain service that extracts document content, applying count fallback logic."""

    def __init__(
        self,
        document_text_port: DocumentTextPort,
        content_extraction_port: ContentExtractionPort,
        character_count_port: CharacterCountPort,
    ) -> None:
        self._document_text_port = document_text_port
        self._content_extraction_port = content_extraction_port
        self._character_count_port = character_count_port

    def extract_content(self, docx_path: str) -> DocumentContentDTO:
        """Extract document content, replacing counts with accurate values when available."""
        paragraphs = self._document_text_port.read_paragraphs(path=docx_path)
        base = self._content_extraction_port.extract(paragraphs=paragraphs, docx_path=docx_path)
        try:
            counts = self._character_count_port.count(docx_path=docx_path)
        except CharacterCountUnavailable:
            return base
        if counts is None:
            return base
        return replace(
            base,
            word_count=counts.word_count,
            char_count=counts.char_count,
            paragraph_count=counts.paragraph_count,
        )
