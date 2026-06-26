from dataclasses import replace

from src.domain.document.character_count_port import CharacterCountPort
from src.domain.document.content_extraction_port import ContentExtractionPort
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.exceptions.count_errors import CharacterCountUnavailable
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler


class ExtractContentUseCase:
    """Orchestrates content extraction and optional COM-based count refinement."""

    def __init__(
        self,
        extraction_port: ContentExtractionPort,
        count_port: CharacterCountPort,
    ) -> None:
        self._extraction_port = extraction_port
        self._count_port = count_port

    @generic_error_handler
    def execute(self, paragraphs: list[str], docx_path: str | None = None) -> DocumentContentDTO:
        """Extract structured content and optionally refine counts from the .docx file.

        When docx_path is None, returns text-based counts from the extraction port.
        When docx_path is provided and the count port succeeds, merges accurate counts
        into the result via dataclasses.replace; falls back to text-based counts otherwise.
        """
        base = self._extraction_port.extract(paragraphs, docx_path)
        if docx_path is None:
            return base
        try:
            counts = self._count_port.count(docx_path)
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
