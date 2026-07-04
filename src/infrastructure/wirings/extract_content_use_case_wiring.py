from src.application.extract_content_use_case import ExtractContentUseCase
from src.domain.document.character_count_port import CharacterCountPort
from src.domain.document.content_extraction_port import ContentExtractionPort
from src.infrastructure.adapters.document.paragraph_content_adapter import ParagraphContentAdapter
from src.infrastructure.adapters.document.win32com_word_count_adapter import (
    Win32ComWordCountAdapter,
)


class ExtractContentUseCaseWiring:
    """Factory for building a ready-to-use ExtractContentUseCase."""

    def create_use_case(self) -> ExtractContentUseCase:
        return ExtractContentUseCase(
            extraction_port=self._get_extraction_port(),
            count_port=self._get_count_port(),
        )

    def _get_count_port(self) -> CharacterCountPort:
        return Win32ComWordCountAdapter()

    def _get_extraction_port(self) -> ContentExtractionPort:
        return ParagraphContentAdapter()
