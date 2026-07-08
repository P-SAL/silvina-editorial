from unittest import TestCase

from src.domain.document.document_content_extractor import DocumentContentExtractor
from src.domain.dtos.character_count_dto import CharacterCountDTO
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.exceptions.count_errors import CharacterCountUnavailable
from src.domain.tests.document.fake_character_count_port import FakeCharacterCountPort
from src.domain.tests.document.fake_content_extraction_port import FakeContentExtractionPort
from src.domain.tests.document.fake_document_text_port import FakeDocumentTextPort


class TestDocumentContentExtractor(TestCase):
    def test_extract_content_returns_base_dto_when_counts_available(self):
        document_text_port = FakeDocumentTextPort(paragraphs=["Para 0", "Para 1"])
        base = DocumentContentDTO(word_count=1, char_count=1, paragraphs=["Para 0", "Para 1"])
        content_extraction_port = FakeContentExtractionPort(result=base)
        counts = CharacterCountDTO(word_count=10, char_count=100, paragraph_count=2)
        character_count_port = FakeCharacterCountPort(result=counts)

        extractor = DocumentContentExtractor(
            document_text_port=document_text_port,
            content_extraction_port=content_extraction_port,
            character_count_port=character_count_port,
        )
        result = extractor.extract_content(docx_path="test.docx")

        self.assertEqual(result.word_count, 10)
        self.assertEqual(result.char_count, 100)
        self.assertEqual(result.paragraph_count, 2)
        self.assertEqual(result.paragraphs, ["Para 0", "Para 1"])

    def test_extract_content_falls_back_to_base_when_character_count_unavailable(self):
        document_text_port = FakeDocumentTextPort(paragraphs=["Para 0"])
        base = DocumentContentDTO(word_count=1, char_count=1, paragraphs=["Para 0"])
        content_extraction_port = FakeContentExtractionPort(result=base)
        character_count_port = FakeCharacterCountPort(error=CharacterCountUnavailable())

        extractor = DocumentContentExtractor(
            document_text_port=document_text_port,
            content_extraction_port=content_extraction_port,
            character_count_port=character_count_port,
        )
        result = extractor.extract_content(docx_path="test.docx")

        self.assertIs(result, base)

    def test_extract_content_falls_back_to_base_when_counts_is_none(self):
        document_text_port = FakeDocumentTextPort(paragraphs=["Para 0"])
        base = DocumentContentDTO(word_count=1, char_count=1, paragraphs=["Para 0"])
        content_extraction_port = FakeContentExtractionPort(result=base)
        character_count_port = FakeCharacterCountPort(result=None)

        extractor = DocumentContentExtractor(
            document_text_port=document_text_port,
            content_extraction_port=content_extraction_port,
            character_count_port=character_count_port,
        )
        result = extractor.extract_content(docx_path="test.docx")

        self.assertIs(result, base)

    def test_extract_content_reads_paragraphs_and_passes_them_to_content_port(self):
        document_text_port = FakeDocumentTextPort(paragraphs=["Alpha", "Beta"])
        received_calls: list[tuple[list[str], str | None]] = []

        class RecordingContentExtractionPort(FakeContentExtractionPort):
            def extract(
                self, paragraphs: list[str], docx_path: str | None = None
            ) -> DocumentContentDTO:
                received_calls.append((paragraphs, docx_path))
                return super().extract(paragraphs, docx_path)

        content_extraction_port = RecordingContentExtractionPort(
            result=DocumentContentDTO(word_count=0, char_count=0, paragraphs=["Alpha", "Beta"])
        )
        character_count_port = FakeCharacterCountPort(result=None)

        extractor = DocumentContentExtractor(
            document_text_port=document_text_port,
            content_extraction_port=content_extraction_port,
            character_count_port=character_count_port,
        )
        result = extractor.extract_content(docx_path="test.docx")

        self.assertEqual(received_calls, [(["Alpha", "Beta"], "test.docx")])
        self.assertEqual(result.paragraphs, ["Alpha", "Beta"])
