import dataclasses
from unittest import TestCase
from unittest.mock import MagicMock

from src.application.extract_content_use_case import ExtractContentUseCase
from src.domain.document.character_count_port import CharacterCountPort
from src.domain.dtos.character_count_dto import CharacterCountDTO
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.exceptions.count_errors import CharacterCountUnavailable
from src.domain.exceptions.document_errors import DocumentEmpty
from src.domain.tests.document.fake_character_count_port import FakeCharacterCountPort
from src.domain.tests.document.fake_content_extraction_port import FakeContentExtractionPort

_BASE_DTO = DocumentContentDTO(word_count=10, char_count=50, paragraph_count=3)
_ACCURATE = CharacterCountDTO(word_count=100, char_count=500, paragraph_count=10)


class TestExtractContentUseCase(TestCase):
    def test_without_path_returns_extraction_port_result(self):
        use_case = ExtractContentUseCase(
            extraction_port=FakeContentExtractionPort(result=_BASE_DTO),
            count_port=MagicMock(spec=CharacterCountPort),
        )

        result = use_case.execute(paragraphs=["para"])

        self.assertEqual(result, _BASE_DTO)

    def test_without_path_count_port_not_called(self):
        count_port = MagicMock(spec=CharacterCountPort)
        use_case = ExtractContentUseCase(
            extraction_port=FakeContentExtractionPort(result=_BASE_DTO),
            count_port=count_port,
        )

        use_case.execute(paragraphs=["para"])

        count_port.count.assert_not_called()

    def test_with_path_and_counts_merges_accurate_counts(self):
        use_case = ExtractContentUseCase(
            extraction_port=FakeContentExtractionPort(result=_BASE_DTO),
            count_port=FakeCharacterCountPort(result=_ACCURATE),
        )

        result = use_case.execute(paragraphs=["para"], docx_path="doc.docx")

        self.assertEqual(result.word_count, _ACCURATE.word_count)
        self.assertEqual(result.char_count, _ACCURATE.char_count)
        self.assertEqual(result.paragraph_count, _ACCURATE.paragraph_count)

    def test_with_path_and_counts_preserves_other_fields(self):
        base = dataclasses.replace(_BASE_DTO, title="My Title", abstract="My Abstract")
        use_case = ExtractContentUseCase(
            extraction_port=FakeContentExtractionPort(result=base),
            count_port=FakeCharacterCountPort(result=_ACCURATE),
        )

        result = use_case.execute(paragraphs=["para"], docx_path="doc.docx")

        self.assertEqual(result.title, "My Title")
        self.assertEqual(result.abstract, "My Abstract")

    def test_with_path_and_no_counts_returns_base_dto(self):
        use_case = ExtractContentUseCase(
            extraction_port=FakeContentExtractionPort(result=_BASE_DTO),
            count_port=FakeCharacterCountPort(result=None),
        )

        result = use_case.execute(paragraphs=["para"], docx_path="doc.docx")

        self.assertEqual(result, _BASE_DTO)

    def test_character_count_unavailable_falls_back_to_base_dto(self):
        use_case = ExtractContentUseCase(
            extraction_port=FakeContentExtractionPort(result=_BASE_DTO),
            count_port=FakeCharacterCountPort(error=CharacterCountUnavailable()),
        )

        result = use_case.execute(paragraphs=["para"], docx_path="doc.docx")

        self.assertEqual(result, _BASE_DTO)

    def test_character_count_unavailable_does_not_propagate(self):
        use_case = ExtractContentUseCase(
            extraction_port=FakeContentExtractionPort(result=_BASE_DTO),
            count_port=FakeCharacterCountPort(error=CharacterCountUnavailable()),
        )

        try:
            use_case.execute(paragraphs=["para"], docx_path="doc.docx")
        except CharacterCountUnavailable:
            self.fail("CharacterCountUnavailable propagated to caller")

    def test_document_empty_propagates(self):
        use_case = ExtractContentUseCase(
            extraction_port=FakeContentExtractionPort(error=DocumentEmpty()),
            count_port=FakeCharacterCountPort(),
        )

        with self.assertRaises(DocumentEmpty):
            use_case.execute(paragraphs=[])
