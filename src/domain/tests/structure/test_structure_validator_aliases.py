from unittest import TestCase

from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.enums.section_name import SectionName
from src.domain.structure.structure_validator import StructureValidator


def _make_document(paragraphs: list[str]) -> DocumentContentDTO:
    return DocumentContentDTO(word_count=0, char_count=0, paragraphs=paragraphs)


class TestStructureValidatorAliases(TestCase):
    def setUp(self):
        self.validator = StructureValidator()

    def test_english_alias_abstract_maps_to_resumen(self):
        present = self.validator._extract_present_sections(["Abstract"])
        self.assertIn(SectionName.SUMMARY, present)

    def test_multiple_aliases_detected(self):
        paragraphs = ["Metodologia", "Methodology", "Discussion", "Results"]
        present = self.validator._extract_present_sections(paragraphs)
        self.assertIn(SectionName.METHODOLOGY, present)
        self.assertIn(SectionName.DISCUSSION, present)
        self.assertIn(SectionName.RESULTS, present)

    def test_fuentes_bibliograficas_maps_to_referencias(self):
        present = self.validator._extract_present_sections(["Fuentes bibliográficas"])
        self.assertIn(SectionName.REFERENCES, present)

    def test_long_body_text_not_detected_as_header(self):
        present = self.validator._extract_present_sections(["x" * 100])
        self.assertEqual(present, [])

    def test_short_header_under_100_chars_detected(self):
        present = self.validator._extract_present_sections([SectionName.INTRODUCTION.value])
        self.assertIn(SectionName.INTRODUCTION, present)

    def test_inline_colon_keyword_detected_regardless_of_length(self):
        long_inline = "resumen: " + "x" * 95
        present = self.validator._extract_present_sections([long_inline])
        self.assertIn(SectionName.SUMMARY, present)

    def test_inline_keyword_with_space_before_colon_detected(self):
        present = self.validator._extract_present_sections(
            ["resumen : Este es el resumen del artículo"]
        )
        self.assertIn(SectionName.SUMMARY, present)

    def test_introduccion_without_accent_detected(self):
        present = self.validator._extract_present_sections(["introduccion"])
        self.assertIn(SectionName.INTRODUCTION, present)
