import dataclasses
from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.citation_dto import CitationDTO
from src.domain.dtos.citation_extraction_result_dto import CitationExtractionResultDTO
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.enums.citation_type import CitationType


class TestCitationExtractionResultDTO(TestCase):
    def _make_dto(self) -> CitationExtractionResultDTO:
        citation = CitationDTO(
            text="(Smith, 2020)", citation_type=CitationType.AUTHOR_YEAR, location=0
        )
        reference = ReferenceDTO(text="Smith, J. (2020). Title. Journal.")
        return CitationExtractionResultDTO(
            citations=[citation],
            references=[reference],
            section_type="Referencias",
        )

    def test_s4a_frozen_citations_field_raises_frozen_instance_error(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.citations = []

    def test_s4a_frozen_references_field_raises_frozen_instance_error(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.references = []

    def test_s4a_frozen_section_type_field_raises_frozen_instance_error(self):
        dto = self._make_dto()
        with self.assertRaises(FrozenInstanceError):
            dto.section_type = "Other"

    def test_s4b_all_fields_are_required_no_defaults(self):
        fields = dataclasses.fields(CitationExtractionResultDTO)
        for field in fields:
            if field.name in ("citations", "references", "section_type"):
                self.assertIs(field.default, dataclasses.MISSING)
                self.assertIs(field.default_factory, dataclasses.MISSING)
