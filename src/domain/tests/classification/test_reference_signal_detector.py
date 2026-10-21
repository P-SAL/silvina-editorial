from unittest import TestCase

from src.domain.classification.reference_signal_detector import ReferenceSignalDetector
from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.reference_dto import ReferenceDTO


class TestReferenceSignalDetector(TestCase):
    def setUp(self) -> None:
        self._detector = ReferenceSignalDetector()

    def test_reference_count_signal_fires_at_exactly_twelve_references(self) -> None:
        document_content = self._build_document_content(reference_count=12)

        self.assertTrue(self._detector.has_sufficient_count(document_content))

    def test_reference_count_signal_does_not_fire_at_eleven_references(self) -> None:
        document_content = self._build_document_content(reference_count=11)

        self.assertFalse(self._detector.has_sufficient_count(document_content))

    def test_reference_recency_signal_uses_maximum_year_per_reference(self) -> None:
        document_content = self._build_document_content(
            references=[ReferenceDTO(text="Autor, A. (1998). Cita previa de 2024.")]
        )

        self.assertTrue(self._detector.has_recent_majority(document_content))

    def test_no_references_yields_false_for_both_reference_signals(self) -> None:
        document_content = self._build_document_content(references=[])

        self.assertFalse(self._detector.has_sufficient_count(document_content))
        self.assertFalse(self._detector.has_recent_majority(document_content))

    def _build_document_content(
        self,
        reference_count: int | None = None,
        references: list[ReferenceDTO] | None = None,
    ) -> DocumentContentDTO:
        if references is None:
            references = (
                [
                    ReferenceDTO(text="Autor, A. (2020). Texto de referencia.")
                    for _ in range(reference_count)
                ]
                if reference_count is not None
                else []
            )
        return DocumentContentDTO(
            word_count=100,
            char_count=1000,
            paragraphs=[],
            references=references,
        )
