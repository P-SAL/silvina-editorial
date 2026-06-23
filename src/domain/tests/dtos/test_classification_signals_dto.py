from dataclasses import FrozenInstanceError
from unittest import TestCase

from src.domain.dtos.base_dto import BaseDTO
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


class TestClassificationSignalsDTO(TestCase):
    def test_classification_signals_is_subclass_of_base_dto(self):
        self.assertTrue(issubclass(ClassificationSignalsDTO, BaseDTO))

    def test_classification_signals_field_values_match_constructor_arguments(self):
        signals = ClassificationSignalsDTO(
            has_sufficient_reference_count=True,
            has_recent_references=False,
            has_methodological_vocabulary=True,
            has_research_intent=False,
            has_evidence_based_contribution=True,
            has_theoretical_justification=False,
        )

        self.assertTrue(signals.has_sufficient_reference_count)
        self.assertFalse(signals.has_recent_references)
        self.assertTrue(signals.has_methodological_vocabulary)
        self.assertFalse(signals.has_research_intent)
        self.assertTrue(signals.has_evidence_based_contribution)
        self.assertFalse(signals.has_theoretical_justification)

    def test_classification_signals_is_immutable(self):
        signals = ClassificationSignalsDTO(
            has_sufficient_reference_count=True,
            has_recent_references=True,
            has_methodological_vocabulary=True,
            has_research_intent=True,
            has_evidence_based_contribution=True,
            has_theoretical_justification=True,
        )

        with self.assertRaises(FrozenInstanceError):
            signals.has_sufficient_reference_count = False
