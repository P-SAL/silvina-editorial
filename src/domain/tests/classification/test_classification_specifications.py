from unittest import TestCase

from src.domain.classification.all_of_specification import AllOfSpecification
from src.domain.classification.has_evidence_based_contribution_specification import (
    HasEvidenceBasedContributionSpecification,
)
from src.domain.classification.has_methodological_vocabulary_specification import (
    HasMethodologicalVocabularySpecification,
)
from src.domain.classification.has_recent_references_specification import (
    HasRecentReferencesSpecification,
)
from src.domain.classification.has_research_intent_specification import (
    HasResearchIntentSpecification,
)
from src.domain.classification.has_sufficient_reference_count_specification import (
    HasSufficientReferenceCountSpecification,
)
from src.domain.classification.has_theoretical_justification_specification import (
    HasTheoreticalJustificationSpecification,
)
from src.domain.dtos.classification_signals_dto import ClassificationSignalsDTO


def _build_signals(**overrides: bool) -> ClassificationSignalsDTO:
    defaults = {
        "has_sufficient_reference_count": False,
        "has_recent_references": False,
        "has_methodological_vocabulary": False,
        "has_research_intent": False,
        "has_evidence_based_contribution": False,
        "has_theoretical_justification": False,
    }
    defaults.update(overrides)
    return ClassificationSignalsDTO(**defaults)


class TestHasSufficientReferenceCountSpecification(TestCase):
    def test_is_satisfied_when_signal_is_true(self) -> None:
        signals = _build_signals(has_sufficient_reference_count=True)

        self.assertTrue(HasSufficientReferenceCountSpecification().is_satisfied_by(signals))

    def test_is_not_satisfied_when_signal_is_false(self) -> None:
        signals = _build_signals(has_sufficient_reference_count=False)

        self.assertFalse(HasSufficientReferenceCountSpecification().is_satisfied_by(signals))


class TestHasRecentReferencesSpecification(TestCase):
    def test_is_satisfied_when_signal_is_true(self) -> None:
        signals = _build_signals(has_recent_references=True)

        self.assertTrue(HasRecentReferencesSpecification().is_satisfied_by(signals))

    def test_is_not_satisfied_when_signal_is_false(self) -> None:
        signals = _build_signals(has_recent_references=False)

        self.assertFalse(HasRecentReferencesSpecification().is_satisfied_by(signals))


class TestHasMethodologicalVocabularySpecification(TestCase):
    def test_is_satisfied_when_signal_is_true(self) -> None:
        signals = _build_signals(has_methodological_vocabulary=True)

        self.assertTrue(HasMethodologicalVocabularySpecification().is_satisfied_by(signals))

    def test_is_not_satisfied_when_signal_is_false(self) -> None:
        signals = _build_signals(has_methodological_vocabulary=False)

        self.assertFalse(HasMethodologicalVocabularySpecification().is_satisfied_by(signals))


class TestHasResearchIntentSpecification(TestCase):
    def test_is_satisfied_when_signal_is_true(self) -> None:
        signals = _build_signals(has_research_intent=True)

        self.assertTrue(HasResearchIntentSpecification().is_satisfied_by(signals))

    def test_is_not_satisfied_when_signal_is_false(self) -> None:
        signals = _build_signals(has_research_intent=False)

        self.assertFalse(HasResearchIntentSpecification().is_satisfied_by(signals))


class TestHasEvidenceBasedContributionSpecification(TestCase):
    def test_is_satisfied_when_signal_is_true(self) -> None:
        signals = _build_signals(has_evidence_based_contribution=True)

        self.assertTrue(HasEvidenceBasedContributionSpecification().is_satisfied_by(signals))

    def test_is_not_satisfied_when_signal_is_false(self) -> None:
        signals = _build_signals(has_evidence_based_contribution=False)

        self.assertFalse(HasEvidenceBasedContributionSpecification().is_satisfied_by(signals))


class TestHasTheoreticalJustificationSpecification(TestCase):
    def test_is_satisfied_when_signal_is_true(self) -> None:
        signals = _build_signals(has_theoretical_justification=True)

        self.assertTrue(HasTheoreticalJustificationSpecification().is_satisfied_by(signals))

    def test_is_not_satisfied_when_signal_is_false(self) -> None:
        signals = _build_signals(has_theoretical_justification=False)

        self.assertFalse(HasTheoreticalJustificationSpecification().is_satisfied_by(signals))


class TestAllOfSpecification(TestCase):
    def test_is_satisfied_when_every_specification_is_satisfied(self) -> None:
        signals = _build_signals(has_sufficient_reference_count=True, has_recent_references=True)
        specification = AllOfSpecification(
            HasSufficientReferenceCountSpecification(), HasRecentReferencesSpecification()
        )

        self.assertTrue(specification.is_satisfied_by(signals))

    def test_is_not_satisfied_when_one_specification_is_not_satisfied(self) -> None:
        signals = _build_signals(has_sufficient_reference_count=True, has_recent_references=False)
        specification = AllOfSpecification(
            HasSufficientReferenceCountSpecification(), HasRecentReferencesSpecification()
        )

        self.assertFalse(specification.is_satisfied_by(signals))

    def test_can_nest_another_all_of_specification(self) -> None:
        signals = _build_signals(
            has_sufficient_reference_count=True,
            has_recent_references=True,
            has_methodological_vocabulary=True,
        )
        inner = AllOfSpecification(
            HasSufficientReferenceCountSpecification(), HasRecentReferencesSpecification()
        )
        outer = AllOfSpecification(inner, HasMethodologicalVocabularySpecification())

        self.assertTrue(outer.is_satisfied_by(signals))
