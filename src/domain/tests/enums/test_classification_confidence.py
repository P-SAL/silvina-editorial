from unittest import TestCase

from src.domain.enums.classification_confidence import ClassificationConfidence


class TestClassificationConfidence(TestCase):
    def test_enum_has_exactly_five_members_with_english_names(self):
        self.assertEqual(len(ClassificationConfidence), 5)
        member_names = {member.name for member in ClassificationConfidence}
        self.assertEqual(
            member_names,
            {
                "IMRYD_OVERRIDE",
                "FULL_SIGNAL_MATCH",
                "RECENT_BIBLIOGRAPHY_SUPPORT",
                "COMPLETE_BIBLIOGRAPHY_SUPPORT",
                "SUFFICIENT_REFERENCE_COUNT",
            },
        )
        member_values = {member.value for member in ClassificationConfidence}
        self.assertEqual(member_values, {0.95, 0.90, 0.86, 0.85, 0.83})

    def test_enum_members_behave_as_plain_floats(self):
        self.assertEqual(ClassificationConfidence.IMRYD_OVERRIDE, 0.95)
        self.assertEqual(ClassificationConfidence.IMRYD_OVERRIDE * 2, 1.90)
