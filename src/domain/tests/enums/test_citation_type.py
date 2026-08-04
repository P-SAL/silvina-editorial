from unittest import TestCase

from src.domain.enums.citation_type import CitationType


class TestCitationType(TestCase):
    def test_members_and_values(self):
        self.assertEqual(CitationType.AUTHOR_YEAR.value, "author_year")
        self.assertEqual(CitationType.NUMERIC.value, "numeric")
        self.assertEqual(CitationType.FOOTNOTE.value, "footnote")
        self.assertEqual(CitationType.UNKNOWN.value, "unknown")

    def test_member_count(self):
        self.assertEqual(len(CitationType), 4)
