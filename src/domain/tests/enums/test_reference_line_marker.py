from unittest import TestCase

from src.domain.enums.reference_line_marker import ReferenceLineMarker


class TestReferenceLineMarker(TestCase):
    def test_enum_has_exactly_four_members(self):
        self.assertEqual(len(ReferenceLineMarker), 4)

    def test_enum_members_have_expected_values(self):
        self.assertEqual(ReferenceLineMarker.HTTP.value, "http")
        self.assertEqual(ReferenceLineMarker.DOI.value, "doi.org")
        self.assertEqual(ReferenceLineMarker.HTTPS.value, "https")
        self.assertEqual(ReferenceLineMarker.ISBN.value, "ISBN")
