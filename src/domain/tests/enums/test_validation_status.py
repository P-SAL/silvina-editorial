from unittest import TestCase

from src.domain.enums.validation_status import ValidationStatus


class TestValidationStatus(TestCase):
    def test_members_and_values(self):
        self.assertEqual(ValidationStatus.PASSED.value, "passed")
        self.assertEqual(ValidationStatus.FAILED.value, "failed")
        self.assertEqual(ValidationStatus.WARNING.value, "warning")
        self.assertEqual(ValidationStatus.NOT_APPLICABLE.value, "not_applicable")

    def test_member_count(self):
        self.assertEqual(len(ValidationStatus), 4)
