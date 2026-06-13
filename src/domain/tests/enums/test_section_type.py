from unittest import TestCase

from src.domain.enums.section_type import SectionType


class TestSectionType(TestCase):
    def test_member_count_is_23(self):
        self.assertEqual(len(SectionType), 23)

    def test_all_bilingual_section_names_present(self):
        expected_members = [
            "TITLE",
            "RESUMEN",
            "ABSTRACT",
            "INTRODUCCION",
            "INTRODUCTION",
            "METODOLOGIA",
            "METHODOLOGY",
            "CONCLUSIONES",
            "CONCLUSIONS",
            "REFERENCIAS",
            "REFERENCES",
        ]
        member_names = [member.name for member in SectionType]
        for name in expected_members:
            self.assertIn(name, member_names)
