from unittest import TestCase

from src.domain.enums.apa_error_type import ApaErrorType


class TestApaErrorTypeEnum(TestCase):
    def test_conjunction_error_value(self):
        self.assertEqual(ApaErrorType.CONJUNCTION_ERROR, "Conjunción incorrecta")

    def test_comma_error_value(self):
        self.assertEqual(ApaErrorType.COMMA_ERROR, "Puntuación incorrecta")

    def test_capitalization_error_value(self):
        self.assertEqual(ApaErrorType.CAPITALIZATION_ERROR, "Mayúsculas/minúsculas incorrectas")

    def test_et_al_format_error_value(self):
        self.assertEqual(ApaErrorType.ET_AL_FORMAT_ERROR, "Formato 'et al.' incorrecto")

    def test_page_format_error_value(self):
        self.assertEqual(ApaErrorType.PAGE_FORMAT_ERROR, "Formato de página incorrecto")

    def test_spacing_error_value(self):
        self.assertEqual(ApaErrorType.SPACING_ERROR, "Espaciado incorrecto")

    def test_year_format_error_value(self):
        self.assertEqual(ApaErrorType.YEAR_FORMAT_ERROR, "Formato de año incorrecto")

    def test_parentheses_error_value(self):
        self.assertEqual(ApaErrorType.PARENTHESES_ERROR, "Paréntesis incorrectos")

    def test_all_8_members_exist(self):
        expected = {
            "CONJUNCTION_ERROR",
            "COMMA_ERROR",
            "CAPITALIZATION_ERROR",
            "ET_AL_FORMAT_ERROR",
            "PAGE_FORMAT_ERROR",
            "SPACING_ERROR",
            "YEAR_FORMAT_ERROR",
            "PARENTHESES_ERROR",
        }
        actual = {m.name for m in ApaErrorType}
        self.assertEqual(expected, actual)
