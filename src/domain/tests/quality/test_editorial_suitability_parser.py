from unittest import TestCase

from src.domain.quality.editorial_suitability_parser import EditorialSuitabilityParser


class TestEditorialSuitabilityParserContribution(TestCase):
    def setUp(self):
        self.parser = EditorialSuitabilityParser()

    def test_case_insensitive_labels_are_recognized(self):
        raw = (
            "veredicto: sustentada\n"
            "contribucion: Aporta un marco novedoso para el análisis operativo.\n"
            "observacion: Sera ignorada porque el veredicto es sustentada.\n"
        )

        verdict, phrase, observation = self.parser.parse_contribution(raw)

        self.assertEqual(verdict, "SUSTENTADA")
        self.assertEqual(phrase, "Aporta un marco novedoso para el análisis operativo.")
        self.assertEqual(
            observation,
            "Contribución sustentada — Aporta un marco novedoso para el análisis operativo.",
        )

    def test_no_sustentada_verdict_sets_fixed_observation(self):
        raw = "VEREDICTO: NO SUSTENTADA\nOBSERVACION: Esto debe ser ignorado por completo.\n"

        verdict, _phrase, observation = self.parser.parse_contribution(raw)

        self.assertEqual(verdict, "NO SUSTENTADA")
        self.assertEqual(observation, "Sin contribución observada o declarada.")

    def test_parcial_verdict_sets_fixed_observation(self):
        raw = (
            "VEREDICTO: PARCIAL\n"
            "CONTRIBUCION: Se menciona un aporte pero sin desarrollo suficiente.\n"
        )

        verdict, _phrase, observation = self.parser.parse_contribution(raw)

        self.assertEqual(verdict, "PARCIAL")
        self.assertEqual(observation, "Contribución declarada pero no suficientemente sustentada.")

    def test_sustentada_without_phrase_uses_fallback_observation(self):
        raw = "VEREDICTO: SUSTENTADA\n"

        verdict, phrase, observation = self.parser.parse_contribution(raw)

        self.assertEqual(verdict, "SUSTENTADA")
        self.assertEqual(phrase, "")
        self.assertEqual(observation, "Contribución sustentada.")


class TestEditorialSuitabilityParserAlignment(TestCase):
    def setUp(self):
        self.parser = EditorialSuitabilityParser()

    def test_no_alineado_is_not_misdetected_as_alineado(self):
        raw = "VEREDICTO: NO ALINEADO\nJUSTIFICACION: No se identifica relacion con ninguna linea de investigacion vigente.\n"

        verdict, _lines, justification = self.parser.parse_alignment(raw)

        self.assertEqual(verdict, "NO ALINEADO")
        self.assertEqual(
            justification,
            "No se identifica relacion con ninguna linea de investigacion vigente.",
        )

    def test_parcialmente_alineado_is_not_misdetected_as_alineado(self):
        raw = "VEREDICTO: PARCIALMENTE ALINEADO\nLINEAS: Linea 3 (tecnologia).\n"

        verdict, lines, _justification = self.parser.parse_alignment(raw)

        self.assertEqual(verdict, "PARCIALMENTE ALINEADO")
        self.assertEqual(lines, "Linea 3 (tecnologia).")

    def test_long_justification_truncated_at_word_boundary_with_ellipsis(self):
        raw = f"VEREDICTO: ALINEADO\nJUSTIFICACION: {'WORD ' * 30}END.\n"

        verdict, _lines, justification = self.parser.parse_alignment(raw)

        expected_prefix = "WORD " * 22 + "WORD"
        self.assertEqual(verdict, "ALINEADO")
        self.assertEqual(justification, expected_prefix + "…")
        self.assertLess(len(justification), 120)
        self.assertFalse(justification.endswith("WOR…"))

    def test_long_lines_truncated_to_eighty_characters_at_word_boundary(self):
        raw = f"VEREDICTO: ALINEADO\nLINEAS: {'WORD ' * 20}END.\n"

        _verdict, lines, _justification = self.parser.parse_alignment(raw)

        expected_prefix = "WORD " * 14 + "WORD"
        self.assertEqual(lines, expected_prefix + "…")
        self.assertLess(len(lines), 80)
