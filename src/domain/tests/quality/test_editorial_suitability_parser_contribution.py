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
