from unittest import TestCase

from src.domain.enums.quality_dimension import QualityDimension
from src.domain.quality.quality_response_parser import QualityResponseParser

VALID_RESPONSE_TWO = """**1. Argumentación** [Puntuación: 8/10]
Los argumentos presentados son solidos y estan bien fundamentados.

**2. Conclusiones** [Puntuación: 8/10]
Las conclusiones se desprenden claramente del contenido desarrollado.
"""


class TestQualityResponseParser(TestCase):
    def test_numbered_and_unnumbered_headers_both_parse_to_same_score(self):
        numbered_response = """**1. Claridad** [Puntuación: 8/10]
Texto de retroalimentacion suficientemente largo para superar el minimo.
"""
        unnumbered_response = """**Claridad** [Puntuación: 8/10]
Texto de retroalimentacion suficientemente largo para superar el minimo.
"""
        parser = QualityResponseParser()

        result_numbered = parser.parse(numbered_response)
        result_unnumbered = parser.parse(unnumbered_response)

        self.assertEqual(result_numbered.scores[QualityDimension.CLARITY].score, 8.0)
        self.assertEqual(result_unnumbered.scores[QualityDimension.CLARITY].score, 8.0)

    def test_score_inferred_from_narrative_when_explicit_score_absent(self):
        response = """**1. Claridad**
El argumento es bastante bueno y adecuado en su desarrollo general del tema.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        self.assertEqual(result.scores[QualityDimension.CLARITY].score, 7.5)

    def test_excelente_keyword_infers_eight_point_five(self):
        response = """**1. Claridad**
El trabajo es excelente y sobresaliente en su desarrollo argumentativo general.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        self.assertEqual(result.scores[QualityDimension.CLARITY].score, 8.5)

    def test_aceptable_keyword_infers_six_point_zero(self):
        response = """**1. Claridad**
El trabajo resulta aceptable y suficiente en su desarrollo argumentativo general.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        self.assertEqual(result.scores[QualityDimension.CLARITY].score, 6.0)

    def test_deficiente_keyword_infers_four_point_zero(self):
        response = """**1. Claridad**
El trabajo resulta deficiente y debil en su desarrollo argumentativo general.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        self.assertEqual(result.scores[QualityDimension.CLARITY].score, 4.0)

    def test_no_keyword_match_uses_neutral_default_score(self):
        response = """**1. Claridad**
Este texto no contiene ninguna palabra clave narrativa reconocida en absoluto.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        self.assertEqual(result.scores[QualityDimension.CLARITY].score, 7.0)

    def test_feedback_shorter_than_ten_characters_becomes_neutral_default(self):
        response = """**1. Claridad** [Puntuación: 8/10]
Corto.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        self.assertEqual(result.scores[QualityDimension.CLARITY].feedback, "No disponible")

    def test_feedback_longer_than_three_sentences_is_truncated(self):
        response = """**1. Claridad** [Puntuación: 8/10]
Primera oracion larga y descriptiva. Segunda oracion tambien larga. Tercera oracion mas. Cuarta oracion final. Quinta oracion sobrante.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        feedback = result.scores[QualityDimension.CLARITY].feedback
        sentence_count = len([s for s in feedback.split(".") if s.strip()])
        self.assertEqual(sentence_count, 3)
        self.assertTrue(feedback.endswith("."))

    def test_argumentacion_block_is_not_misclassified_as_claridad(self):
        response = """**1. Argumentación** [Puntuación: 8/10]
La argumentacion presenta un argumento solido y bien fundamentado en el texto.

**2. Conclusiones** [Puntuación: 8/10]
Las conclusiones se desprenden claramente del contenido desarrollado.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        self.assertEqual(
            result.scores[QualityDimension.ARGUMENTATION].feedback,
            "La argumentacion presenta un argumento solido y bien fundamentado en el texto.",
        )

    def test_markdown_list_markers_are_stripped_from_feedback(self):
        response = """**1. Argumentación** [Puntuación: 8/10]
* El autor identifica tres explicaciones mecanicistas del razonamiento emergente.
+ Es un buen resumen del trabajo de los autores en general.

**2. Conclusiones** [Puntuación: 8/10]
Las conclusiones se desprenden claramente del contenido desarrollado.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        feedback = result.scores[QualityDimension.ARGUMENTATION].feedback
        self.assertFalse(feedback.startswith("*"))
        self.assertNotIn("+ Es un buen resumen", feedback)
        self.assertIn("El autor identifica tres explicaciones mecanicistas", feedback)

    def test_one_missing_dimension_in_otherwise_valid_response_keeps_the_rest(self):
        response = """**1. Claridad** [Puntuación: 8/10]
El argumento central es claro y facil de seguir en todo el texto.
"""
        parser = QualityResponseParser()

        result = parser.parse(response)

        self.assertEqual(result.scores[QualityDimension.CLARITY].score, 8.0)
        self.assertEqual(result.scores[QualityDimension.COHERENCE].score, 7.0)
        self.assertEqual(result.scores[QualityDimension.COHERENCE].feedback, "No disponible")
        self.assertEqual(
            result.matched_dimensions,
            frozenset({QualityDimension.CLARITY}),
        )
