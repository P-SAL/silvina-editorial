from unittest import TestCase

from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.enums.quality_level import QualityLevel
from src.domain.exceptions.quality_errors import QualityAnalysisFailed
from src.domain.quality.quality_analyzer import QualityAnalyzer


class FakeLlmGeneratorPort:
    def __init__(self, responses: list[str]) -> None:
        self._responses = responses
        self.call_count = 0
        self.received_prompts: list[str] = []

    def generate(self, prompt: str) -> str:
        self.received_prompts.append(prompt)
        response = self._responses[self.call_count]
        self.call_count += 1
        return response


def build_document_content(
    paragraphs: list[str], title: str | None = "Title"
) -> DocumentContentDTO:
    full_text = " ".join(paragraphs)
    return DocumentContentDTO(
        word_count=len(full_text.split()),
        char_count=len(full_text),
        paragraph_count=len(paragraphs),
        title=title,
        paragraphs=paragraphs,
    )


VALID_RESPONSE_ONE = """**1. Claridad del argumento** [Puntuación: 8/10]
El argumento central es claro y facil de seguir en todo el texto.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""

VALID_RESPONSE_TWO = """**1. Argumentación** [Puntuación: 8/10]
Los argumentos presentados son solidos y estan bien fundamentados.

**2. Conclusiones** [Puntuación: 8/10]
Las conclusiones se desprenden claramente del contenido desarrollado.
"""


class TestTextSampling(TestCase):
    def test_short_document_uses_full_text_fallback_instead_of_sample(self):
        paragraphs = ["Intro corta."] * 3 + ["Parrafo de relleno."] * 2 + ["Conclusion breve."]
        document_content = build_document_content(paragraphs)
        fake_port = FakeLlmGeneratorPort([VALID_RESPONSE_ONE, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        analyzer.analyze(document_content, article_type=None)

        full_text = " ".join(paragraphs)
        self.assertIn(full_text[:200], fake_port.received_prompts[0])

    def test_long_document_uses_strategic_sample_not_full_text(self):
        excluded_paragraph = "PARRAFO_EXCLUIDO_UNICO " + ("relleno " * 10)
        paragraphs = (
            [
                "Intro uno " + "palabra " * 60,
                "Intro dos " + "palabra " * 60,
                "Intro tres " + "palabra " * 60,
            ]
            + [excluded_paragraph]
            + ["Relleno extra uno " + "palabra " * 100]
            + ["Relleno extra dos " + "palabra " * 100]
            + ["Relleno medio uno " + "palabra " * 100]
            + ["Relleno medio dos " + "palabra " * 100]
            + ["Relleno extra tres " + "palabra " * 100]
            + ["Conclusion final " + "palabra " * 100]
        )
        document_content = build_document_content(paragraphs)
        fake_port = FakeLlmGeneratorPort([VALID_RESPONSE_ONE, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        analyzer.analyze(document_content, article_type=None)

        self.assertNotIn("PARRAFO_EXCLUIDO_UNICO", fake_port.received_prompts[0])


class TestConclusionDetection(TestCase):
    def test_conclusion_paragraphs_exclude_reference_like_lines(self):
        paragraphs = (
            ["Intro uno.", "Intro dos.", "Intro tres."]
            + ["Relleno extra uno " + "palabra " * 100]
            + ["Relleno medio uno " + "palabra " * 100]
            + ["Relleno medio dos " + "palabra " * 100]
            + ["Relleno extra dos " + "palabra " * 100]
            + ["En conclusion, el trabajo demuestra " + "palabra " * 100]
            + ["https://doi.org/10.1234 referencia bibliografica excluida " + "palabra " * 100]
            + ["Conclusion final reafirmada " + "palabra " * 100]
        )
        document_content = build_document_content(paragraphs)
        fake_port = FakeLlmGeneratorPort([VALID_RESPONSE_ONE, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        analyzer.analyze(document_content, article_type=None)

        self.assertNotIn("referencia bibliografica excluida", fake_port.received_prompts[0])


class TestHeaderFormats(TestCase):
    def test_numbered_and_unnumbered_headers_both_parse_to_same_score(self):
        numbered_response = """**1. Claridad** [Puntuación: 8/10]
Texto de retroalimentacion suficientemente largo para superar el minimo.
"""
        unnumbered_response = """**Claridad** [Puntuación: 8/10]
Texto de retroalimentacion suficientemente largo para superar el minimo.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])

        fake_port_numbered = FakeLlmGeneratorPort([numbered_response, VALID_RESPONSE_TWO])
        analyzer_numbered = QualityAnalyzer(fake_port_numbered)
        result_numbered = analyzer_numbered.analyze(document_content, article_type=None)

        fake_port_unnumbered = FakeLlmGeneratorPort([unnumbered_response, VALID_RESPONSE_TWO])
        analyzer_unnumbered = QualityAnalyzer(fake_port_unnumbered)
        result_unnumbered = analyzer_unnumbered.analyze(document_content, article_type=None)

        self.assertEqual(result_numbered.dimension_scores["claridad"]["score"], 8.0)
        self.assertEqual(result_unnumbered.dimension_scores["claridad"]["score"], 8.0)


class TestNarrativeScoreInference(TestCase):
    def test_score_inferred_from_narrative_when_explicit_score_absent(self):
        response_one = """**1. Claridad**
El argumento es bastante bueno y adecuado en su desarrollo general del tema.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["score"], 7.5)

    def test_excelente_keyword_infers_eight_point_five(self):
        response_one = """**1. Claridad**
El trabajo es excelente y sobresaliente en su desarrollo argumentativo general.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["score"], 8.5)

    def test_aceptable_keyword_infers_six_point_zero(self):
        response_one = """**1. Claridad**
El trabajo resulta aceptable y suficiente en su desarrollo argumentativo general.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["score"], 6.0)

    def test_deficiente_keyword_infers_four_point_zero(self):
        response_one = """**1. Claridad**
El trabajo resulta deficiente y debil en su desarrollo argumentativo general.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["score"], 4.0)

    def test_no_keyword_match_uses_neutral_default_score(self):
        response_one = """**1. Claridad**
Este texto no contiene ninguna palabra clave narrativa reconocida en absoluto.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["score"], 7.0)


class TestFeedbackExtraction(TestCase):
    def test_feedback_shorter_than_ten_characters_becomes_neutral_default(self):
        response_one = """**1. Claridad** [Puntuación: 8/10]
Corto.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["feedback"], "No disponible")

    def test_feedback_longer_than_three_sentences_is_truncated(self):
        response_one = """**1. Claridad** [Puntuación: 8/10]
Primera oracion larga y descriptiva. Segunda oracion tambien larga. Tercera oracion mas. Cuarta oracion final. Quinta oracion sobrante.

**2. Coherencia** [Puntuación: 8/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        feedback = result.dimension_scores["claridad"]["feedback"]
        sentence_count = len([s for s in feedback.split(".") if s.strip()])
        self.assertEqual(sentence_count, 3)
        self.assertTrue(feedback.endswith("."))


class TestDimensionMapping(TestCase):
    def test_argumentacion_block_is_not_misclassified_as_claridad(self):
        response_two = """**1. Argumentación** [Puntuación: 8/10]
La argumentacion presenta un argumento solido y bien fundamentado en el texto.

**2. Conclusiones** [Puntuación: 8/10]
Las conclusiones se desprenden claramente del contenido desarrollado.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([VALID_RESPONSE_ONE, response_two])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(
            result.dimension_scores["argumentacion"]["feedback"],
            "La argumentacion presenta un argumento solido y bien fundamentado en el texto.",
        )

    def test_one_missing_dimension_in_otherwise_valid_response_keeps_the_rest(self):
        response_one = """**1. Claridad** [Puntuación: 8/10]
El argumento central es claro y facil de seguir en todo el texto.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["score"], 8.0)
        self.assertEqual(result.dimension_scores["coherencia"]["score"], 7.0)
        self.assertEqual(result.dimension_scores["coherencia"]["feedback"], "No disponible")


class TestDirectPerCallAssignment(TestCase):
    def test_claridad_and_coherencia_always_come_from_call_one(self):
        response_two_with_claridad_like_header = """**1. Argumentación** [Puntuación: 3/10]
Argumentos debiles presentados en el desarrollo del texto analizado aqui.

**2. Conclusiones** [Puntuación: 8/10]
Las conclusiones se desprenden claramente del contenido desarrollado.

**Claridad** [Puntuación: 1/10]
Este bloque de claridad nunca deberia usarse porque viene de la llamada dos.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort(
            [VALID_RESPONSE_ONE, response_two_with_claridad_like_header]
        )
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["score"], 8.0)
        self.assertEqual(result.dimension_scores["coherencia"]["score"], 8.0)

    def test_argumentacion_and_conclusiones_always_come_from_call_two(self):
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([VALID_RESPONSE_ONE, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.dimension_scores["argumentacion"]["score"], 8.0)
        self.assertEqual(result.dimension_scores["conclusiones"]["score"], 8.0)


class TestFullPerCallParseFailure(TestCase):
    def test_both_dimensions_failing_to_parse_in_one_call_raises_quality_analysis_failed(self):
        response_one_without_headers = (
            "Este texto no contiene ningun encabezado de dimension reconocible."
        )
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one_without_headers, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        with self.assertRaises(QualityAnalysisFailed):
            analyzer.analyze(document_content, article_type=None)


class TestPortCallCountAndOverallScore(TestCase):
    def test_generate_is_called_exactly_twice_per_analysis(self):
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([VALID_RESPONSE_ONE, VALID_RESPONSE_TWO])
        analyzer = QualityAnalyzer(fake_port)

        analyzer.analyze(document_content, article_type=None)

        self.assertEqual(fake_port.call_count, 2)

    def test_overall_score_is_mean_of_four_dimension_scores(self):
        response_one = """**1. Claridad** [Puntuación: 8/10]
El argumento central es claro y facil de seguir en todo el texto.

**2. Coherencia** [Puntuación: 6/10]
Las ideas se conectan de forma parcial entre las distintas secciones del texto.
"""
        response_two = """**1. Argumentación** [Puntuación: 7/10]
Los argumentos presentados son razonables y estan fundamentados en el texto.

**2. Conclusiones** [Puntuación: 9/10]
Las conclusiones se desprenden claramente del contenido desarrollado en detalle.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, response_two])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.overall_score, 7.5)

    def test_overall_score_of_seven_resolves_to_good_quality_level(self):
        response_one = """**1. Claridad** [Puntuación: 7/10]
El argumento central es claro y facil de seguir en todo el texto.

**2. Coherencia** [Puntuación: 7/10]
Las ideas se conectan logicamente entre las distintas secciones del texto.
"""
        response_two = """**1. Argumentación** [Puntuación: 7/10]
Los argumentos presentados son razonables y estan fundamentados en el texto.

**2. Conclusiones** [Puntuación: 7/10]
Las conclusiones se desprenden claramente del contenido desarrollado en detalle.
"""
        document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])
        fake_port = FakeLlmGeneratorPort([response_one, response_two])
        analyzer = QualityAnalyzer(fake_port)

        result = analyzer.analyze(document_content, article_type=None)

        self.assertEqual(result.quality_level, QualityLevel.GOOD)

    def test_domain_service_has_zero_infrastructure_imports(self):
        from pathlib import Path

        source = Path("src/domain/quality/quality_analyzer.py").read_text(encoding="utf-8")

        self.assertNotIn("src.infrastructure", source)
        self.assertNotIn("import ollama", source)
        self.assertNotIn("from ollama", source)
