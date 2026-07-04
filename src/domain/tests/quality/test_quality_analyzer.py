from unittest import TestCase

from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.dtos.quality_level_thresholds_dto import QualityLevelThresholdsDTO
from src.domain.enums.quality_level import QualityLevel
from src.domain.exceptions.quality_errors import QualityAnalysisFailed
from src.domain.quality.quality_analyzer import QualityAnalyzer
from src.domain.quality.quality_level_resolver import QualityLevelResolver
from src.domain.quality.quality_response_parser import QualityResponseParser
from src.domain.quality.quality_text_sampler import QualityTextSampler
from src.domain.tests.quality.fake_llm_generator_adapter import FakeLlmGeneratorAdapter


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


CLARITY_COHERENCE_PROMPT_TEMPLATE = """Eres un revisor editorial académico experto.

TEXTO A ANALIZAR:
{text_sample}

Evalúa Claridad y Coherencia."""

ARGUMENTATION_CONCLUSIONS_PROMPT_TEMPLATE = """Eres un revisor editorial académico experto.

TEXTO A ANALIZAR:
{text_sample}

Evalúa Argumentación y Conclusiones."""


def build_analyzer(fake_adapter: FakeLlmGeneratorAdapter) -> QualityAnalyzer:
    return QualityAnalyzer(
        llm_generator=fake_adapter,
        text_sampler=QualityTextSampler(),
        response_parser=QualityResponseParser(),
        clarity_coherence_prompt_template=CLARITY_COHERENCE_PROMPT_TEMPLATE,
        argumentation_conclusions_prompt_template=ARGUMENTATION_CONCLUSIONS_PROMPT_TEMPLATE,
        resolver=QualityLevelResolver(
            thresholds=QualityLevelThresholdsDTO(
                excellent_threshold=9.0,
                good_threshold=7.0,
                acceptable_threshold=5.0,
                needs_improvement_threshold=3.0,
            )
        ),
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


class TestQualityAnalyzer(TestCase):
    def setUp(self):
        self.document_content = build_document_content(["Parrafo uno.", "Parrafo dos."])

    def test_generate_is_called_exactly_twice_per_analysis(self):
        fake_adapter = FakeLlmGeneratorAdapter([VALID_RESPONSE_ONE, VALID_RESPONSE_TWO])
        analyzer = build_analyzer(fake_adapter)

        analyzer.analyze(self.document_content, article_type=None)

        self.assertEqual(fake_adapter.call_count, 2)

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
        fake_adapter = FakeLlmGeneratorAdapter([response_one, response_two])
        analyzer = build_analyzer(fake_adapter)

        result = analyzer.analyze(self.document_content, article_type=None)

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
        fake_adapter = FakeLlmGeneratorAdapter([response_one, response_two])
        analyzer = build_analyzer(fake_adapter)

        result = analyzer.analyze(self.document_content, article_type=None)

        self.assertEqual(result.quality_level, QualityLevel.GOOD)

    def test_domain_service_has_zero_infrastructure_imports(self):
        from pathlib import Path

        source = Path("src/domain/quality/quality_analyzer.py").read_text(encoding="utf-8")

        self.assertNotIn("src.infrastructure", source)
        self.assertNotIn("import ollama", source)
        self.assertNotIn("from ollama", source)

    def test_claridad_and_coherencia_always_come_from_call_one(self):
        response_two_with_claridad_like_header = """**1. Argumentación** [Puntuación: 3/10]
Argumentos debiles presentados en el desarrollo del texto analizado aqui.

**2. Conclusiones** [Puntuación: 8/10]
Las conclusiones se desprenden claramente del contenido desarrollado.

**Claridad** [Puntuación: 1/10]
Este bloque de claridad nunca deberia usarse porque viene de la llamada dos.
"""
        fake_adapter = FakeLlmGeneratorAdapter(
            [VALID_RESPONSE_ONE, response_two_with_claridad_like_header]
        )
        analyzer = build_analyzer(fake_adapter)

        result = analyzer.analyze(self.document_content, article_type=None)

        self.assertEqual(result.dimension_scores["claridad"]["score"], 8.0)
        self.assertEqual(result.dimension_scores["coherencia"]["score"], 8.0)

    def test_argumentacion_and_conclusiones_always_come_from_call_two(self):
        fake_adapter = FakeLlmGeneratorAdapter([VALID_RESPONSE_ONE, VALID_RESPONSE_TWO])
        analyzer = build_analyzer(fake_adapter)

        result = analyzer.analyze(self.document_content, article_type=None)

        self.assertEqual(result.dimension_scores["argumentacion"]["score"], 8.0)
        self.assertEqual(result.dimension_scores["conclusiones"]["score"], 8.0)

    def test_both_dimensions_failing_to_parse_in_one_call_raises_quality_analysis_failed(self):
        response_one_without_headers = (
            "Este texto no contiene ningun encabezado de dimension reconocible."
        )
        fake_adapter = FakeLlmGeneratorAdapter([response_one_without_headers, VALID_RESPONSE_TWO])
        analyzer = build_analyzer(fake_adapter)

        with self.assertRaises(QualityAnalysisFailed):
            analyzer.analyze(self.document_content, article_type=None)

    def test_rendered_prompt_preserves_legacy_wording_with_sample_interpolated(self):
        fake_adapter = FakeLlmGeneratorAdapter([VALID_RESPONSE_ONE, VALID_RESPONSE_TWO])
        analyzer = build_analyzer(fake_adapter)

        analyzer.analyze(self.document_content, article_type=None)

        text_sample = QualityTextSampler().build_sample(self.document_content)
        self.assertIn(
            "Eres un revisor editorial académico experto.", fake_adapter.received_prompts[0]
        )
        self.assertIn(text_sample, fake_adapter.received_prompts[0])

    def test_quality_analyzer_module_defines_exactly_one_class(self):
        import ast
        from pathlib import Path

        source = Path("src/domain/quality/quality_analyzer.py").read_text(encoding="utf-8")
        tree = ast.parse(source)
        top_level_classes = [
            node.name for node in ast.iter_child_nodes(tree) if isinstance(node, ast.ClassDef)
        ]

        self.assertEqual(top_level_classes, ["QualityAnalyzer"])
