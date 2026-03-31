"""
article_classifier.py
Classifies articles into categories using a 6-signal hybrid approach.
Part of Silvina Editorial Assistant v0.9
"""

from typing import Optional
import ollama
from domain.models import DocumentContent, ClassificationResult
from domain.enums import ArticleType
from business_logic.structure_analyzer import StructureAnalyzer

# ─── Methodological vocabulary for Signal 3 ───────────────────────────────────
_METHODOLOGICAL_VOCAB = [
    # Spanish
    "metodología", "método", "muestra", "hipótesis", "variables", "variable dependiente",
    "variable independiente", "cuantitativo", "cualitativo", "mixto", "diseño de investigación",
    "diseño experimental", "cuasi-experimental", "correlación", "regresión", "análisis estadístico",
    "significancia", "validez", "confiabilidad", "instrumento", "encuesta", "entrevista",
    "observación sistemática", "triangulación", "marco teórico", "marco metodológico",
    "población", "unidad de análisis", "categorías de análisis", "codificación",
    "datos primarios", "datos secundarios", "trabajo de campo",
    # English (bilingual articles)
    "methodology", "sample", "hypothesis", "quantitative", "qualitative",
    "experimental design", "statistical analysis", "regression", "correlation",
    "validity", "reliability", "instrument", "survey", "field work",
]

class ArticleClassifier:
    """Classifies academic articles using a 6-signal hybrid approach."""

    def __init__(self, model_name: str = "llama3-gradient:8b-instruct-1048k-q4_K_M",
                 base_url: str = "http://localhost:11434"):
        self.model_name = model_name
        self.base_url = base_url
        self.client = ollama.Client(host=base_url)

    # ══════════════════════════════════════════════════════════════════════════
    # PUBLIC ENTRY POINT
    # ══════════════════════════════════════════════════════════════════════════

    def classify_article(self, document_content: DocumentContent) -> ClassificationResult:
        if not document_content or not document_content.paragraphs:
            raise ValueError("DocumentContent.paragraphs is empty")

        from domain.enums import classify_article_size
        article_size = classify_article_size(document_content.char_count)

        # ── Signal 1: IMRyD override (deterministic) ──────────────────────────
        structure = StructureAnalyzer().analyze(document_content)

        if structure["imryd_complete"] and article_size.name != "FUERA_RANGO":
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.95,
                reasoning="Estructura IMRyD completa detectada (override determinístico)."
            )

        # ── Signals 2-5: collect boolean results ──────────────────────────────
        text_sample = self._build_text_sample(document_content)

        s2 = self._signal_reference_density(document_content)
        s3 = self._signal_methodological_vocab(document_content)
        s4 = self._signal_research_question(text_sample, document_content.title)
        s5 = self._signal_conceptual_closure(text_sample, document_content.title)

        signals = [s2, s3, s4, s5]
        signal_count = sum(signals)

        # ── Signal 6: apply classification rule ───────────────────────────────
        return self._apply_rule(signals, signal_count, article_size)

    # ══════════════════════════════════════════════════════════════════════════
    # SIGNAL HELPERS
    # ══════════════════════════════════════════════════════════════════════════

    def _build_text_sample(self, document_content: DocumentContent) -> str:
        """Return a text sample suitable for LLM signals."""
        return " ".join(document_content.paragraphs[:60])[:7000]

    # ── Signal 2 ──────────────────────────────────────────────────────────────

    def _signal_reference_density(self, document_content: DocumentContent) -> bool:
        """
        True if reference density >= 3 references per 1000 words.
        Reads document_content.references (list[str]).
        """
        if document_content.word_count == 0:
            return False
        ref_count = len(document_content.references) if document_content.references else 0
        density = ref_count / (document_content.word_count / 1000)
        return density >= 3.0

    # ── Signal 3 ──────────────────────────────────────────────────────────────

    def _signal_methodological_vocab(self, document_content: DocumentContent) -> bool:
        """
        True if >= 3 distinct methodological terms are found in the full text.
        Case-insensitive scan of all paragraphs.
        """
        full_text = " ".join(document_content.paragraphs).lower()
        hits = sum(1 for term in _METHODOLOGICAL_VOCAB if term.lower() in full_text)
        return hits >= 3

    # ── Signal 4 ──────────────────────────────────────────────────────────────

    def _signal_research_question(self, text_sample: str, title: str) -> bool:
        """
        True if the LLM detects an explicit research question or objective
        (pregunta de investigación / objetivo de investigación).
        """
        prompt = f"""Analiza el siguiente fragmento de un artículo académico.

TÍTULO: {title}

TEXTO:
{text_sample}

TAREA: ¿El artículo formula explícitamente una pregunta de investigación, objetivo de investigación o hipótesis central?

Responde SOLO con una de estas dos opciones (sin explicación adicional):
SI
NO"""

        try:
            response = self.client.generate(
                model=self.model_name,
                prompt=prompt,
                options={"temperature": 0.1, "num_predict": 10}
            )
            answer = response["response"].strip().upper()
            return answer.startswith("SI")
        except Exception as e:
            print(f"⚠️  Signal 4 (research question) error: {e}")
            return False

    # ── Signal 5 ──────────────────────────────────────────────────────────────

    def _signal_conceptual_closure(self, text_sample: str, title: str) -> bool:
        """
        True if the LLM detects that the article reaches conclusions grounded
        in evidence or systematic analysis (not mere opinion).
        """
        prompt = f"""Analiza el siguiente fragmento de un artículo académico.

TÍTULO: {title}

TEXTO:
{text_sample}

TAREA: ¿El artículo presenta conclusiones basadas en evidencia, datos o análisis sistemático (no sólo en opinión del autor)?

Responde SOLO con una de estas dos opciones (sin explicación adicional):
SI
NO"""

        try:
            response = self.client.generate(
                model=self.model_name,
                prompt=prompt,
                options={"temperature": 0.1, "num_predict": 10}
            )
            answer = response["response"].strip().upper()
            return answer.startswith("SI")
        except Exception as e:
            print(f"⚠️  Signal 5 (conceptual closure) error: {e}")
            return False

    # ══════════════════════════════════════════════════════════════════════════
    # SIGNAL 6 — CLASSIFICATION RULE
    # ══════════════════════════════════════════════════════════════════════════

    def _apply_rule(self, signals: list[bool], signal_count: int,
                    article_size) -> ClassificationResult:
        """
        Apply the agreed classification rule to signals 2-5:
          4-5 signals → CIENTÍFICO  (0.85)
          3   signals → CIENTÍFICO  (0.70)
          2   signals → DIVULGACIÓN (0.75)
          0-1 signals → OPINIÓN     (0.65)
        """
        s2, s3, s4, s5 = signals
        signal_labels = {
            "Densidad de referencias": s2,
            "Vocabulario metodológico": s3,
            "Pregunta de investigación": s4,
            "Cierre conceptual": s5,
        }
        active = [label for label, val in signal_labels.items() if val]
        inactive = [label for label, val in signal_labels.items() if not val]

        def _reasoning(conf_signals: list[str], absent: list[str]) -> str:
            parts = []
            if conf_signals:
                parts.append(f"Señales presentes: {', '.join(conf_signals)}.")
            if absent:
                parts.append(f"Señales ausentes: {', '.join(absent)}.")
            return " ".join(parts)

        if signal_count >= 4:
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.85,
                reasoning=f"4 de 4 señales científicas confirmadas. "
                           + _reasoning(active, inactive)
            )

        if signal_count == 3:
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.70,
                reasoning=f"3 de 4 señales científicas confirmadas. "
                           + _reasoning(active, inactive)
            )

        if signal_count == 2:
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size,
                confidence=0.75,
                reasoning=f"2 de 4 señales científicas. Artículo de divulgación académica. "
                           + _reasoning(active, inactive)
            )

        # 0-1 signals
        return ClassificationResult(
            article_type=ArticleType.OPINION,
            article_size=article_size,
            confidence=0.65,
            reasoning=f"{'1' if signal_count == 1 else '0'} de 4 señales científicas. "
                       "Texto argumentativo u opinión sin validación empírica. "
                       + _reasoning(active, inactive)
        )


# ─── Convenience function ─────────────────────────────────────────────────────

def classify_document(document: DocumentContent,
                      model_name: str = "llama3-gradient:8b-instruct-1048k-q4_K_M") -> ClassificationResult:
    classifier = ArticleClassifier(model_name=model_name)
    return classifier.classify_article(document)


# ─── Quick smoke test ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    doc = DocumentContent(
        word_count=6000,
        char_count=35000,
        title="Análisis cuantitativo de la cohesión institucional en unidades militares conjuntas",
        abstract="Este estudio analiza los efectos de la integración conjunta.",
        references=[f"Referencia {i}" for i in range(25)],  # 25 refs → density 4.17/1000
        paragraphs=[
            "Este estudio analiza los efectos de la integración conjunta utilizando una metodología cuantitativa.",
            "La hipótesis central sostiene que la cohesión aumenta con la formación conjunta.",
            "Se aplicó una encuesta a una muestra de 240 oficiales. El diseño es cuasi-experimental.",
            "El análisis estadístico muestra correlación significativa (r=0.72, p<0.01).",
            "Los resultados validan la hipótesis y permiten formular recomendaciones institucionales.",
        ] * 10,
    )

    classifier = ArticleClassifier()
    result = classifier.classify_article(doc)
    print(result)
