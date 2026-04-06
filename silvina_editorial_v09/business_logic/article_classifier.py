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

        s2a = self._signal_reference_count(document_content)
        s2b = self._signal_reference_recency(document_content)
        s3 = self._signal_methodological_vocab(document_content)
        s4 = self._signal_research_question(text_sample, document_content.title)
        s5 = self._signal_conceptual_closure(text_sample, document_content.title)
        
        signals = [s2a, s2b, s3, s4, s5]
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
    # ── Signal 2a ─────────────────────────────────────────────────────────────

    def _signal_reference_count(self, document_content: DocumentContent) -> bool:
        """
        True if total references >= 15 (EUMIC minimum for científico).
        """
        references = document_content.references or []
        return len(references) >= 15

    # ── Signal 2b ─────────────────────────────────────────────────────────────

    def _signal_reference_recency(self, document_content: DocumentContent) -> bool:
        """
        True if >= 50% of references are recent (year >= current_year - 4).
        Year extracted from Reference.text via regex.
        """
        import re
        from datetime import datetime

        references = document_content.references or []
        if not references:
            return False

        recent_threshold = datetime.now().year - 4
        year_pattern = re.compile(r'\b((?:19|20)\d{2})\b')

        recent_count = 0
        for ref in references:
            years = [int(y) for y in year_pattern.findall(ref.text)]
            if years and max(years) >= recent_threshold:
                recent_count += 1

        return (recent_count / len(references)) >= 0.5

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
        Apply the agreed classification rule to signals 2a-5 (5 total):
          5/5 → CIENTÍFICO  (0.90)
          4/5 → CIENTÍFICO  (0.80)
          2-3/5 → DIVULGACIÓN (0.75)
          1/5 → DIVULGACIÓN (0.60)
          0/5 → OPINIÓN     (0.65)
        """
        s2a, s2b, s3, s4, s5 = signals
        signal_labels = {
            "Referencias ≥ 15": s2a,
            "Referencias recientes": s2b,
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

        if signal_count == 5:
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.90,
                reasoning=f"5 de 5 señales científicas confirmadas. "
                           + _reasoning(active, inactive)
            )

        if signal_count == 4:
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.80,
                reasoning=f"4 de 5 señales científicas confirmadas. "
                           + _reasoning(active, inactive)
            )

        if signal_count in (2, 3):
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size,
                confidence=0.75,
                reasoning=f"{signal_count} de 5 señales científicas. Artículo de divulgación académica. "
                           + _reasoning(active, inactive)
            )

        if signal_count == 1:
            return ClassificationResult(
                article_type=ArticleType.OPINION,
                article_size=article_size,
                confidence=0.65,
                reasoning=f"1 de 5 señales científicas. Insuficiente evidencia académica para divulgación. "
                           + _reasoning(active, inactive)
            )

        # 0 signals
        return ClassificationResult(
            article_type=ArticleType.OPINION,
            article_size=article_size,
            confidence=0.65,
            reasoning="0 de 5 señales científicas. "
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
    from domain.models import Reference

    doc = DocumentContent(
        word_count=6000,
        char_count=35000,
        title="Análisis cuantitativo de la cohesión institucional en unidades militares conjuntas",
        abstract="Este estudio analiza los efectos de la integración conjunta.",
        references=[
            Reference(text=f"Autor, A. (202{i % 5}). Título de referencia {i}. Revista Académica, {i+1}, 1-10.")
            for i in range(20)
        ],  # 20 refs, all 2020-2024 → S2a=True, S2b=True
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

