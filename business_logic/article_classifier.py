"""
article_classifier.py
Classifies articles into categories using a hybrid signal approach.
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
    "metodología", "hipótesis", "variables", "variable dependiente",
    "variable independiente", "cuantitativo", "cualitativo", "mixto", "diseño de investigación",
    "diseño experimental", "cuasi-experimental", "correlación", "regresión", "análisis estadístico",
    "significancia", "validez", "confiabilidad", "encuesta", "entrevista",
    "observación sistemática", "triangulación", "marco metodológico",
    "población", "unidad de análisis", "categorías de análisis", "codificación",
    "datos primarios", "datos secundarios", "trabajo de campo",
    # English (bilingual articles)
    "methodology", "sample", "hypothesis", "quantitative", "qualitative",
    "experimental design", "statistical analysis", "regression", "correlation",
    "validity", "reliability", "field work",
]


class ArticleClassifier:
    """Classifies academic articles using a hybrid signal approach."""

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

        # ── Signals 2a, 2b, 3, 4, 5 ──────────────────────────────────────────
        text_sample = self._build_text_sample(document_content)
        
        s2a = self._signal_reference_count(document_content)
        s2b = self._signal_reference_recency(document_content)
        s3  = self._signal_methodological_vocab(document_content)
        s4  = self._signal_research_question(text_sample, document_content.title)
        s5  = self._signal_conceptual_closure(text_sample, document_content.title)

        signals = [s2a, s2b, s3, s4, s5]

        # ── Signal 6: apply classification rule ───────────────────────────────
        return self._apply_rule(signals, article_size)

    # ══════════════════════════════════════════════════════════════════════════
    # SIGNAL HELPERS
    # ══════════════════════════════════════════════════════════════════════════

    def _build_text_sample(self, document_content: DocumentContent) -> str:
        """
        Return a text sample for LLM signals.
        Takes first 3500 chars (intro/research questions) +
        last 2500 chars (conclusions), skipping bibliography.
        """
        full_text = " ".join(document_content.paragraphs)

        # Exclude bibliography — detect first short standalone paragraph
        # containing a bibliography marker (≤ 30 chars = section header, not body prose)
        bib_markers = ['referencias', 'bibliografía', 'bibliography', 'fuentes bibliográficas']
        bib_pos = len(full_text)
        char_pos = 0
        for para in document_content.paragraphs:
            para_lower = para.strip().lower()
            if len(para_lower) <= 30 and any(marker in para_lower for marker in bib_markers):
                bib_pos = char_pos
                break
            char_pos += len(para) + 1

        clean_text = full_text[:bib_pos] if bib_pos > 0 else full_text
        if not clean_text:
            clean_text = full_text
        intro  = clean_text[:3500]
        ending = clean_text[-2500:] if len(clean_text) > 3500 else ""
        return (intro + " " + ending).strip() or full_text[:6000]

    # ── Signal 2a ─────────────────────────────────────────────────────────────

    def _signal_reference_count(self, document_content: DocumentContent) -> bool:
        """True if total references >= 15 (EUMIC minimum for científico)."""
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
        True if >= 4 distinct methodological terms found AND at least 1 hard term.
        S3 acts as tiebreaker only — see _apply_rule().
        """
        _HARD_TERMS = {
            "cuasi-experimental", "análisis estadístico", "triangulación",
            "marco metodológico", "unidad de análisis", "datos primarios",
            "datos secundarios", "statistical analysis", "experimental design",
            "diseño experimental", "diseño de investigación", "observación sistemática",
            "categorías de análisis"
        }

        full_text = " ".join(document_content.paragraphs).lower()
        found_terms = [term for term in _METHODOLOGICAL_VOCAB if term.lower() in full_text]
        found_hard  = [term for term in found_terms if term.lower() in _HARD_TERMS]

        return len(found_terms) >= 4 and len(found_hard) >= 1

    # ── Signal 4 ──────────────────────────────────────────────────────────────

    def _signal_research_question(self, text_sample: str, title: str) -> bool:
        """
        True if the LLM detects an explicit research question, objective,
        or central hypothesis.
        """
        prompt = f"""Analiza el siguiente fragmento de un artículo académico.

TÍTULO: {title}

TEXTO:
{text_sample}

TAREA: Determina si el artículo expresa explícitamente una intención de investigación mediante CUALQUIERA de estas formas:

1. Verbos de intención investigativa: examinar, analizar, identificar, determinar, explorar, comprender, evaluar, investigar, estudiar, revisar, sintetizar
2. Marcadores de alcance: "el presente estudio", "esta investigación", "la presente revisión", "el presente trabajo"
3. Marcadores de problema: "el problema central", "el objetivo es", "la pregunta que guía", "se busca responder"
4. Expresión de preguntas o hipótesis: una o múltiples, numeradas o no, directas o secuenciales

Responde SI si CUALQUIERA de estas formas está presente. Responde NO solo si ninguna está presente.

Responde SOLO con una de estas dos opciones (sin explicación adicional):
SI
NO"""
                  
        try:
            response = self.client.generate(
                model=self.model_name,
                prompt=prompt,
                options={"temperature": 0.1, "num_predict": 100}
            )
            answer = response["response"].strip().upper()
            return "SI" in answer[:100]
                        
        except Exception as e:
            print(f"⚠️  Signal 4 (research question) error: {e}")
            return False

    # ── Signal 5 ──────────────────────────────────────────────────────────────

    def _signal_conceptual_closure(self, text_sample: str, title: str) -> bool:
        """
        True if the LLM detects conclusions grounded in evidence, systematic
        analysis, systematic review, or academic literature synthesis.
        """
        prompt = f"""Analiza el siguiente fragmento de un artículo académico.

TÍTULO: {title}

TEXTO:
{text_sample}

TAREA: Determina si el artículo presenta conclusiones que exterioricen una contribución mediante CUALQUIERA de estas formas:

1. Hallazgos expresados como resultado de un proceso sistemático, replicable o verificable: "los resultados demuestran", "la evidencia indica", "el análisis revela", "se identificaron", "se observó que"
2. Propuesta de marco teórico, modelo, taxonomía, clasificación o esquema conceptual derivado del análisis
3. Recomendaciones específicas derivadas de evidencia o análisis sistemático, no de opinión personal
4. Identificación y abordaje de una brecha de conocimiento: "este estudio contribuye", "a diferencia de estudios previos", "se propone", "se demuestra que"
5. Síntesis que va más allá de la descripción: integra múltiples fuentes o perspectivas para arribar a una posición nueva

Responde SI si CUALQUIERA de estas formas está presente. Responde NO solo si las conclusiones son puramente descriptivas u opinativas sin sustento sistemático.

Responde SOLO con una de estas dos opciones (sin explicación adicional):
SI
NO"""

        try:
            response = self.client.generate(
                model=self.model_name,
                prompt=prompt,
                options={"temperature": 0.1, "num_predict": 100}
            )
            answer = response["response"].strip().upper()
            return "SI" in answer[:100]
                        
        except Exception as e:
            print(f"⚠️  Signal 5 (conceptual closure) error: {e}")
            return False

    # ══════════════════════════════════════════════════════════════════════════
    # SIGNAL 6 — CLASSIFICATION RULE
    # ══════════════════════════════════════════════════════════════════════════

    def _apply_rule(self, signals: list[bool],
                    article_size) -> ClassificationResult:
        """
        Classification rule — Option B:

        CIENTÍFICO:
          S4 + S5 + (S2a OR S2b) → CIENTÍFICO (0.85)

        DIVULGACIÓN:
          S4 + S5 (no S2)        → DIVULGACIÓN (0.75) — scientific intent, weak references
          S2a + S2b + S3 (no S4/S5) → DIVULGACIÓN (0.70) — referenced but no research intent
          S4 OR S5 alone         → DIVULGACIÓN (0.65) — partial scientific signal

        OPINIÓN:
          0 signals              → OPINIÓN (0.65)
        """
        s2a, s2b, s3, s4, s5 = signals

        signal_labels = {
            "Referencias ≥ 15":         s2a,
            "Referencias recientes":     s2b,
            "Vocabulario metodológico":  s3,
            "Pregunta de investigación": s4,
            "Contribución conclusiva":   s5,
        }
        active   = [label for label, val in signal_labels.items() if val]
        inactive = [label for label, val in signal_labels.items() if not val]

        def _reasoning(present: list[str], absent: list[str]) -> str:
            parts = []
            if present:
                parts.append(f"Señales presentes: {', '.join(present)}.")
            if absent:
                parts.append(f"Señales ausentes: {', '.join(absent)}.")
            return " ".join(parts)

        # ── CIENTÍFICO ────────────────────────────────────────────────────────
        if s4 and s5 and (s2a or s2b):
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.85,
                reasoning="Intención investigativa + contribución conclusiva + respaldo bibliográfico. "
                          + _reasoning(active, inactive)
            )

        # ── DIVULGACIÓN ───────────────────────────────────────────────────────
        if s4 and s5:
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size,
                confidence=0.75,
                reasoning="Intención investigativa y contribución presentes, sin respaldo bibliográfico suficiente. "
                          + _reasoning(active, inactive)
            )

        if s2a and s2b and s3:
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size,
                confidence=0.70,
                reasoning="Respaldo bibliográfico sólido con vocabulario metodológico, sin intención investigativa explícita. "
                          + _reasoning(active, inactive)
            )

        if s4 or s5:
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size,
                confidence=0.65,
                reasoning="Señal científica parcial detectada. Artículo de divulgación académica. "
                          + _reasoning(active, inactive)
            )

        # ── OPINIÓN ───────────────────────────────────────────────────────────
        return ClassificationResult(
            article_type=ArticleType.OPINION,
            article_size=article_size,
            confidence=0.65,
            reasoning="Sin señales científicas detectadas. "
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
        ],
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
