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
    # Spanish — general methodology
    "metodología", "hipótesis", "variables", "variable dependiente",
    "variable independiente", "cuantitativo", "cualitativo", "mixto", "diseño de investigación",
    "diseño experimental", "cuasi-experimental", "correlación", "regresión", "análisis estadístico",
    "significancia", "validez", "confiabilidad", "encuesta", "entrevista",
    "observación sistemática", "triangulación", "marco metodológico",
    "población", "unidad de análisis", "categorías de análisis", "codificación",
    "datos primarios", "datos secundarios", "trabajo de campo",
    # Spanish — experimental design and simulation
    "laboratorio", "simulación", "maqueta", "escenario experimental",
    "se diseñó", "se implementó", "se simuló", "se construyó",
    "experimento", "prototipo", "banco de pruebas",
    # Spanish — quantitative results and validation
    "los resultados demuestran", "los experimentos mostraron", "los resultados obtenidos",
    "se validó", "se comprobó", "se verificó", "se demostró",
    "confirmando", "validando la hipótesis", "resultados preliminares",
    "tiempo para detectar", "tiempo para responder", "tasa de",
    "reducción del", "mejora del", "incremento del",
    # Spanish — metrics and measurement
    "métricas", "indicadores", "parámetros", "benchmark",
    "precisión", "recall", "exactitud", "rendimiento",
    # Spanish — qualitative social science
    "etnografía", "análisis del discurso", "teoría fundamentada", "análisis temático",
    "investigación acción", "estudio de caso", "análisis de contenido", "fenomenología",
    "hermenéutica", "narrativa", "interpretativo", "constructivismo",
    "saturación teórica", "muestreo teórico", "codificación axial",
    # Spanish — systematic review / evidence synthesis
    "revisión sistemática", "revision sistematica", "meta-análisis", "meta-analisis",
    "síntesis de evidencia", "criterios de inclusión", "criterios de exclusion",
    "búsqueda bibliográfica", "busqueda sistematica", "reproducibilidad",
    "protocolo de revisión", "seleccion de estudios", "extraccion de datos",
    # English (bilingual articles)
    "methodology", "sample", "hypothesis", "quantitative", "qualitative",
    "experimental design", "statistical analysis", "regression", "correlation",
    "validity", "reliability", "field work",
    "laboratory", "simulation", "experiment", "prototype",
    "results demonstrate", "experiments showed", "validated",
    "systematic review", "meta-analysis", "grounded theory",
    "thematic analysis", "discourse analysis",
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
        s4, s5 = self._signal_s4_s5(text_sample, document_content.title)

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
        """True if total references >= 12 (adjusted threshold — EUMIC guideline is 15)."""
        references = document_content.references or []
        return len(references) >= 12

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
        S3 acts as mandatory gate — see _apply_rule().
        Covers: quantitative, qualitative, experimental, systematic review methodologies.
        """
        _HARD_TERMS = {
            # quantitative
            "cuasi-experimental", "análisis estadístico", "triangulación",
            "marco metodológico", "unidad de análisis", "datos primarios",
            "datos secundarios", "statistical analysis", "experimental design",
            "diseño experimental", "diseño de investigación", "observación sistemática",
            "categorías de análisis",
            # experimental evidence
            "laboratorio", "simulación", "escenario experimental",
            "se validó", "se comprobó", "los resultados demuestran",
            "los experimentos mostraron", "validando la hipótesis",
            "tiempo para detectar", "tiempo para responder",
            "resultados preliminares", "banco de pruebas",
            # qualitative social science
            "teoría fundamentada", "análisis del discurso", "análisis temático",
            "saturación teórica", "codificación axial", "muestreo teórico",
            "grounded theory", "thematic analysis", "discourse analysis",
            "análisis de contenido", "fenomenología", "hermenéutica",
            # systematic review
            "revisión sistemática", "revision sistematica",
            "meta-análisis", "meta-analisis",
            "síntesis de evidencia",
        }

        import unicodedata

        def _normalize(s: str) -> str:
            return unicodedata.normalize("NFD", s.lower()).encode("ascii", "ignore").decode()

        full_text_norm = _normalize(" ".join(document_content.paragraphs))
        found_terms = [term for term in _METHODOLOGICAL_VOCAB if _normalize(term) in full_text_norm]
        found_hard  = [term for term in found_terms if _normalize(term) in {_normalize(t) for t in _HARD_TERMS}]

        return len(found_terms) >= 4 and len(found_hard) >= 1

    # ── Signals 4 and 5 (combined LLM call) ──────────────────────────────────

    def _signal_s4_s5(self, text_sample: str, title: str) -> tuple[bool, bool]:
        """
        Single LLM call returning (S4, S5).
        S4: explicit research intent detected.
        S5: evidence-based conclusive contribution detected.
        """
        prompt = f"""Analiza el siguiente fragmento de un artículo académico.

TÍTULO: {title}

TEXTO:
{text_sample}

TAREA: Responde DOS preguntas independientes sobre el texto.

PREGUNTA 1 — INTENCIÓN DE INVESTIGACIÓN (S4):
¿El artículo expresa explícitamente una intención de investigación mediante CUALQUIERA de estas formas?
- Verbos de intención: examinar, analizar, identificar, determinar, explorar, comprender, evaluar, investigar, estudiar, revisar, sintetizar
- Marcadores de alcance: "el presente estudio", "esta investigación", "la presente revisión", "el presente trabajo"
- Marcadores de problema: "el problema central", "el objetivo es", "la pregunta que guía", "se busca responder"
- Marcadores de propuesta experimental: "para fundamentar esta propuesta", "este trabajo combina", "este trabajo incluye", "a través de la simulación", "se propone demostrar", "se busca validar"
- Preguntas o hipótesis explícitas: una o múltiples, numeradas o no

PREGUNTA 2 — CONTRIBUCIÓN CONCLUSIVA (S5):
¿El artículo presenta conclusiones que exterioricen una contribución mediante CUALQUIERA de estas formas?
- Hallazgos de proceso sistemático: "los resultados demuestran", "la evidencia indica", "el análisis revela", "se identificaron"
- Propuesta de marco teórico, modelo, taxonomía o clasificación derivado del análisis
- Recomendaciones específicas derivadas de evidencia, no de opinión personal
- Identificación de brecha de conocimiento: "este estudio contribuye", "se propone", "se demuestra que"
- Síntesis que integra múltiples fuentes para arribar a una posición nueva
- Resultados experimentales cuantitativos: mejoras porcentuales, reducciones de tiempo, métricas de rendimiento
- Confirmación de hipótesis: "confirmando que", "lo que confirma", "los experimentos demostraron mejoras", "los resultados preliminares obtenidos fueron"

Responde EXACTAMENTE en este formato (dos líneas, sin nada más):
S4: SI
S5: SI"""

        try:
            response = self.client.generate(
                model=self.model_name,
                prompt=prompt,
                options={"temperature": 0.1, "num_predict": 20}
            )
            lines = response["response"].strip().upper().splitlines()
            s4 = any("S4" in l and "SI" in l for l in lines)
            s5 = any("S5" in l and "SI" in l for l in lines)
            return s4, s5

        except Exception as e:
            print(f"⚠️  Signals S4/S5 (combined) error: {e}")
            return False, False

    # ══════════════════════════════════════════════════════════════════════════
    # CLASSIFICATION RULE
    # ══════════════════════════════════════════════════════════════════════════

    def _apply_rule(self, signals: list[bool],
                    article_size) -> ClassificationResult:
        """
        Classification rule — Option B revised:

        CIENTÍFICO: three mandatory gates S3+S4+S5, S2 modulates confidence
          S3+S4+S5 + S2a + S2b → CIENTÍFICO 0.95  full bibliometric support
          S3+S4+S5 + S2b only  → CIENTÍFICO 0.88  recency confirmed
          S3+S4+S5 + S2a only  → CIENTÍFICO 0.80  volume confirmed
          S3+S4+S5 (no S2)     → CIENTÍFICO 0.72  methodological core only

        DIVULGACIÓN:
          S4+S5 (no S3)        → 0.75  intent present, no explicit methodology
          S4 OR S5 alone       → 0.65  partial scientific signal
          S3 alone             → 0.70  methodological vocabulary, no research intent

        OPINIÓN:
          no signals           → 0.65
        """
        s2a, s2b, s3, s4, s5 = signals

        signal_labels = {
            "Referencias ≥ 12":          s2a,
            "Referencias recientes":      s2b,
            "Vocabulario metodológico":   s3,
            "Pregunta de investigación":  s4,
            "Contribución conclusiva":    s5,
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

        # ── CIENTÍFICO: three mandatory gates S3+S4+S5 ───────────────────────
        if s3 and s4 and s5:
            if s2b and s2a:
                return ClassificationResult(
                    article_type=ArticleType.CIENTIFICO,
                    article_size=article_size,
                    confidence=0.95,
                    reasoning="Núcleo científico completo con pleno respaldo bibliométrico. "
                              + _reasoning(active, inactive)
                )
            if s2b:
                return ClassificationResult(
                    article_type=ArticleType.CIENTIFICO,
                    article_size=article_size,
                    confidence=0.88,
                    reasoning="Núcleo científico completo. Recencia de referencias confirmada; "
                              "cantidad por debajo del umbral EUMIC. "
                              + _reasoning(active, inactive)
                )
            if s2a:
                return ClassificationResult(
                    article_type=ArticleType.CIENTIFICO,
                    article_size=article_size,
                    confidence=0.80,
                    reasoning="Núcleo científico completo. Volumen de referencias adecuado; "
                              "recencia no confirmada. "
                              + _reasoning(active, inactive)
                )
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.72,
                reasoning="Núcleo metodológico mínimo presente. "
                          "Sin respaldo bibliométrico suficiente. "
                          + _reasoning(active, inactive)
            )

        # ── DIVULGACIÓN ───────────────────────────────────────────────────────
        if s4 and s5:
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size,
                confidence=0.75,
                reasoning="Intención investigativa y contribución conclusiva presentes, "
                          "sin metodología explícita (S3 ausente). "
                          + _reasoning(active, inactive)
            )

        if s3 and s5:
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.72,
                reasoning="Metodología experimental y contribución conclusiva detectadas. "
                          "Intención investigativa no confirmada en esta ejecución. "
                          "Clasificación CIENTÍFICO con confianza reducida — "
                          "se recomienda verificación editorial manual. "
                          + _reasoning(active, inactive)
            )

        if s3 and s4:
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.72,
                reasoning="Metodología experimental e intención investigativa detectadas. "
                          "Contribución conclusiva no confirmada en esta ejecución. "
                          "Clasificación CIENTÍFICO con confianza reducida — "
                          "se recomienda verificación editorial manual. "
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
       
        if s3 and (s2a or s2b):
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.70,
                reasoning="Vocabulario metodológico sólido con respaldo bibliográfico detectados. "
                          "Intención investigativa y contribución conclusiva no confirmadas en esta ejecución. "
                          "Clasificación preliminar CIENTÍFICO con confianza reducida — "
                          "se recomienda verificación editorial manual. "
                          + _reasoning(active, inactive)
            )

        if s3:
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.60,
                reasoning="Vocabulario metodológico detectado sin respaldo bibliográfico confirmado. "
                          "Intención investigativa y contribución conclusiva no confirmadas en esta ejecución. "
                          "Clasificación preliminar CIENTÍFICO con confianza muy reducida — "
                          "se recomienda verificación editorial manual. "
                          + _reasoning(active, inactive)
            )
        
                      
        # ── OPINIÓN ───────────────────────────────────────────────────────────
        return ClassificationResult(
            article_type=ArticleType.OPINION,
            article_size=article_size,
            confidence=0.65,
            reasoning="Sin pregunta de investigación ni contribución conclusiva. "
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

    # Assert S3 fires deterministically on methodological vocabulary
    s3_result = classifier._signal_methodological_vocab(doc)
    assert s3_result == True, f"S3 should fire on test paragraphs — got {s3_result}"

    # Assert classification is never OPINIÓN on a methodological document
    assert result.article_type != ArticleType.OPINION, \
        f"Should not be OPINIÓN on methodological doc — got {result.article_type}"

    print(f"✅ Smoke test passed — S3: {s3_result} | "
          f"Classification: {result.article_type.value} | "
          f"Confidence: {result.confidence}")
    print(f"Reasoning: {result.reasoning}")
