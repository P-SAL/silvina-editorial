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
        s4, s5, s6 = self._signal_s4_s5_s6(text_sample, document_content.title)

        signals = [s2a, s2b, s3, s4, s5, s6]

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

    # ── Signals 4, 5 and 6 (combined LLM call) ──────────────────────────────────

    def _signal_s4_s5_s6(self, text_sample: str, title: str) -> tuple[bool, bool, bool]:
        """
        Single LLM call returning (S4, S5, S6).
        S4: explicit research intent detected.
        S5: evidence-based conclusive contribution detected.
        S6: theoretical framework justification / knowledge gap identified.
        """
        prompt = f"""Analiza el siguiente fragmento de un artículo académico.

TÍTULO: {title}

TEXTO:
{text_sample}

TAREA: Responde TRES preguntas independientes sobre el texto.

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

PREGUNTA 3 — JUSTIFICACIÓN TEÓRICA (S6):
¿El artículo justifica la selección de su marco teórico o identifica un vacío en el conocimiento existente que su investigación aborda mediante CUALQUIERA de estas formas?
- Referencia al estado del arte o literatura previa: "estudios previos han demostrado", "la literatura indica", "la literatura previa señala", "estudios previos muestran"
- Identificación de vacío: "sin embargo, no se ha explorado", "los estudios existentes no abordan", "existe un vacío en la literatura", "vacío en el conocimiento"
- Justificación del marco teórico: "se adopta el enfoque X porque", "este marco permite", "se seleccionó esta metodología porque"
- Anclaje en investigación previa: "a diferencia de estudios anteriores", "extendiendo el trabajo de", "en línea con"

FORMATO DE RESPUESTA — escribe ÚNICAMENTE estas tres líneas, sin encabezados, sin explicaciones:
S4: SI o NO
S5: SI o NO
S6: SI o NO"""

        try:
            response = self.client.generate(
                model=self.model_name,
                prompt=prompt,
                options={"temperature": 0.1, "num_predict": 300}
            )
            raw = response["response"].strip()
            import re
            raw_upper = raw.upper()
            s4 = bool(re.search(r'S4\s*:\s*SI', raw_upper))
            s5 = bool(re.search(r'S5\s*:\s*SI', raw_upper))
            s6 = bool(re.search(r'S6\s*:\s*SI', raw_upper))
                        
            return s4, s5, s6

        except Exception as e:
            print(f"⚠️  Signals S4/S5/S6 (combined) error: {e}")
            return False, False, False
    
    
    # ══════════════════════════════════════════════════════════════════════════
    # CLASSIFICATION RULE
    # ══════════════════════════════════════════════════════════════════════════

    def _apply_rule(self, signals: list[bool],
                    article_size) -> ClassificationResult:
        """
        Classification rule — v0.9 revised (S6 added).
        Reference: 19-case table, session May 2026.

        CIENTÍFICO requires S3+S4+S5 + structural support (confidence >= 0.83):
          case 2: S3+S4+S5+S2a+S2b+S6  → 0.90
          case 3: S3+S4+S5+S2b+S6       → 0.86
          case 4: S3+S4+S5+S2a+S2b      → 0.85
          case 5: S3+S4+S5+S2a+S6       → 0.83

        DIVULGACIÓN near-miss (cases 6–9): S3+S4+S5 present but below threshold.
        DIVULGACIÓN standard (cases 10–18): missing qualitative core signals.
        OPINIÓN (case 19): no signals detected.

        Confidence levels apply exclusively to CIENTÍFICO.
        DIVULGACIÓN and OPINIÓN carry confidence=None.
        """
        s2a, s2b, s3, s4, s5, s6 = signals

        signal_labels = {
            "Referencias ≥ 12":         s2a,
            "Referencias recientes":     s2b,
            "Vocabulario metodológico":  s3,
            "Intención investigativa":   s4,
            "Contribución conclusiva":   s5,
            "Justificación teórica":     s6,
        }
        active   = [label for label, val in signal_labels.items() if val]
        inactive = [label for label, val in signal_labels.items() if not val]

        def _sig(present: list, absent: list) -> str:
            parts = []
            if present: parts.append(f"Señales presentes: {', '.join(present)}.")
            if absent:  parts.append(f"Señales ausentes: {', '.join(absent)}.")
            return " ".join(parts)

        # ── CIENTÍFICO paths — S3+S4+S5 required ─────────────────────────────
        if s3 and s4 and s5:

            if s2a and s2b and s6:                              # case 2 — 0.90
                return ClassificationResult(
                    article_type=ArticleType.CIENTIFICO,
                    article_size=article_size, confidence=0.90,
                    reasoning=(
                        "El artículo reúne la totalidad de los indicadores científicos: "
                        "vocabulario metodológico (S3), intención investigativa (S4), "
                        "contribución evidenciada (S5), justificación del marco teórico y "
                        "vacío en la literatura (S6), cantidad de referencias suficiente (S2a) "
                        "y bibliografía actualizada (S2b). Artículo científico con muy elevada "
                        "confianza. " + _sig(active, inactive)
                    )
                )
            elif s2b and s6:                                    # case 3 — 0.86
                return ClassificationResult(
                    article_type=ArticleType.CIENTIFICO,
                    article_size=article_size, confidence=0.86,
                    reasoning=(
                        "Vocabulario metodológico (S3), intención investigativa (S4), "
                        "contribución evidenciada (S5) y justificación teórica (S6) presentes. "
                        "Bibliografía reciente (S2b), aunque por debajo del umbral de cantidad "
                        "mínima (S2a ausente). Artículo científico con confianza elevada. "
                        + _sig(active, inactive)
                    )
                )
            elif s2a and s2b:                                   # case 4 — 0.85
                return ClassificationResult(
                    article_type=ArticleType.CIENTIFICO,
                    article_size=article_size, confidence=0.85,
                    reasoning=(
                        "Vocabulario metodológico (S3), intención investigativa (S4), "
                        "contribución evidenciada (S5) y respaldo bibliográfico completo en "
                        "cantidad y actualidad (S2a, S2b). No se detectó justificación del "
                        "marco teórico ni identificación de vacío en la literatura (S6 ausente). "
                        "Artículo científico de rigor metodológico; calificación de confianza "
                        "media por ausencia de S6. " + _sig(active, inactive)
                    )
                )
            elif s2a and s6:                                    # case 5 — 0.83
                return ClassificationResult(
                    article_type=ArticleType.CIENTIFICO,
                    article_size=article_size, confidence=0.83,
                    reasoning=(
                        "Vocabulario metodológico (S3), intención investigativa (S4), "
                        "contribución evidenciada (S5) y justificación teórica (S6) presentes. "
                        "Cantidad de referencias suficiente (S2a). La bibliografía no alcanza "
                        "el umbral de actualidad requerido (S2b ausente). Artículo científico "
                        "de rigor metodológico; calificación de confianza media por ausencia "
                        "de S2b. " + _sig(active, inactive)
                    )
                )

            # ── Near-miss: S3+S4+S5 present but below 0.83 threshold ─────────
            _rec = "Revisión editorial recomendada: "
            if s6:                                              # case 6
                body = (
                    "El artículo muestra indicadores cualitativos sólidos (S3, S4, S5, S6), "
                    "pero carece del respaldo bibliográfico mínimo requerido "
                    "(S2a y S2b ausentes). "
                )
                rec = (
                    _rec + "con la incorporación de respaldo bibliográfico suficiente en "
                    "cantidad y actualidad, el artículo podría alcanzar el umbral para "
                    "artículo científico."
                )
            elif s2b:                                           # case 7
                body = (
                    "Vocabulario metodológico, intención investigativa y contribución "
                    "evidenciada presentes (S3, S4, S5), con bibliografía reciente (S2b). "
                    "Sin justificación del marco teórico (S6) ni cantidad suficiente "
                    "de referencias (S2a). "
                )
                rec = (
                    _rec + "con la incorporación de justificación del marco teórico (S6) "
                    "y ampliación del número de referencias (S2a), el artículo podría "
                    "alcanzar el umbral para artículo científico."
                )
            elif s2a:                                           # case 8
                body = (
                    "Vocabulario metodológico, intención investigativa y contribución "
                    "evidenciada presentes (S3, S4, S5), con cantidad de referencias "
                    "suficiente (S2a). Sin justificación del marco teórico (S6) ni "
                    "actualidad bibliográfica (S2b). "
                )
                rec = (
                    _rec + "con la incorporación de justificación del marco teórico (S6) "
                    "y actualización de la bibliografía (S2b), el artículo podría alcanzar "
                    "el umbral para artículo científico."
                )
            else:                                               # case 9
                body = (
                    "Vocabulario metodológico, intención investigativa y contribución "
                    "evidenciada presentes (S3, S4, S5). Sin justificación del marco "
                    "teórico (S6) ni respaldo bibliográfico (S2a, S2b). Las señales "
                    "cualitativas sin soporte estructural son insuficientes para "
                    "artículo científico. "
                )
                rec = (
                    _rec + "con la incorporación de justificación del marco teórico (S6) "
                    "y fortalecimiento del respaldo bibliográfico en cantidad y actualidad "
                    "(S2a, S2b), el artículo podría alcanzar el umbral para artículo "
                    "científico."
                )
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=body + rec + " " + _sig(active, inactive)
            )

        # ── DIVULGACIÓN standard (cases 10–18) ───────────────────────────────
        if s3 and s4:                                           # case 10
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Vocabulario metodológico (S3) e intención investigativa (S4) presentes. "
                    "No se detectó contribución basada en evidencia (S5 ausente). Sin los tres "
                    "pilares cualitativos completos, la clasificación como artículo científico "
                    "no es posible. " + _sig(active, inactive)
                )
            )
        if s3 and s5:                                           # case 11
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Vocabulario metodológico (S3) y contribución basada en evidencia (S5) "
                    "presentes. No se detectó intención investigativa explícita (S4 ausente). "
                    "Sin los tres pilares cualitativos completos, la clasificación como "
                    "artículo científico no es posible. " + _sig(active, inactive)
                )
            )
        if s3 and s2a and s2b:                                  # case 12
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Vocabulario metodológico (S3) y respaldo bibliográfico completo (S2a, S2b) "
                    "presentes. No se detectaron intención investigativa (S4) ni contribución "
                    "basada en evidencia (S5). Las señales cuantitativas sin pilares cualitativos "
                    "son insuficientes para artículo científico. " + _sig(active, inactive)
                )
            )
        if s3 and s2a:                                          # case 13
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Vocabulario metodológico (S3) y cantidad de referencias suficiente (S2a). "
                    "Sin intención investigativa (S4), contribución evidenciada (S5) ni "
                    "justificación teórica (S6). Evidencia insuficiente para artículo "
                    "científico. " + _sig(active, inactive)
                )
            )
        if s3 and s2b:                                          # case 14
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Vocabulario metodológico (S3) y bibliografía reciente (S2b). Sin intención "
                    "investigativa (S4), contribución evidenciada (S5) ni justificación teórica "
                    "(S6). Evidencia insuficiente para artículo científico. "
                    + _sig(active, inactive)
                )
            )
        if s3:                                                  # case 15
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Vocabulario metodológico presente (S3). Sin intención investigativa (S4), "
                    "contribución evidenciada (S5), justificación teórica (S6) ni respaldo "
                    "bibliográfico (S2a, S2b). El vocabulario técnico por sí solo es "
                    "insuficiente para clasificar como artículo científico. "
                    + _sig(active, inactive)
                )
            )
        if s4 and s5:                                           # case 16
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Intención investigativa (S4) y contribución evidenciada (S5) detectadas, "
                    "pero sin vocabulario metodológico formal (S3 ausente). El artículo carece "
                    "del sustento terminológico que distingue la investigación científica de la "
                    "divulgación especializada. " + _sig(active, inactive)
                )
            )
        if s4:                                                  # case 17
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Intención investigativa detectada (S4), pero sin vocabulario metodológico "
                    "(S3), contribución evidenciada (S5) ni respaldo bibliográfico. La sola "
                    "presencia de intención investigativa es insuficiente para artículo "
                    "científico. " + _sig(active, inactive)
                )
            )
        if s5:                                                  # case 18
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size, confidence=None,
                reasoning=(
                    "Contribución basada en evidencia detectada (S5), pero sin vocabulario "
                    "metodológico (S3) ni intención investigativa (S4). Una contribución "
                    "evidenciada sin proceso metodológico explícito no es suficiente para "
                    "artículo científico. " + _sig(active, inactive)
                )
            )

        # ── OPINIÓN: no signals (case 19) ─────────────────────────────────────
        return ClassificationResult(
            article_type=ArticleType.OPINION,
            article_size=article_size, confidence=None,
            reasoning=(
                "No se detectaron señales de investigación científica ni de divulgación "
                "especializada. El artículo expone puntos de vista, argumentos o reflexiones "
                "sin respaldo metodológico ni evidencia sistemática. "
                + _sig(active, inactive)
            )
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

    if result.article_type == ArticleType.CIENTIFICO:
        assert result.confidence is not None and result.confidence >= 0.83, \
            f"CIENTÍFICO confidence should be >= 0.83 — got {result.confidence}"

    conf_display = f"{result.confidence:.2f}" if result.confidence is not None else "—"
    print(f"✅ Smoke test passed — S3: {s3_result} | "
          f"Classification: {result.article_type.value} | "
          f"Confidence: {conf_display}")
    print(f"Reasoning: {result.reasoning}")
   