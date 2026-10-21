import unicodedata

from src.domain.dtos.document_content_dto import DocumentContentDTO


class MethodologicalVocabularyDetector:
    """Domain service that detects methodological-vocabulary signal S3 in a document."""

    _METHODOLOGICAL_VOCABULARY = (
        # Spanish — general methodology
        "metodología",
        "hipótesis",
        "variables",
        "variable dependiente",
        "variable independiente",
        "cuantitativo",
        "cualitativo",
        "mixto",
        "diseño de investigación",
        "diseño experimental",
        "cuasi-experimental",
        "correlación",
        "regresión",
        "análisis estadístico",
        "significancia",
        "validez",
        "confiabilidad",
        "encuesta",
        "entrevista",
        "observación sistemática",
        "triangulación",
        "marco metodológico",
        "población",
        "unidad de análisis",
        "categorías de análisis",
        "codificación",
        "datos primarios",
        "datos secundarios",
        "trabajo de campo",
        # Spanish — experimental design and simulation
        "laboratorio",
        "simulación",
        "maqueta",
        "escenario experimental",
        "se diseñó",
        "se implementó",
        "se simuló",
        "se construyó",
        "experimento",
        "prototipo",
        "banco de pruebas",
        # Spanish — quantitative results and validation
        "los resultados demuestran",
        "los experimentos mostraron",
        "los resultados obtenidos",
        "se validó",
        "se comprobó",
        "se verificó",
        "se demostró",
        "confirmando",
        "validando la hipótesis",
        "resultados preliminares",
        "tiempo para detectar",
        "tiempo para responder",
        "tasa de",
        "reducción del",
        "mejora del",
        "incremento del",
        # Spanish — metrics and measurement
        "métricas",
        "indicadores",
        "parámetros",
        "benchmark",
        "precisión",
        "recall",
        "exactitud",
        "rendimiento",
        # Spanish — qualitative social science
        "etnografía",
        "análisis del discurso",
        "teoría fundamentada",
        "análisis temático",
        "investigación acción",
        "estudio de caso",
        "análisis de contenido",
        "fenomenología",
        "hermenéutica",
        "narrativa",
        "interpretativo",
        "constructivismo",
        "saturación teórica",
        "muestreo teórico",
        "codificación axial",
        # Spanish — systematic review / evidence synthesis
        "revisión sistemática",
        "revision sistematica",
        "meta-análisis",
        "meta-analisis",
        "síntesis de evidencia",
        "criterios de inclusión",
        "criterios de exclusion",
        "búsqueda bibliográfica",
        "busqueda sistematica",
        "reproducibilidad",
        "protocolo de revisión",
        "seleccion de estudios",
        "extraccion de datos",
        # English (bilingual articles)
        "methodology",
        "sample",
        "hypothesis",
        "quantitative",
        "qualitative",
        "experimental design",
        "statistical analysis",
        "regression",
        "correlation",
        "validity",
        "reliability",
        "field work",
        "laboratory",
        "simulation",
        "experiment",
        "prototype",
        "results demonstrate",
        "experiments showed",
        "validated",
        "systematic review",
        "meta-analysis",
        "grounded theory",
        "thematic analysis",
        "discourse analysis",
    )
    _HARD_METHODOLOGICAL_TERMS = frozenset(
        {
            # quantitative
            "cuasi-experimental",
            "análisis estadístico",
            "triangulación",
            "marco metodológico",
            "unidad de análisis",
            "datos primarios",
            "datos secundarios",
            "statistical analysis",
            "experimental design",
            "diseño experimental",
            "diseño de investigación",
            "observación sistemática",
            "categorías de análisis",
            # experimental evidence
            "laboratorio",
            "simulación",
            "escenario experimental",
            "se validó",
            "se comprobó",
            "los resultados demuestran",
            "los experimentos mostraron",
            "validando la hipótesis",
            "tiempo para detectar",
            "tiempo para responder",
            "resultados preliminares",
            "banco de pruebas",
            # qualitative social science
            "teoría fundamentada",
            "análisis del discurso",
            "análisis temático",
            "saturación teórica",
            "codificación axial",
            "muestreo teórico",
            "grounded theory",
            "thematic analysis",
            "discourse analysis",
            "análisis de contenido",
            "fenomenología",
            "hermenéutica",
            # systematic review
            "revisión sistemática",
            "revision sistematica",
            "meta-análisis",
            "meta-analisis",
            "síntesis de evidencia",
        }
    )

    def __init__(self, minimum_term_count: int = 4, minimum_hard_term_count: int = 1) -> None:
        self._minimum_term_count = minimum_term_count
        self._minimum_hard_term_count = minimum_hard_term_count

    def detect(self, document_content: DocumentContentDTO) -> bool:
        """Return whether the document satisfies the methodological-vocabulary signal."""
        full_text_normalized = self._normalize_text(" ".join(document_content.paragraphs))
        found_terms = [
            term
            for term in self._METHODOLOGICAL_VOCABULARY
            if self._normalize_text(term) in full_text_normalized
        ]
        normalized_hard_terms = {
            self._normalize_text(term) for term in self._HARD_METHODOLOGICAL_TERMS
        }
        found_hard_terms = [
            term for term in found_terms if self._normalize_text(term) in normalized_hard_terms
        ]
        return (
            len(found_terms) >= self._minimum_term_count
            and len(found_hard_terms) >= self._minimum_hard_term_count
        )

    def _normalize_text(self, text: str) -> str:
        return unicodedata.normalize("NFD", text.lower()).encode("ascii", "ignore").decode()
