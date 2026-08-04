"""
quality_analyzer.py
Analyzes document quality across multiple SEMANTIC dimensions using LLM.
TIER 2 - Focuses on content quality, not grammar/spelling (that's Tier 1)
Part of Silvina Editorial Assistant v0.9
"""
from __future__ import annotations

import ollama
import re
from typing import Dict, Any
from domain.enums import ClassificationCategory, QualityLevel, get_quality_level_from_score
from domain.models import DocumentContent, QualityResult
from domain.models import QualityAnalysisResult


# ── Líneas de investigación FMC (reference material for Pertinencia) ─────────
_LINEAS_INVESTIGACION = (
    "1. Ámbitos de conflicto y su interacción — evolución de ámbitos de conflicto "
    "tradicionales y emergentes; operaciones militares conjuntas y combinadas; geografía "
    "militar, logística, doctrina, toma de decisiones, operaciones de paz, juegos de guerra, "
    "capacidades militares, amenazas y espacios jurisdiccionales.\n"
    "2. Estrategia, conflicto y política internacional — estrategias nacionales de defensa; "
    "articulación con política exterior; escenarios geopolíticos globales y regionales; "
    "fundamentos teóricos y dimensión aplicada del fenómeno estratégico.\n"
    "3. Recursos humanos para la defensa — gestión del capital humano en organizaciones de "
    "defensa; formación, liderazgo, factores humanos en IA, capacitación en tecnologías "
    "emergentes; aspectos psicosociales del personal militar.\n"
    "4. Dinámica de los conflictos en el ciberespacio — conflictos en el ciberespacio; "
    "convergencia ciber-física, ciber-resiliencia, geopolítica tecnológica, IA, tecnologías "
    "cuánticas, protección de infraestructuras críticas, guerra de información, marcos "
    "jurídicos internacionales.\n"
    "5. Conocimiento, gestión y protección de recursos estratégicos — identificación, "
    "evaluación y protección de recursos críticos para la defensa; recursos naturales, "
    "tecnológicos e industriales; Objetivos de Valor Estratégico (OVE); sustentabilidad "
    "en operaciones militares.\n"
    "6. Inteligencia en los diversos ámbitos de conflicto — producción de inteligencia en "
    "niveles estratégico, operacional y táctico; ciclo de inteligencia, IA, capacidades "
    "geoespaciales, fuentes abiertas digitales, estudios prospectivos, desinformación, "
    "convergencia humano-artificial.\n"
    "7. Ciencia, tecnología y producción en la defensa — desarrollo e implementación de "
    "tecnologías para la defensa; base industrial nacional, autonomía tecnológica, "
    "transferencia tecnológica, IA, computación cuántica, industria 4.0, sistemas autónomos, "
    "soberanía tecnológica."
)


class QualityAnalyzer:
    """Analyzes academic document quality across semantic dimensions."""

    def __init__(self, model_name: str = "gemma2:27b",
                 base_url: str = "http://localhost:11434"):
        self.model_name = model_name
        import ollama
        self.ollama = ollama
        self.base_url = base_url
        self.client = ollama.Client(host=self.base_url)

    # ══════════════════════════════════════════════════════════════════════════
    # PUBLIC ENTRY POINT
    # ══════════════════════════════════════════════════════════════════════════

    def analyze_quality(self, document_content, article_type) -> QualityAnalysisResult:
        print("      ⏳ Analizando con Ollama...")

        # Sample text - strategic sampling
        parts = []
        parts.append(document_content.title or "")
        parts.extend(document_content.paragraphs[:3])  # Intro
        mid = len(document_content.paragraphs) // 2
        parts.extend(document_content.paragraphs[mid:mid+2])  # Middle

        # Find conclusion section explicitly, fallback to last non-reference paragraphs
        conclusion_paras = []
        in_conclusion = False
        for p in document_content.paragraphs:
            if re.search(r'conclusi', p, re.IGNORECASE):
                in_conclusion = True
            if in_conclusion and not any(x in p[:80] for x in ['http', 'doi.org', 'https', 'ISBN']):
                conclusion_paras.append(p)
        if conclusion_paras:
            parts.extend(conclusion_paras[:3])
        else:
            non_ref = [p for p in document_content.paragraphs
                       if not any(x in p[:80] for x in ['http', 'doi.org', 'https', 'ISBN'])]
            parts.extend(non_ref[-2:])

        text_sample = ' '.join(parts)[:8000]
        # For short documents, use full text instead of sample
        if len(text_sample.split()) < 400:
            text_sample = ' '.join(document_content.paragraphs)[:8000]

        ollama_options = {
            'temperature': 0.2,
            'num_predict': 1000,
            'num_ctx': 4096,
            'repeat_penalty': 1.1,
            'timeout': 120
        }

        # ── CALL 1: Claridad + Coherencia ─────────────────────────────────
        prompt_1 = f"""Eres un revisor editorial académico experto. Analiza este fragmento en DOS dimensiones.

TEXTO A ANALIZAR:
{text_sample}

INSTRUCCIONES:
1. Evalúa SOLO lo que está presente en el texto
2. Sé específico: menciona qué funciona bien y qué necesita mejorar
3. La ortografía y gramática ya fueron verificadas - enfócate en el CONTENIDO

FORMATO DE RESPUESTA (OBLIGATORIO):

**1. Claridad del argumento** [Puntuación: X/10]
[Analiza si el argumento central es claro. ¿El lector entiende fácilmente el mensaje principal?]

**2. Coherencia** [Puntuación: X/10]
[Analiza si las ideas se conectan lógicamente. ¿Hay transiciones claras entre secciones?]

CRITERIOS: 9-10 Excelente | 7-8 Bueno | 5-6 Aceptable | 3-4 Deficiente | 0-2 Inaceptable
"""

        # ── CALL 2: Argumentación + Conclusiones ──────────────────────────
        prompt_2 = f"""Eres un revisor editorial académico experto. Analiza este fragmento en DOS dimensiones.

TEXTO A ANALIZAR:
{text_sample}

INSTRUCCIONES:
1. Evalúa SOLO lo que está presente en el texto
2. Para Conclusiones: si no hay sección formal, infiere del contenido final del texto
3. La ortografía y gramática ya fueron verificadas - enfócate en el CONTENIDO

FORMATO DE RESPUESTA (OBLIGATORIO):

**1. Argumentación** [Puntuación: X/10]
[Si hay argumentos, evalúa su calidad. Si no los hay, indícalo claramente y asigna una puntuación baja.]

**2. Conclusiones** [Puntuación: X/10]
[OBLIGATORIO: Evalúa siempre. Si no hay sección formal, analiza el párrafo final del texto y asigna puntuación.]

CRITERIOS: 9-10 Excelente | 7-8 Bueno | 5-6 Aceptable | 3-4 Deficiente | 0-2 Inaceptable
"""

        try:
            # Call 1
            response_1 = self.ollama.generate(
                model=self.model_name,
                prompt=prompt_1,
                options=ollama_options
            )
            text_1 = response_1.get('response', '').strip()
            print(f"      ✓ Llamada 1 completada: {len(text_1.split())} palabras")

            # Call 2
            response_2 = self.ollama.generate(
                model=self.model_name,
                prompt=prompt_2,
                options=ollama_options
            )
            text_2 = response_2.get('response', '').strip()
            print(f"      ✓ Llamada 2 completada: {len(text_2.split())} palabras")

            # Parse both responses
            scores_1 = self._parse_llm_response(text_1)
            scores_2 = self._parse_llm_response(text_2)

            # Merge: prefer whichever call returned real feedback for each dim
            scores = {}
            for dim in ["claridad", "coherencia", "argumentacion", "conclusiones"]:
                s1 = scores_1.get(dim, {"score": 7.0, "feedback": "No disponible"})
                s2 = scores_2.get(dim, {"score": 7.0, "feedback": "No disponible"})
                scores[dim] = s1 if s1["feedback"] != "No disponible" else s2

            overall = sum(d["score"] for d in scores.values()) / len(scores)
            quality_level = get_quality_level_from_score(overall)

            # Call 3: Idoneidad editorial
            print(f"      ⏳ Evaluando idoneidad editorial...")
            idoneidad = self._analyze_idoneidad(text_sample)
            print(f"      ✓ Llamada 3 completada")

            print(f"      ✓ Análisis generado: {overall:.1f}/10\n")

            return QualityAnalysisResult(
                overall_score=overall,
                quality_level=quality_level,
                dimension_scores=scores,
                idoneidad_editorial=idoneidad
            )

        except Exception as e:
            print(f"      ⚠️  Error en LLM: {e}")
            default = {
                d: {"score": 7.0, "feedback": "Análisis no disponible"}
                for d in ["claridad", "coherencia", "argumentacion", "conclusiones"]
            }
            return QualityAnalysisResult(
                overall_score=7.0,
                quality_level=QualityLevel.ACCEPTABLE,
                dimension_scores=default,
                idoneidad_editorial={}
            )

    # ══════════════════════════════════════════════════════════════════════════
    # CALL 1 & 2 PARSER
    # ══════════════════════════════════════════════════════════════════════════

    def _parse_llm_response(self, text: str) -> Dict[str, Dict[str, Any]]:
        """
        Extract feedback and scores from LLM response.
        Handles both numbered (**1. Dim) and unnumbered (**Dim) header formats
        to cope with LLM non-determinism.
        """
        result = {
            "claridad":      {"score": 7.0, "feedback": "No disponible"},
            "coherencia":    {"score": 7.0, "feedback": "No disponible"},
            "argumentacion": {"score": 7.0, "feedback": "No disponible"},
            "conclusiones":  {"score": 7.0, "feedback": "No disponible"}
        }

        # Split on numbered OR unnumbered dimension headers
        blocks = re.split(
            r'(?=\*\*(?:\d+\.\s*)?(?:Claridad|Coherencia|Argumentaci[oó]n|Conclusiones))',
            text.strip(), flags=re.IGNORECASE
        )

        for block in blocks:
            if not block.strip():
                continue

            # Search entire block for score
            score_match = re.search(
                r'\[Puntuaci[oó]n:\s*(\d+(?:\.\d+)?)(?:/10)?\]|(\d+(?:\.\d+)?)\s*/\s*10',
                block, re.IGNORECASE
            )
            if not score_match:
                block_lower = block.lower()
                if any(w in block_lower for w in ['excelente', 'sobresaliente', 'muy bueno']):
                    score = 8.5
                elif any(w in block_lower for w in ['bueno', 'adecuado', 'correcto']):
                    score = 7.5
                elif any(w in block_lower for w in ['aceptable', 'suficiente', 'regular']):
                    score = 6.0
                elif any(w in block_lower for w in ['deficiente', 'débil', 'pobre', 'insuficiente']):
                    score = 4.0
                else:
                    score = 7.0
            else:
                score_str = score_match.group(1) or score_match.group(2)
                try:
                    score = max(0.0, min(10.0, float(score_str)))
                except Exception:
                    score = 7.0

            # Extract feedback from remaining lines
            lines = block.strip().split('\n')
            feedback_lines = [l.strip() for l in lines[1:] if l.strip()]
            feedback = ' '.join(feedback_lines)
            feedback = re.sub(r'\*\*RECOMENDACIÓN.*', '', feedback,
                              flags=re.DOTALL | re.IGNORECASE).strip()
            feedback = ' '.join(feedback.split())

            if len(feedback) < 10:
                feedback = "No disponible"

            # Truncate to 3 sentences
            sentences = [s.strip() for s in feedback.split('.') if s.strip()]
            if len(sentences) > 3:
                feedback = '. '.join(sentences[:3]) + '.'

            # Map to dimension
            block_lower = block[:200].lower()
            if 'argumentaci' in block_lower:
                result["argumentacion"] = {"score": score, "feedback": feedback}
            elif 'conclusi' in block_lower:
                result["conclusiones"] = {"score": score, "feedback": feedback}
            elif 'coherencia' in block_lower:
                result["coherencia"] = {"score": score, "feedback": feedback}
            elif 'claridad' in block_lower or 'argumento' in block_lower:
                result["claridad"] = {"score": score, "feedback": feedback}

        return result

    # ══════════════════════════════════════════════════════════════════════════
    # CALL 3: IDONEIDAD EDITORIAL
    # ══════════════════════════════════════════════════════════════════════════

    def _analyze_idoneidad(self, text_sample: str) -> Dict[str, Dict]:
        """
        Call 3: evaluates Contribución and Pertinencia as qualitative verdicts.
        Returns dict with 'contribucion' and 'pertinencia' keys.
        """
        prompt_contribucion = f"""Eres un revisor editorial académico. Analiza el siguiente fragmento de un artículo académico.

TEXTO:
{text_sample}

TAREA: Evalúa si el artículo hace una reclamación explícita de contribución al conocimiento y si sus conclusiones la respaldan. Evalúa únicamente lo que está visible en el texto:
- ¿El autor identifica explícitamente qué aporta este trabajo?
- ¿Las conclusiones contienen verbos de producción (propone, demuestra, desarrolla, valida, identifica, construye) o solo verbos de resumen (describe, presenta, repasa)?
- ¿Hay coherencia entre lo que el artículo promete en la introducción y lo que entrega en las conclusiones?

Responde EXACTAMENTE en este formato (tres líneas, sin nada más):
VEREDICTO: [SUSTENTADA / PARCIAL / NO SUSTENTADA]
CONTRIBUCION: [Una frase muy breve que nombre el aporte, por ejemplo "Innovación en tecnología de drones" o "Sin contribución observada o declarada"]
OBSERVACION: ["Sin contribución observada o declarada" / "Contribución declarada pero no suficientemente sustentada" / "Contribución sustentada — [frase muy breve del aporte]"]"""

        prompt_pertinencia = f"""Eres un revisor editorial académico. Analiza el siguiente fragmento de un artículo académico.

TEXTO:
{text_sample}

LÍNEAS DE INVESTIGACIÓN DE LA FACULTAD MILITAR CONJUNTA:
{_LINEAS_INVESTIGACION}

TAREA: Evalúa si el artículo se alinea con una o más de estas líneas de investigación.

Responde EXACTAMENTE en este formato (tres líneas, sin nada más):
VEREDICTO: [ALINEADO / PARCIALMENTE ALINEADO / NO ALINEADO]
LINEAS: [Número/s de línea/s con las que se alinea y tema específico, por ejemplo "Línea 4 — avances en tecnología criptográfica" o "Ninguna"]
JUSTIFICACION: [Frase muy breve que nombre la relación específica, por ejemplo "Avances en tecnología criptográfica" o "No se identificó alineación temática"]"""

        options = {"temperature": 0.1, "num_predict": 300}

        default_contribucion = {
            "veredicto": "No disponible",
            "contribucion": "No disponible",
            "observacion": "No disponible"
        }
        default_pertinencia = {
            "veredicto": "No disponible",
            "lineas": "No disponible",
            "justificacion": "No disponible"
        }

        try:
            r1 = self.client.generate(
                model=self.model_name, prompt=prompt_contribucion, options=options
            )
            contribucion = self._parse_idoneidad_response(r1["response"].strip(), "contribucion")
        except Exception as e:
            print(f"      ⚠️  Error en Contribución: {e}")
            contribucion = default_contribucion

        try:
            r2 = self.client.generate(
                model=self.model_name, prompt=prompt_pertinencia, options=options
            )
            pertinencia = self._parse_idoneidad_response(r2["response"].strip(), "pertinencia")
        except Exception as e:
            print(f"      ⚠️  Error en Pertinencia: {e}")
            pertinencia = default_pertinencia

        return {"contribucion": contribucion, "pertinencia": pertinencia}

    def _parse_idoneidad_response(self, raw: str, dimension: str) -> Dict[str, str]:
        """
        Parse qualitative verdict response for Contribución or Pertinencia.
        Extracts VEREDICTO and dimension-specific fields using regex.
        """
        result = {}

        v_match = re.search(r'VEREDICTO\s*:\s*(.+)', raw, re.IGNORECASE)
        result["veredicto"] = v_match.group(1).strip() if v_match else "No disponible"

        def _trim(text: str, max_chars: int = 120) -> str:
            """First sentence, truncated at word boundary."""
            sentence = text.split('.')[0].strip()
            if len(sentence) <= max_chars:
                return sentence
            trimmed = sentence[:max_chars]
            last_space = trimmed.rfind(' ')
            return trimmed[:last_space] + '…' if last_space > 0 else trimmed + '…'

        def _observacion_from_veredicto(veredicto: str, raw_obs: str) -> str:
            """Ensure observacion is consistent with veredicto."""
            v = veredicto.upper()
            if "NO SUSTENTADA" in v:
                return "Sin contribución observada o declarada."
            if "PARCIAL" in v:
                return "Contribución declarada pero no suficientemente sustentada."
            if "SUSTENTADA" in v:
                # Extract the contribution phrase from raw observacion or contribucion field
                obs = _trim(raw_obs)
                if obs and "no disponible" not in obs.lower() and "sin contribución" not in obs.lower():
                    return f"Contribución sustentada — {obs}"
                return "Contribución sustentada."
            return _trim(raw_obs)

        if dimension == "contribucion":
            c_match = re.search(r'CONTRIBUCION\s*:\s*(.+)', raw, re.IGNORECASE)
            o_match = re.search(r'OBSERVACION\s*:\s*(.+)', raw, re.IGNORECASE)
            contrib_text = _trim(c_match.group(1)) if c_match else "No disponible"
            raw_obs = _trim(o_match.group(1)) if o_match else ""
            result["contribucion"] = contrib_text
            result["observacion"] = _observacion_from_veredicto(result["veredicto"], contrib_text)
        else:
            l_match = re.search(r'LINEAS?\s*:\s*(.+)', raw, re.IGNORECASE)
            j_match = re.search(r'JUSTIFICACION\s*:\s*(.+)', raw, re.IGNORECASE)
            result["lineas"] = _trim(l_match.group(1), 80) if l_match else "No disponible"
            result["justificacion"] = _trim(j_match.group(1)) if j_match else "No disponible"
        
        
        return result


# ── Convenience function ──────────────────────────────────────────────────────

def analyze_document_quality(document: DocumentContent,
                             category: ClassificationCategory,
                             model_name: str = "llama3-gradient:8b-instruct-1048k-q4_K_M") -> QualityResult:
    analyzer = QualityAnalyzer(model_name=model_name)
    return analyzer.analyze_quality(document, document.full_text, category)
