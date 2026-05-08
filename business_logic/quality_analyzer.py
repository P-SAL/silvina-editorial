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


class QualityAnalyzer:
    """Analyzes academic document quality across semantic dimensions."""

    def __init__(self, model_name: str = "llama3-gradient:8b-instruct-1048k-q4_K_M",
                 base_url: str = "http://localhost:11434"):
        self.model_name = model_name
        import ollama
        self.ollama = ollama
        self.base_url = base_url
        self.client = ollama.Client(host=self.base_url)

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
            'num_predict': 800,
            'num_ctx': 4096,
            'repeat_penalty': 1.1,
            'timeout': 120
        }

        # ── CALL 1: Claridad + Coherencia ────────────────────────────────
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

        # ── CALL 2: Argumentación + Conclusiones ─────────────────────────
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
            # First call
            response_1 = self.ollama.generate(
                model=self.model_name,
                prompt=prompt_1,
                options=ollama_options
            )
            text_1 = response_1.get('response', '').strip()
            print(f"      ✓ Llamada 1 completada: {len(text_1.split())} palabras")

            # Second call
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

            print(f"      ✓ Análisis generado: {overall:.1f}/10\n")

            return QualityAnalysisResult(
                overall_score=overall,
                quality_level=quality_level,
                dimension_scores=scores
            )

        except Exception as e:
            print(f"      ⚠️  Error en LLM: {e}")
            default = {
                d: {"score": 7.0, "feedback": "Análisis no disponible"}
                for d in ["claridad", "coherencia", "argumentacion", "conclusiones"]
            }
            return QualityAnalysisResult(7.0, QualityLevel.ACCEPTABLE, default)

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

            first_line = block.split('\n')[0]

            # Search entire block for score
            score_match = re.search(
                r'\[Puntuaci[oó]n:\s*(\d+(?:\.\d+)?)(?:/10)?\]|(\d+(?:\.\d+)?)\s*/\s*10',
                block, re.IGNORECASE
            )
            if not score_match:
                continue
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

            # Map to dimension — search first 200 chars of block
            # Order matters: check argumentaci before claridad to avoid false match on "argumento"
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


# Convenience function
def analyze_document_quality(document: DocumentContent,
                             category: ClassificationCategory,
                             model_name: str = "llama3-gradient:8b-instruct-1048k-q4_K_M") -> QualityResult:
    analyzer = QualityAnalyzer(model_name=model_name)
    return analyzer.analyze_quality(document, document.full_text, category)
