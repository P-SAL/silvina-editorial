"""
quality_analyzer.py
Analyzes document quality across multiple SEMANTIC dimensions using LLM.
TIER 2 - Focuses on content quality, not grammar/spelling (that's Tier 1)
Part of Silvina Editorial Assistant v0.7
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
        """
        Initialize the quality analyzer.
        
        Args:
            model_name: Name of the Ollama model to use
            base_url: Base URL for Ollama API
        """
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
        parts.extend(document_content.paragraphs[-2:])  # Conclusion
        text_sample = ' '.join(parts)[:8000]
        
        # IMPROVED PROMPT - Better instructions, clearer format
        prompt = f"""Eres un revisor editorial académico experto. Analiza este fragmento de artículo académico en CUATRO dimensiones semánticas.

TEXTO A ANALIZAR:
{text_sample}

INSTRUCCIONES IMPORTANTES:
1. Evalúa SOLO lo que realmente está presente en el texto
2. Si ves argumentos, reconócelos (no digas "no presenta argumentos" si los hay)
3. Sé específico: menciona qué funciona bien y qué necesita mejorar
4. No repitas información - cada dimensión debe aportar algo nuevo
5. La ortografía y gramática ya fueron verificadas - enfócate en el CONTENIDO

FORMATO DE RESPUESTA (OBLIGATORIO):

**1. Claridad del argumento** [Puntuación: X/10]
[Analiza si el argumento central es claro. ¿El lector entiende fácilmente el mensaje principal?]

**2. Coherencia** [Puntuación: X/10]
[Analiza si las ideas se conectan lógicamente. ¿Hay transiciones claras entre secciones?]

**3. Argumentación** [Puntuación: X/10]
[Analiza la calidad de los argumentos. ¿Están respaldados con evidencia, ejemplos o citas?]

**4. Conclusiones** [Puntuación: X/10]
[Analiza si las conclusiones son claras y se derivan del contenido presentado.]

CRITERIOS DE PUNTUACIÓN:
- 9-10: Excelente calidad, listo para publicación
- 7-8: Buena calidad, mejoras menores recomendadas
- 5-6: Calidad aceptable, requiere trabajo adicional
- 3-4: Calidad deficiente, necesita revisión profunda
- 0-2: Calidad inaceptable

RECOMENDACIÓN FINAL (elige UNA opción):
- Si la puntuación promedio es ≥7: "APTO PARA PUBLICACIÓN"
- Si la puntuación promedio es <7: "NO SE RECOMIENDA PUBLICAR"

NO agregues notas adicionales después de la recomendación.
"""

        try:
            response = self.ollama.generate(
                model=self.model_name,
                prompt=prompt,
                options={
                    'temperature': 0.2,  # Lower = more factual
                    'num_predict': 800,   # Shorter to avoid repetition
                    'num_ctx': 4096,
                    'timeout': 120}  # 120 seconds timeout
)
            
            
            
            analysis_text = response.get('response', '').strip()

            # Clean up unwanted sections
            # 1. Remove everything after RECOMENDACIÓN FINAL
            analysis_text = re.sub(
                r'(\*\*RECOMENDACIÓN FINAL\*\*:.*?)(?:\n\nNota:.*)?$',
                r'\1',
                analysis_text,
                flags=re.DOTALL | re.IGNORECASE
            )
            # 2. Remove any remaining standalone "Nota:" sections
            analysis_text = re.sub(r'\n*Nota:.*', '', analysis_text, flags=re.DOTALL | re.IGNORECASE)
            analysis_text = analysis_text.strip()

            
            word_count = len(analysis_text.split())
            print(f"      ✓ Análisis generado: {word_count} palabras\n")
            
            # Parse into structured scores
            scores = self._parse_llm_response(analysis_text)
            overall = sum(d["score"] for d in scores.values()) / len(scores)
            quality_level = get_quality_level_from_score(overall)
           
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
        Extract feedback AND scores from LLM response.
        Handles BOTH formats:
        - Format 1: **Dimension** [Puntuación: 8/10]
        - Format 2: **Dimension**: 8/10
        """
        
        result = {
            "claridad": {"score": 7.0, "feedback": "No disponible"},
            "coherencia": {"score": 7.0, "feedback": "No disponible"},
            "argumentacion": {"score": 7.0, "feedback": "No disponible"},
            "conclusiones": {"score": 7.0, "feedback": "No disponible"}
        }
        
        # Pattern handles both [Puntuación: X/10] and : X/10 formats
        pattern = r'\*\*(?:\d+\.\s*)?([^*:\[]+?)(?:\**)?\s*(?:\[Puntuación:\s*(\d+(?:\.\d+)?)(?:/10)?\]|:\s*(\d+(?:\.\d+)?)(?:/10)?)\s*[:\n]+(.*?)(?=\n\*\*|\n*$)'
        matches = re.findall(pattern, text, re.DOTALL | re.IGNORECASE)
        
        for match in matches:
            name = match[0].strip()
            # Score can be in position 1 or 2 depending on format
            score_str = match[1] if match[1] else match[2]
            content = match[3]
            
            name_lower = name.strip().lower()
            
            # Parse score
            try:
                score = float(score_str)
                score = max(0.0, min(10.0, score))
            except (ValueError, TypeError):
                score = 7.0
            
            # Clean feedback content - remove everything after RECOMENDACIÓN
            clean_content = re.sub(
                r'\*\*RECOMENDACIÓN.*',
                '',
                content,
                flags=re.DOTALL | re.IGNORECASE
            )
            clean_content = re.sub(r'^\([^)]+\)\s*', '', clean_content.strip())
            clean_content = ' '.join(clean_content.split())
            
            # Ensure minimum feedback length
            if len(clean_content) < 10:
                clean_content = "Análisis no disponible para esta dimensión."
            
            # Limit length at sentence boundary
            if len(clean_content) > 500:
                sentences = clean_content[:500].split('.')
                if len(sentences) > 1:
                    clean_content = '.'.join(sentences[:-1]) + '.'
                else:
                    clean_content = clean_content[:500] + '...'
            
            # Map to dimension
            if "claridad" in name_lower or ("argumento" in name_lower and "argumentaci" not in name_lower):
                result["claridad"]["score"] = score
                result["claridad"]["feedback"] = clean_content
            elif "coherencia" in name_lower:
                result["coherencia"]["score"] = score
                result["coherencia"]["feedback"] = clean_content
            elif "argumentaci" in name_lower:
                result["argumentacion"]["score"] = score
                result["argumentacion"]["feedback"] = clean_content
            elif "conclusion" in name_lower:
                result["conclusiones"]["score"] = score
                result["conclusiones"]["feedback"] = clean_content
        
        return result


# Convenience function
def analyze_document_quality(document: DocumentContent,
                            category: ClassificationCategory,
                            model_name: str = "llama3-gradient:8b-instruct-1048k-q4_K_M") -> QualityResult:
    """
    Analyze document quality using default analyzer.
    
    Args:
        document: DocumentContent to analyze
        category: Article classification
        model_name: Ollama model to use
        
    Returns:
        QualityResult
    """
    analyzer = QualityAnalyzer(model_name=model_name)
    return analyzer.analyze_quality(document, document.full_text, category)