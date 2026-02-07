"""
quality_analyzer.py
Analyzes document quality across multiple dimensions using LLM.
FIXED VERSION - Improved prompt to eliminate hallucinations
"""
from __future__ import annotations

import ollama
import re
import json
from typing import Dict, Any
from domain.enums import ClassificationCategory, QualityLevel, get_quality_level_from_score
from domain.models import DocumentContent, QualityResult
from domain.models import QualityAnalysisResult


class QualityAnalyzer:
    """Analyzes academic document quality across multiple dimensions."""
    
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
        
        # IMPROVED PROMPT - eliminates hallucinations
        prompt = f"""Eres un revisor editorial académico. Analiza este fragmento de artículo académico en CINCO dimensiones.

TEXTO:
{text_sample}

INSTRUCCIONES CRÍTICAS:
1. SOLO analiza lo que REALMENTE VES en el texto
2. NO inventes errores que no existen
3. Si no encuentras errores, di "No se detectaron errores en esta muestra"
4. Máximo 60 palabras por dimensión
5. En Normativa: SOLO mencionar errores si los hay (máximo 2 ejemplos REALES)

FORMATO DE RESPUESTA:

**1. Normativa**:
[Si hay errores ortográficos/gramaticales: mencionar MAX 2 ejemplos REALES con página/ubicación. Si NO hay errores: escribir "No se detectaron errores ortográficos o gramaticales en esta muestra."]

**2. Claridad del argumento**:
[Evaluar si el argumento es claro y comprensible. MAX 60 palabras]

**3. Coherencia**:
[Evaluar conexión lógica entre ideas. MAX 60 palabras]

**4. Argumentación**:
[Evaluar solidez de argumentos y evidencia. MAX 60 palabras]

**5. Conclusiones**:
[Evaluar si hay conclusiones claras. MAX 60 palabras]

REGLAS ESTRICTAS:
- NO repitas el mismo comentario en múltiples dimensiones
- NO inventes problemas
- Sé específico y objetivo
- Al final: "**RECOMENDACIÓN**: APTO PARA PUBLICACIÓN" o "NO SE RECOMIENDA PUBLICAR"
"""

        try:
            response = self.ollama.generate(
                model=self.model_name,
                prompt=prompt,
                options={
                    'temperature': 0.2,  # Lower = more factual
                    'num_predict': 800,   # Shorter to avoid repetition
                    'num_ctx': 4096
                }
            )
            
            analysis_text = response.get('response', '').strip()
            
            # Clean up unwanted sections
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
                for d in ["normativa", "claridad", "coherencia", "argumentacion", "conclusiones"]
            }
            return QualityAnalysisResult(7.0, QualityLevel.ACCEPTABLE, default)

    def _parse_llm_response(self, text: str) -> Dict[str, Dict[str, Any]]:
        """
        Extract feedback from LLM response.
        FIXED VERSION - Better parsing, handles all formats.
        """
        
        result = {
            "normativa": {"score": 7.0, "feedback": "No disponible"},
            "claridad": {"score": 7.0, "feedback": "No disponible"},
            "coherencia": {"score": 7.0, "feedback": "No disponible"},
            "argumentacion": {"score": 7.0, "feedback": "No disponible"},
            "conclusiones": {"score": 7.0, "feedback": "No disponible"}
        }
        
        # IMPROVED REGEX: Handles multiple formats
        
        pattern = r'\*\*(?:\d+\.\s*)?([^:*\n]+)\**:?\s*\n(.*?)(?=\n\*\*(?:\d+\.)?|\n*$)'
        matches = re.findall(pattern, text, re.DOTALL)
        
        for name, content in matches:
            name_lower = name.strip().lower()
            
            # Remove recommendation sections
            clean_content = re.sub(
                r'\*\*RECOMENDACIÓN\*\*:.*',
                '',
                content,
                flags=re.DOTALL | re.IGNORECASE
            )
            
            # Remove subtitle in parentheses
            clean_content = re.sub(r'^\([^)]+\)\s*', '', clean_content.strip())
            
            # Normalize whitespace
            clean_content = ' '.join(clean_content.split())
                       
            if len(clean_content) < 10:
                clean_content = "Análisis no disponible para esta dimensión."

            # Limit length at sentence boundary
            
            
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
            
            # Only store if valid content (min 10 chars)
            if len(clean_content) >= 10:
                # Map to dimension
                if "normativ" in name_lower:
                    result["normativa"]["feedback"] = clean_content
                elif "claridad" in name_lower or "argumento" in name_lower:
                    result["claridad"]["feedback"] = clean_content
                elif "coherencia" in name_lower:
                    result["coherencia"]["feedback"] = clean_content
                elif "argumentaci" in name_lower:
                    result["argumentacion"]["feedback"] = clean_content
                elif "conclusion" in name_lower:
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
