"""
quality_analyzer.py
Analyzes document quality across multiple dimensions using LLM.
Part of Silvina Editorial Assistant v0.7
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
        
        # Sample text - INCREASED from 3500 to 8000 for better context
        parts = []
        parts.append(document_content.title or "")
        parts.extend(document_content.paragraphs[:3])
        mid = len(document_content.paragraphs) // 2
        parts.extend(document_content.paragraphs[mid:mid+2])
        parts.extend(document_content.paragraphs[-2:])
        text_sample = ' '.join(parts)[:8000]  # INCREASED
        
        prompt = f"""Analiza la calidad editorial de este artículo académico. Para cada dimensión, proporciona un análisis breve (máximo 70 palabras) y específico.

TEXTO:
{text_sample}

FORMATO DE RESPUESTA (5 dimensiones):

**1. Normativa**: (Corrección ortográfica y gramatical)
[Análisis en 70 palabras máximo. Si hay errores, dar máximo 3 ejemplos con: Página, Error, Corrección, Tipo]

**2. Claridad del argumento**:
[Análisis en 70 palabras máximo]

**3. Coherencia**:
[Análisis en 70 palabras máximo]

**4. Argumentación**:
[Análisis en 70 palabras máximo]

**5. Conclusiones**:
[Análisis en 70 palabras máximo]

INSTRUCCIONES:
- Máximo 70 palabras por dimensión
- En Normativa: máximo 3 ejemplos de errores
- Conciso, específico, tono editorial académico
- AL FINAL: Añadir "**RECOMENDACIÓN**: [APTO PARA PUBLICACIÓN / NO SE RECOMIENDA PUBLICAR] porque [razón breve]"
"""

        try:
            response = self.ollama.generate(
                model=self.model_name,
                prompt=prompt,
                options={'temperature': 0.3, 'num_predict': 1000, 'num_ctx': 4096}
            )
            
            analysis_text = response.get('response', '').strip()
            
            # Remove unwanted "Nota:" sections
            analysis_text = re.sub(r'\n*Nota:.*', '', analysis_text, flags=re.DOTALL | re.IGNORECASE)
            analysis_text = analysis_text.strip()
            
            word_count = len(analysis_text.split())
            print(f"Generando análisis: {word_count} palabras")
            print("✅ Análisis completado\n")
            print(analysis_text)
            
                        
            # Parse the LLM response into structured scores
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
        """Extract feedback from LLM response - IMPROVED VERSION."""
        
        result = {
            "normativa": {"score": 7.0, "feedback": "No disponible"},
            "claridad": {"score": 7.0, "feedback": "No disponible"},
            "coherencia": {"score": 7.0, "feedback": "No disponible"},
            "argumentacion": {"score": 7.0, "feedback": "No disponible"},
            "conclusiones": {"score": 7.0, "feedback": "No disponible"}
        }
        
        # IMPROVED: More flexible pattern that handles variations
        # Matches: **1. Name**: or **1. Name**:
        # Captures everything until next **number. or end of text
        pattern = r'\*\*\d+\.\s*([^:*\n]+)\**:\**\s*(.*?)(?=\n\*\*\d+\.|$)'
        matches = re.findall(pattern, text, re.DOTALL)
        
        for name, content in matches:
            name_lower = name.strip().lower()
            
            # Remove **RECOMENDACIÓN** sections
            clean_content = re.sub(r'\*\*RECOMENDACIÓN\*\*:.*', '', content, flags=re.DOTALL)
            clean_content = clean_content.strip()
            
            # Remove subtitle in parentheses if present
            clean_content = re.sub(r'^\([^)]+\)\s*', '', clean_content)
            
            # Limit length
            clean_content = clean_content[:500]
            
            # Only store if we actually got content
            if len(clean_content) > 10:  # Minimum 10 chars to be valid
                if "normativ" in name_lower:
                    result["normativa"]["feedback"] = clean_content
                elif "claridad" in name_lower:
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
