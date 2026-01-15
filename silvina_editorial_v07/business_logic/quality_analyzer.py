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
        self.base_url = base_url
        self.client = ollama.Client(host=self.base_url)
        
         
    def analyze_quality(self, document, document_text, category) -> 'QualityResult':
        """Analyze document quality across all dimensions efficiently."""
    
        print("      ⏳ Analizando calidad (modo rápido)...")
        
        # Create ONE combined prompt instead of 6 separate calls
        combined_prompt = f"""Analiza este artículo académico en TODAS estas dimensiones y responde en formato JSON:

DOCUMENTO:
Título: {document.title if hasattr(document, "title") else document.paragraphs[0].text}

Palabras: {len(document_text.split())}
Categoría: {category.value}

Texto (primeros 1500 palabras):
{' '.join([p.text for p in document.paragraphs[:50]])[:6000]}

EVALÚA (escala 0-10) y responde SOLO en JSON sin texto adicional:
{{
  "claridad": {{"score": X, "feedback": "..."}},
  "coherencia": {{"score": X, "feedback": "..."}},
  "argumentacion": {{"score": X, "feedback": "..."}},
  "metodologia": {{"score": X, "feedback": "..."}},
  "conclusiones": {{"score": X, "feedback": "..."}},
  "formato": {{"score": X, "feedback": "..."}}
}}"""

        try:
            response = self.client.generate(
                model=self.model_name,
                prompt=combined_prompt,
                options={'temperature': 0.3, 'num_predict': 800}
            )
            
            # Parse JSON response
            import json
            import re
            
            response_text = response['response'].strip()
            # Extract JSON from response
            json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
            if json_match:
                scores_data = json.loads(json_match.group())
            else:
                # Fallback to default scores if parsing fails
                scores_data = {dim: {"score": 7.0, "feedback": "Análisis no disponible"} 
                            for dim in ["claridad", "coherencia", "argumentacion", "metodologia", "conclusiones", "formato"]}
            
            # Calculate overall score
            overall = sum(scores_data[dim]["score"] for dim in scores_data) / len(scores_data)
            
            # Determine quality level
            quality_level = get_quality_level_from_score(overall)
           
            return QualityResult(
                overall_score=overall,
                quality_level=quality_level,
                dimension_scores=scores_data
            )
            
        except Exception as e:
            print(f"      ⚠ Error en análisis de calidad: {e}, usando valores por defecto")
            # Return default acceptable scores
            default_scores = {dim: {"score": 7.0, "feedback": "Análisis automático no completado"} 
                            for dim in ["claridad", "coherencia", "argumentacion", "metodologia", "conclusiones", "formato"]}
            return QualityResult(
                overall_score=7.0,
                quality_level=QualityLevel.ACCEPTABLE,
                dimension_scores=default_scores
            )
    
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