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
        
    def analyze_quality(self, document_content, article_type) -> QualityAnalysisResult:
        print("      ⏳ Analizando calidad...")
    
        # Strategic sampling
        parts = []
        parts.append(document_content.title or "")
        parts.extend([p.text for p in document_content.paragraphs[:2]])
        mid = len(document_content.paragraphs) // 2
        parts.extend([p.text for p in document_content.paragraphs[mid:mid+2]])
        parts.extend([p.text for p in document_content.paragraphs[-2:]])
        text_sample = ' '.join(parts)[:3000]
        
        prompt = f"""Analiza calidad editorial de este artículo.

    TEXTO:
    {text_sample}

    EVALÚA (0-10) en JSON:
    {{
    "claridad": {{"score": X, "feedback": "Una oración justificando"}},
    "coherencia": {{"score": X, "feedback": "Una oración justificando"}},
    "argumentacion": {{"score": X, "feedback": "Una oración justificando"}},
    "conclusiones": {{"score": X, "feedback": "Una oración justificando"}},
    "formato": {{"score": X, "feedback": "Una oración justificando"}}
    }}

    INSTRUCCIONES: Conciso, académico, enfoque editorial EUMIC/APA."""

        try:
            response = self.ollama.generate(
                model=self.model_name,
                prompt=prompt,
                options={'temperature': 0.3, 'num_predict': 600, 'num_ctx': 4096},
                timeout=90
            )
            
            import json, re
            text = response['response'].strip()
            json_match = re.search(r'\{.*\}', text, re.DOTALL)
            
            if json_match:
                scores = json.loads(json_match.group())
            else:
                scores = {
                    "claridad": {"score": 7.0, "feedback": "Análisis no disponible"},
                    "coherencia": {"score": 7.0, "feedback": "Análisis no disponible"},
                    "argumentacion": {"score": 7.0, "feedback": "Análisis no disponible"},
                    "conclusiones": {"score": 7.0, "feedback": "Análisis no disponible"},
                    "formato": {"score": 7.0, "feedback": "Análisis no disponible"}
                }
            
            overall = sum(s["score"] for s in scores.values()) / len(scores)
            quality_level = get_quality_level_from_score(overall)
            
            return QualityAnalysisResult(
                overall_score=overall,
                quality_level=quality_level,
                dimension_scores=scores
            )
            
        except Exception as e:
            print(f"⚠️  Error: {e}")
            default = {d: {"score": 7.0, "feedback": "Error en análisis"} 
                    for d in ["claridad", "coherencia", "argumentacion", "conclusiones", "formato"]}
            return QualityAnalysisResult(7.0, QualityLevel.ACCEPTABLE, default)


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