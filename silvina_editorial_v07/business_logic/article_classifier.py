"""
article_classifier.py
Classifies articles into categories using LLM analysis.
Part of Silvina Editorial Assistant v0.7
"""

from typing import Optional
import ollama
from domain.models import DocumentContent, ClassificationResult
from domain.enums import ClassificationCategory


class ArticleClassifier:
    """Classifies academic articles using LLM analysis."""
    
    def __init__(self, model_name: str = "llama3-gradient:8b-instruct-1048k-q4_K_M", 
                 base_url: str = "http://localhost:11434"):
        """
        Initialize the article classifier.
        
        Args:
            model_name: Name of the Ollama model to use
            base_url: Base URL for Ollama API
        """
        self.model_name = model_name
        self.base_url = base_url
        self.client = ollama.Client(host=base_url)
    
    def classify_article(self, document: DocumentContent) -> ClassificationResult:
        """
        Classify an article into one of the defined categories.
        
        Args:
            document: DocumentContent object with article text
            
        Returns:
            ClassificationResult with category, confidence, and reasoning
        """
        # Prepare the prompt for classification
        prompt = self._create_classification_prompt(document)
        
        try:
            # Call LLM for classification
            response = self.client.chat(
                model=self.model_name,
                messages=[
                    {
                        'role': 'system',
                        'content': self._get_system_prompt()
                    },
                    {
                        'role': 'user',
                        'content': prompt
                    }
                ],
                options={
                    'temperature': 0.3,
                    'num_predict': 500
                }
            )
            
            # Parse the response
            response_text = response['message']['content']
            # Debug: Print raw response
            category, confidence, reasoning = self._parse_classification_response(response_text)
            
            return ClassificationResult(
                category=category,
                confidence=confidence,
                reasoning=reasoning
            )
            
        except Exception as e:
            print(f"Warning: LLM classification failed: {e}")
            # Return fallback classification
            return ClassificationResult(
                category=ClassificationCategory.UNKNOWN,
                confidence=0.0,
                reasoning=f"Classification failed due to error: {str(e)}"
            )
    
    def _get_system_prompt(self) -> str:
        """Get the system prompt for classification."""
        return """Eres un experto en clasificación de artículos académicos según las normas EUMIC.

Debes clasificar artículos en una de estas categorías:

1. RESEARCH_ARTICLE: Investigación original con metodología, resultados y análisis
2. REVIEW_ARTICLE: Revisión sistemática de literatura sobre un tema
3. REFLECTION_ARTICLE: Análisis crítico y reflexivo desde perspectiva del autor
4. SHORT_ARTICLE: Comunicación breve de resultados preliminares
5. CASE_REPORT: Descripción detallada de un caso específico

IMPORTANTE: Debes responder ÚNICAMENTE con estas tres líneas, nada más:

CATEGORY: [escribe exactamente uno: RESEARCH_ARTICLE, REVIEW_ARTICLE, REFLECTION_ARTICLE, SHORT_ARTICLE, o CASE_REPORT]
CONFIDENCE: [un número decimal entre 0.0 y 1.0, ejemplo: 0.85]
REASONING: [una línea breve explicando por qué]

Ejemplo de respuesta válida:
CATEGORY: RESEARCH_ARTICLE
CONFIDENCE: 0.9
REASONING: El artículo presenta metodología clara, resultados experimentales y análisis de datos originales.

NO agregues texto adicional antes o después de estas tres líneas."""
    
    def _create_classification_prompt(self, document: DocumentContent) -> str:
        """Create the classification prompt from document content."""
        # Build document summary for classification
        doc_summary = []
        
        if document.title:
            doc_summary.append(f"TÍTULO: {document.title}")
        
        if document.abstract:
            abstract_preview = document.abstract[:500]
            doc_summary.append(f"RESUMEN: {abstract_preview}")
        
        if document.sections:
            doc_summary.append(f"SECCIONES IDENTIFICADAS: {', '.join(document.sections.keys())}")
        
        doc_summary.append(f"PALABRAS TOTALES: {document.word_count}")
        
        # Add content preview
        content_preview = ' '.join(document.paragraphs[:5])[:1000]
        doc_summary.append(f"CONTENIDO (primeros párrafos): {content_preview}")
        
        prompt = f"""Clasifica el siguiente artículo académico:

{chr(10).join(doc_summary)}

Analiza la estructura, contenido y características del artículo para determinar su categoría."""
        
        return prompt
    
    def _parse_classification_response(self, response_text: str) -> tuple:
        """
        Parse the LLM response to extract category, confidence, and reasoning.
        
        Args:
            response_text: Raw text response from LLM
            
        Returns:
            Tuple of (category, confidence, reasoning)
        """
        import re
        
        # Clean response text
        response_text = response_text.strip()
        
        # Extract category (try multiple patterns)
        category = ClassificationCategory.UNKNOWN
        
        # Pattern 1: Exact format "CATEGORY: RESEARCH_ARTICLE"
        category_match = re.search(r'CATEGORY:\s*([A-Z_]+)', response_text, re.IGNORECASE)
        if category_match:
            category_str = category_match.group(1).upper()
            # Try direct match
            if category_str in ClassificationCategory.__members__:
                category = ClassificationCategory[category_str]
        
        # Pattern 2: If not found, look for category names anywhere in first 200 chars
        if category == ClassificationCategory.UNKNOWN:
            first_part = response_text[:200].upper()
            for cat_name in ClassificationCategory.__members__:
                if cat_name in first_part:
                    category = ClassificationCategory[cat_name]
                    break
        
        # Extract confidence
        confidence = 0.5  # Default
        confidence_match = re.search(r'CONFIDENCE:\s*([\d.]+)', response_text, re.IGNORECASE)
        if confidence_match:
            try:
                confidence = float(confidence_match.group(1))
                confidence = max(0.0, min(1.0, confidence))  # Clamp to [0, 1]
            except ValueError:
                pass
        
        # Extract reasoning
        reasoning = "Clasificación basada en análisis del contenido."
        reasoning_match = re.search(r'REASONING:\s*(.+?)(?:\n\n|\n[A-Z]+:|\Z)', 
                                   response_text, re.IGNORECASE | re.DOTALL)
        if reasoning_match:
            reasoning = reasoning_match.group(1).strip()
            # Limit reasoning length
            if len(reasoning) > 300:
                reasoning = reasoning[:297] + "..."
        
        return category, confidence, reasoning


# Convenience function
def classify_document(document: DocumentContent, 
                     model_name: str = "llama3-gradient:8b-instruct-1048k-q4_K_M") -> ClassificationResult:
    """
    Classify a document using the default classifier.
    
    Args:
        document: DocumentContent to classify
        model_name: Ollama model to use
        
    Returns:
        ClassificationResult
    """
    classifier = ArticleClassifier(model_name=model_name)
    return classifier.classify_article(document)