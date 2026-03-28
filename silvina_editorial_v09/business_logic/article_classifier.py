"""
article_classifier.py
Classifies articles into categories using LLM analysis.
Part of Silvina Editorial Assistant v0.7
"""

from typing import Optional
import ollama
from domain.models import DocumentContent, ClassificationResult
from domain.enums import ArticleType
from business_logic.structure_analyzer import StructureAnalyzer

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
    
    def classify_article(self, document_content: DocumentContent) -> ClassificationResult:
        if not document_content or not document_content.paragraphs:
            raise ValueError("DocumentContent.paragraphs is empty")

        from domain.enums import classify_article_size
        article_size = classify_article_size(document_content.char_count)

        structure = StructureAnalyzer().analyze(document_content)

        # =========================
        # 1. DETERMINISTIC RULES
        # =========================

        if structure["imryd_complete"] and article_size.name != "FUERA_RANGO":
            return ClassificationResult(
                article_type=ArticleType.CIENTIFICO,
                article_size=article_size,
                confidence=0.9,
                reasoning="Estructura IMRyD completa detectada mediante análisis determinístico."
            )

        if structure["has_introduction"] and structure["has_discussion"]:
            return ClassificationResult(
                article_type=ArticleType.DIVULGACION,
                article_size=article_size,
                confidence=0.75,
                reasoning="Artículo con estructura reflexiva sin IMRyD completo."
            )

        if not structure["has_methods"] and not structure["has_results"]:
            return ClassificationResult(
                article_type=ArticleType.OPINION,
                article_size=article_size,
                confidence=0.7,
                reasoning="Texto argumentativo sin validación empírica."
            )

        # =========================
        # 2. LLM FALLBACK
        # =========================

        text_sample = ' '.join(document_content.paragraphs[:50])[:6000]

        prompt = f"""Clasifica este artículo académico según normas EUMIC.

    DOCUMENTO:
    Título: {document_content.title}
    Caracteres: {document_content.char_count:,}
    Palabras: {document_content.word_count}

    Texto:
    {text_sample}

    TIPOS:
    - CIENTIFICO: IMRyD, razonamiento crítico, citas académicas
    - DIVULGACION: Reflexión académica, sin IMRyD rígido
    - OPINION: Crítica reflexiva, sin validación empírica

    RESPONDE (una línea por campo):
    CATEGORY: [CIENTIFICO|DIVULGACION|OPINION]
    CONFIDENCE: [0.0-1.0]
    REASONING: [Máximo 2 oraciones en español]"""

        try:
            response = self.client.generate(
                model=self.model_name,
                prompt=prompt,
                options={'temperature': 0.3, 'num_predict': 200}
            )

            text = response['response'].strip()

            article_type = ArticleType.UNKNOWN
            confidence = 0.5
            reasoning = ""

            for line in text.split('\n'):
                if line.startswith('CATEGORY:'):
                    if 'CIENTIFICO' in line.upper():
                        article_type = ArticleType.CIENTIFICO
                    elif 'DIVULGACION' in line.upper():
                        article_type = ArticleType.DIVULGACION
                    elif 'OPINION' in line.upper():
                        article_type = ArticleType.OPINION
                elif line.startswith('CONFIDENCE:'):
                    try:
                        confidence = float(line.split(':', 1)[1].strip())
                    except ValueError:
                        confidence = 0.5
                elif line.startswith('REASONING:'):
                    reasoning = line.split(':', 1)[1].strip()

            return ClassificationResult(
                article_type=article_type,
                article_size=article_size,
                confidence=confidence,
                reasoning=reasoning
            )

        except Exception as e:
            print(f"⚠️  Error en clasificación LLM: {e}")

        # =========================
        # 3. ABSOLUTE FALLBACK
        # =========================

        return ClassificationResult(
            article_type=ArticleType.UNKNOWN,
            article_size=article_size,
            confidence=0.0,
            reasoning="No se pudo clasificar el artículo"
        )

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

if __name__ == "__main__":
    doc = DocumentContent(
        word_count=6000,
        char_count=35000,
        title="Test Article",
        abstract="This is a short scientific article with introduction, methods and discussion.",
        paragraphs=[
            "Este estudio analiza los efectos de X utilizando una metodología experimental.",
            "Se aplicaron métodos cuantitativos con una muestra de 120 sujetos.",
            "Los resultados muestran una correlación significativa.",
            "Se discuten las implicaciones teóricas y prácticas de los hallazgos."
        ]
    )

    classifier = ArticleClassifier()
    result = classifier.classify_article(doc)
    print(result)

