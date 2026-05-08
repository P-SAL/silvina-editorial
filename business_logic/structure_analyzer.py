"""
structure_analyzer.py
Analyzes structural features of academic documents (IMRyD, signals).
Part of Silvina Editorial Assistant v0.9
"""

from typing import Dict
from domain.models import DocumentContent


class StructureAnalyzer:
    """
    Detects structural patterns in academic documents
    (e.g., IMRyD, opinion, divulgation).

    IMRyD detection scans only short paragraphs (≤ 5 words) as section
    headers to avoid false positives from body prose.
    """

    IMRYD_KEYWORDS = {
        "introduction": [
            "introduction", "background", "context",
            "introducción", "introduccion", "intro"
        ],
        "methods": [
            "method", "methodology", "materials", "procedures",
            "método", "metodo", "metodología", "metodologia",
            "métodos", "metodos", "materiales"
        ],
        "results": [
            "results", "findings",
            "resultados", "hallazgos"
        ],
        "discussion": [
            "discussion",
            "discusión", "discusion"
        ],
        "conclusion": [
            "conclusion", "conclusions", "concluding",
            "conclusión", "conclusiones"
        ]
    }

    def analyze(self, document: DocumentContent) -> Dict[str, bool]:
        """
        Analyze document structure and return detected signals.
        Only short paragraphs (≤ 5 words) are considered section headers.
        """

        # Only scan short paragraphs as section headers — avoids false positives
        # from body prose containing words like "análisis" or "resultados"
        header_candidates = [
            p.strip().lower() for p in document.paragraphs
            if 1 <= len(p.strip().split()) <= 5
        ]

        signals = {
            "has_introduction": False,
            "has_methods": False,
            "has_results": False,
            "has_discussion": False,
            "has_conclusion": False,
            "imryd_complete": False
        }

        for section, keywords in self.IMRYD_KEYWORDS.items():
            for kw in keywords:
                if any(kw in header for header in header_candidates):
                    signals[f"has_{section}"] = True
                    break

        # IMRyD completeness requires all 4 core sections
        signals["imryd_complete"] = (
            signals["has_introduction"]
            and signals["has_methods"]
            and signals["has_results"]
            and signals["has_discussion"]
        )

        return signals


def analyze_structure(document: DocumentContent) -> Dict[str, bool]:
    """Convenience function."""
    return StructureAnalyzer().analyze(document)


if __name__ == "__main__":
    doc = DocumentContent(
        title="Test",
        word_count=5000,
        char_count=30000,
        paragraphs=[
            "Introducción",
            "Este estudio analiza los efectos de la integración conjunta.",
            "Metodología",
            "Se aplicó una encuesta a una muestra de 240 oficiales.",
            "Resultados",
            "Los hallazgos muestran correlación significativa.",
            "Discusión",
            "Estos resultados sugieren implicaciones institucionales.",
        ]
    )

    analyzer = StructureAnalyzer()
    result = analyzer.analyze(doc)
    print(result)
    print(f"IMRyD completo: {result['imryd_complete']}")
