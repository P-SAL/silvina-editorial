"""
structure_analyzer.py
Analyzes structural features of academic documents (IMRyD, signals).
Part of Silvina Editorial Assistant v0.7
"""

from typing import Dict
from domain.models import DocumentContent


class StructureAnalyzer:
    """
    Detects structural patterns in academic documents
    (e.g., IMRyD, opinion, divulgation).
    """

    IMRYD_KEYWORDS = {
    "introduction": [
        "introduction", "background", "context",
        "introducción", "introduccion", "este estudio", "el presente trabajo"
    ],
    "methods": [
        "method", "methodology", "materials", "procedures",
        "método", "metodo", "metodología", "metodologia",
        "se aplicaron", "se utilizaron"
    ],
    "results": [
        "results", "findings", "data",
        "resultados", "los resultados", "hallazgos"
    ],
    "discussion": [
        "discussion", "analysis", "interpretation",
        "discusión", "discusion", "se discuten", "análisis"
    ],
    "conclusion": [
        "conclusion", "concluding",
        "conclusión", "conclusion", "conclusiones"
    ]
}

    def analyze(self, document: DocumentContent) -> Dict[str, bool]:
        """
        Analyze document structure and return detected signals.
        """

        text = " ".join(document.paragraphs).lower()

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
                if kw in text:
                    signals[f"has_{section}"] = True
                    break

        # IMRyD completeness (core 4)
        signals["imryd_complete"] = (
            signals["has_introduction"]
            and signals["has_methods"]
            and signals["has_results"]
            and signals["has_discussion"]
        )

        return signals

    def analyze_structure(document: DocumentContent) -> Dict[str, bool]:
        analyzer = StructureAnalyzer()
        return analyzer.analyze(document)

if __name__ == "__main__":
    doc = DocumentContent(
        title="Test",
        word_count=5000,
        char_count=30000,
        paragraphs=[
            "Introduction This study analyzes...",
            "Methods We applied quantitative methods...",
            "Results The findings show significant correlation...",
            "Discussion These results suggest..."
        ]
    )

    analyzer = StructureAnalyzer()
    print(analyzer.analyze(doc))
