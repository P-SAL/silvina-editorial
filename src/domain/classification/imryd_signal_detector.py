from src.domain.dtos.document_content_dto import DocumentContentDTO


class ImrydSignalDetector:
    """Domain service that detects IMRyD section-presence signals in a document.

    Scans only short paragraphs (<= 5 words) as section-header candidates, to avoid
    false positives from body prose containing section-name words (e.g. "resultados"
    appearing mid-sentence rather than as a heading).
    """

    _HEADER_CANDIDATE_MAX_WORD_COUNT = 5
    _IMRYD_KEYWORDS: dict[str, tuple[str, ...]] = {
        "introduction": (
            "introduction",
            "background",
            "context",
            "introducción",
            "introduccion",
            "intro",
        ),
        "methods": (
            "method",
            "methodology",
            "materials",
            "procedures",
            "método",
            "metodo",
            "metodología",
            "metodologia",
            "métodos",
            "metodos",
            "materiales",
        ),
        "results": ("results", "findings", "resultados", "hallazgos"),
        "discussion": ("discussion", "discusión", "discusion"),
        "conclusion": (
            "conclusion",
            "conclusions",
            "concluding",
            "conclusión",
            "conclusiones",
        ),
    }

    def detect(self, document_content: DocumentContentDTO) -> dict[str, bool]:
        """Return IMRyD section-presence signals for the given document."""
        header_candidates = [
            paragraph.strip().lower()
            for paragraph in document_content.paragraphs
            if 1 <= len(paragraph.strip().split()) <= self._HEADER_CANDIDATE_MAX_WORD_COUNT
        ]

        signals = {
            "has_introduction": False,
            "has_methods": False,
            "has_results": False,
            "has_discussion": False,
            "has_conclusion": False,
            "imryd_complete": False,
        }

        for section, keywords in self._IMRYD_KEYWORDS.items():
            if any(keyword in header for keyword in keywords for header in header_candidates):
                signals[f"has_{section}"] = True

        signals["imryd_complete"] = (
            signals["has_introduction"]
            and signals["has_methods"]
            and signals["has_results"]
            and signals["has_discussion"]
        )

        return signals
