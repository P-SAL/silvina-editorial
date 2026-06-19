from src.domain.dtos.document_content_dto import DocumentContentDTO
from src.domain.enums.article_type import ArticleType
from src.domain.enums.section_name import SectionName
from src.domain.structure.required_sections_provider import RequiredSectionsProvider


class StructureValidator:
    """Domain service that validates the structural completeness of a document."""

    _SECTION_ALIASES: dict[SectionName, list[str]] = {
        SectionName.SUMMARY: ["resumen", "abstract"],
        SectionName.INTRODUCTION: ["introducción", "introduccion", "introduction"],
        SectionName.METHODOLOGY: ["metodología", "metodologia", "methodology"],
        SectionName.RESULTS: ["resultados", "results"],
        SectionName.DISCUSSION: ["discusión", "discusion", "discussion"],
        SectionName.ARGUMENTATION: ["argumentación", "argumentacion", "argumentation"],
        SectionName.DEVELOPMENT: ["desarrollo", "development"],
        SectionName.CONCLUSIONS: ["conclusiones", "conclusión", "conclusion"],
        SectionName.REFERENCES: [
            "referencias",
            "bibliografía",
            "bibliografia",
            "fuentes bibliográficas",
        ],
    }

    def __init__(self) -> None:
        pass

    def validate(
        self,
        document_content: DocumentContentDTO,
        article_type: ArticleType,
    ) -> tuple[list[SectionName], list[SectionName]]:
        """Return (present_sections, missing_sections)."""
        required = self._get_required_sections(article_type)
        present = self._extract_present_sections(document_content.paragraphs)
        present_lower = [s.lower() for s in present]
        missing = [s for s in required if s.lower() not in present_lower]
        return present, missing

    def _extract_present_sections(self, paragraphs: list[str]) -> list[SectionName]:
        """Detect canonical section names from a list of paragraph strings."""
        sections: list[SectionName] = []
        all_keywords = [kw for keywords in self._SECTION_ALIASES.values() for kw in keywords]

        for para in paragraphs:
            text_lower = para.lower().strip()
            is_short_header = len(text_lower) < 100
            is_inline_header = any(
                text_lower.startswith(kw + ":") or text_lower.startswith(kw + " :")
                for kw in all_keywords
            )
            if not (is_short_header or is_inline_header):
                continue
            for section_name, aliases in self._SECTION_ALIASES.items():
                if any(kw in text_lower for kw in aliases):
                    sections.append(section_name)
                    break

        return sections

    def _get_required_sections(self, article_type: ArticleType) -> list[SectionName]:
        """Delegate to RequiredSectionsProvider."""
        return RequiredSectionsProvider.get(article_type)
