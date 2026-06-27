from re import compile, split

from src.domain.document.document_text_port import DocumentTextPort
from src.domain.document.reference_extraction_port import ReferenceExtractionPort
from src.domain.dtos.reference_dto import ReferenceDTO
from src.domain.exceptions.reference_errors import ReferenceParsingFailed  # noqa: F401
from src.domain.exceptions.decorators.generic_error_handler import generic_error_handler

_DEFAULT_SECTION_TYPE = "Referencias"

_SECTION_TYPE_MAP = {
    "ibliograf": "Bibliografía",
    "fuentes": "Fuentes bibliográficas",
}

_YEAR_END_PATTERN = compile(r"\((?:\d{1,2}\s+de\s+\w+\s+de\s+)?\d{4}[a-z]?\)\.?")
_AUTHOR_CLEANUP_PATTERN = compile(
    r"([A-ZÁ-ÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁ-ÚÑ]?[a-záéíóúñ]+)*,\s+[A-ZÁÉÍÓÚÑ]\..*)"
)
_LEADING_BULLETS_PATTERN = compile(r"^[\•\-\*\·\s]+")
_BIB_SECTION_PATTERN = compile(
    r"(Bibliograf[íi]a|Referencias|Fuentes\s*bibliogr[áa]ficas(?:\s*consultadas)?)\s*(.{100,})",
    flags=2 | 16,  # re.IGNORECASE | re.DOTALL
)


class DocxReferenceAdapter(ReferenceExtractionPort):
    def __init__(self, document_text_port: DocumentTextPort) -> None:
        self._document_text_port = document_text_port

    @generic_error_handler
    def extract_references(self, docx_path: str) -> tuple[list[ReferenceDTO], str]:
        full_text = "".join(self._document_text_port.read_paragraphs(docx_path))
        bib_match = _BIB_SECTION_PATTERN.search(full_text)
        if not bib_match:
            return [], _DEFAULT_SECTION_TYPE
        section_type = self._resolve_section_type(bib_match.group(1).lower())
        return self._parse_references(bib_match.group(2)), section_type

    def _resolve_section_type(self, label: str) -> str:
        for key, value in _SECTION_TYPE_MAP.items():
            if key in label:
                return value
        return _DEFAULT_SECTION_TYPE

    def _parse_references(self, bib_text: str) -> list[ReferenceDTO]:
        parts = split(f"({_YEAR_END_PATTERN.pattern})", bib_text)
        references: list[ReferenceDTO] = []
        current_ref = ""
        for part in parts:
            current_ref += part
            if not _YEAR_END_PATTERN.fullmatch(part):
                continue
            ref = self._clean_reference(current_ref)
            current_ref = ""
            if not ref:
                continue
            references.append(ReferenceDTO(text=ref))
        ref = self._clean_reference(current_ref)
        if ref:
            references.append(ReferenceDTO(text=ref))
        return references

    def _clean_reference(self, text: str) -> str:
        clean = _LEADING_BULLETS_PATTERN.sub("", text.strip()).strip()
        author_match = _AUTHOR_CLEANUP_PATTERN.search(clean)
        if author_match:
            clean = author_match.group(1).strip()
        return clean
