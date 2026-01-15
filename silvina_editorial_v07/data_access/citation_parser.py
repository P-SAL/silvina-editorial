"""
Citation Parser - Extracts APA citations from Spanish text
"""

import re
from typing import List
from domain.models import Citation


class CitationParser:
    """Extracts APA citations from Spanish text."""
    
    # Consolidated patterns
    AUTHOR_PATTERN = r'[A-ZÁÉÍÓÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+'
    YEAR_PATTERN = r'\d{4}[a-z]?'
    PAGE_PATTERN = r'(?:pp?\.|párr\.)\s*([\d\-]+)'
    
    def __init__(self):
        # Parenthetical: (García, 2020) or (Ministerio, 2020a) or (CIA, 1985)
        self.pattern_parenthetical = re.compile(
            rf'\(([A-ZÁ-ÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+(?:\s+et\s+al\.)?(?:\s+y\s+[A-ZÁ-ÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)*),?\s+'
            rf'({self.YEAR_PATTERN})(?:,\s*{self.PAGE_PATTERN})?\)',
            re.IGNORECASE
        )
        
        # Narrative: García (2020) or Ministerio (2020a)
        self.pattern_narrative = re.compile(
            rf'([A-ZÁ-ÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+(?:\s+et\s+al\.)?(?:\s+y\s+[A-ZÁ-ÚÑ][a-záéíóúñA-ZÁÉÍÓÚÑ\-]+)*)\s+'
            rf'\(({self.YEAR_PATTERN})(?:,\s*{self.PAGE_PATTERN})?\)',
            re.IGNORECASE
        )
    
    def parse(self, text: str, paragraph_index: int) -> List[Citation]:
        """Extract all citations from one paragraph."""
        citations = []
        
        # Parenthetical citations
        for match in self.pattern_parenthetical.finditer(text):
            authors_raw = match.group(1)
            year = match.group(2)
            page = match.group(3) if match.lastindex >= 3 else None
            
            citations.append(Citation(
                authors=self._parse_authors(authors_raw),
                year=year,
                page=page,
                paragraph_index=paragraph_index,
                citation_type="parentética",
                raw_text=match.group(0)
            ))
        
        # Narrative citations
        for match in self.pattern_narrative.finditer(text):
            authors_raw = match.group(1)
            year = match.group(2)
            page = match.group(3) if match.lastindex >= 3 else None
            
            citations.append(Citation(
                authors=self._parse_authors(authors_raw),
                year=year,
                page=page,
                paragraph_index=paragraph_index,
                citation_type="narrativa",
                raw_text=match.group(0)
            ))
        
        return citations
    
    @staticmethod
    def _parse_authors(authors_text: str) -> List[str]:
        """Parse author string into list."""
        if "et al." in authors_text:
            first_author = authors_text.split("et al.")[0].strip()
            return [f"{first_author} et al."]
        
        if " y " in authors_text:
            return [a.strip() for a in authors_text.split(" y ")]
        
        return [authors_text.strip()]