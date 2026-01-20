"""
Citation Parser - Extracts APA citations from Spanish text
"""

import re
from typing import List
from domain.models import Citation
from domain.enums import CitationType


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
            
            citation_text = match.group(0)
            authors = self._parse_authors(authors_raw)
            author = authors[0] if authors else None

            citations.append(Citation(
                text=citation_text,
                citation_type=CitationType.AUTHOR_YEAR,
                location=paragraph_index,
                author=author,
                year=year
            ))

        
        # Narrative citations
        for match in self.pattern_narrative.finditer(text):
            authors_raw = match.group(1)
            year = match.group(2)
            page = match.group(3) if match.lastindex >= 3 else None
            
            citation_text = match.group(0)
            authors = self._parse_authors(authors_raw)
            author = authors[0] if authors else None

            citations.append(Citation(
                text=citation_text,
                citation_type=CitationType.AUTHOR_YEAR,
                location=paragraph_index,
                author=author,
                year=year
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