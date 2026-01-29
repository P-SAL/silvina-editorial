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
   

    def extract_footnotes(self, doc) -> List[Citation]:
        """Extract Markdown-style footnote references [^1], [^2], etc."""
        citations = []
        footnote_pattern = re.compile(r'\[\^(\d+)\]')
        
        seen_numbers = set()
        
        # Search in ALL paragraphs (including footnote section)
        for i, para in enumerate(doc.paragraphs):
            matches = footnote_pattern.findall(para.text)
            for num in matches:
                if num not in seen_numbers:
                    seen_numbers.add(num)
                    citations.append(Citation(
                        text=f"[^{num}]",
                        citation_type=CitationType.FOOTNOTE,
                        location=i,
                        author=None,
                        year=None
                    ))
        
        # Also search in tables (footnotes might be there)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    matches = footnote_pattern.findall(cell.text)
                    for num in matches:
                        if num not in seen_numbers:
                            seen_numbers.add(num)
                            citations.append(Citation(
                                text=f"[^{num}]",
                                citation_type=CitationType.FOOTNOTE,
                                location=-1,
                                author=None,
                                year=None
                            ))
        
        if len(citations) > 0:
            print(f"      ✓ {len(citations)} notas al pie detectadas [^1]-[^{max(seen_numbers)}]")
            print(f"      ⚠️  FORMATO NO APA: Use citas parentéticas (Autor, Año)")
        # REMOVED: paragraph count warning (was around line 75)
        
        return citations

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
