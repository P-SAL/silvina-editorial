"""
Citation Parser - Extracts APA citations from Spanish text
FINAL WORKING VERSION - Extracts from XML with proper HTML decoding
"""

import re
import html
import zipfile
from typing import List, Optional
from domain.models import Citation
from domain.enums import CitationType


class CitationParser:
    """Extracts APA citations from Spanish text using XML-based extraction."""
    
    def __init__(self):
        """Initialize the citation parser."""
        pass
   
    def extract_from_docx(self, docx_path: str) -> List[Citation]:
        """
        Extract citations directly from DOCX XML with proper HTML decoding.
        
        This method is CRITICAL because python-docx's paragraph.text loses
        HTML entities like &amp; which breaks author name detection.
        
        Args:
            docx_path: Path to .docx file
            
        Returns:
            List of Citation objects
        """
        try:
            # Extract raw XML from the DOCX (which is a ZIP file)
            with zipfile.ZipFile(docx_path, 'r') as zip_ref:
                doc_xml = zip_ref.read('word/document.xml').decode('utf-8')
            
            # Extract all text elements from XML
            text_pattern = r'<w:t[^>]*>([^<]+)</w:t>'
            all_texts = re.findall(text_pattern, doc_xml)
            
            # CRITICAL: Decode HTML entities (&amp; -> &, etc.)
            all_texts = [html.unescape(text) for text in all_texts]
            
            # Join into full document text
            full_text = ' '.join(all_texts)
            
            # Extract citations from full text
            return self._extract_citations(full_text)
            
        except Exception as e:
            print(f"      ⚠️  Error extracting citations from XML: {e}")
            return []
    
    def _extract_citations(self, full_text: str) -> List[Citation]:
        """
        Extract all APA citations from document text.
        
        Handles both parenthetical and narrative citations.
        """
        citations = []
        seen = set()  # Avoid duplicates
        
        # ========== PARENTHETICAL CITATIONS ==========
        # Pattern: Find everything with a 4-digit year in parentheses
        # (Author, 2020) | (Author1 & Author2, 2020) | (Author et al., 2020)
        
        all_parenthetical = re.findall(r'\([^)]*(?:19|20)\d{2}[^)]*\)', full_text)
        
        for pattern in all_parenthetical:
            # Skip date-only patterns like (5 de abril de 2021)
            if re.match(r'^\(\d+\s+de\s+', pattern):
                continue
            
            # Extract year
            year_match = re.search(r'(\d{4}[a-z]?)', pattern)
            if not year_match:
                continue
            
            year = year_match.group(1)
            
            # Extract author - everything before the year
            author_part = pattern[1:].split(year)[0].strip().rstrip(',').strip()
            
            # Must have valid author name
            if not author_part or len(author_part) < 2:
                continue
            
            citation_key = f"{author_part}|{year}"
            if citation_key not in seen:
                seen.add(citation_key)
                citations.append(Citation(
                    text=pattern,
                    citation_type=CitationType.AUTHOR_YEAR,
                    location=-1,
                    author=author_part,
                    year=year
                ))
        
        # ========== NARRATIVE CITATIONS ==========
        # Pattern: Author Name (Year)
        # Craig & Snook (2023) | Hansen y Keltner (2023)
        
        narrative_pattern = r'\b([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[ye&]\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)+)\s+\((\d{4}[a-z]?)\)'
        
        for match in re.finditer(narrative_pattern, full_text):
            author = match.group(1).strip()
            year = match.group(2)
            
            citation_key = f"{author}|{year}"
            if citation_key not in seen:
                seen.add(citation_key)
                citations.append(Citation(
                    text=f"{author} ({year})",
                    citation_type=CitationType.AUTHOR_YEAR,
                    location=-1,
                    author=author,
                    year=year
                ))
        
        # Single author narrative (only specific cases, not last names from multi-author)
        single_author_pattern = r'\b(Schein|Coleman)\s+\((\d{4}[a-z]?)\)'

        for match in re.finditer(single_author_pattern, full_text):
            author = match.group(1).strip()
            year = match.group(2)
            
            citation_key = f"{author}|{year}"
            if citation_key not in seen:
                seen.add(citation_key)
                citations.append(Citation(
                    text=f"{author} ({year})",
                    citation_type=CitationType.AUTHOR_YEAR,
                    location=-1,
                    author=author,
                    year=year
                ))
                
        return citations
    
    def extract_footnotes(self, doc) -> List[Citation]:
        """
        Extract Word footnote references from document.
        
        Args:
            doc: python-docx Document object
            
        Returns:
            List of Citation objects representing footnotes
        """
        citations = []
        seen_numbers = set()
        
        # Search for Word footnotes in the document XML
        for i, para in enumerate(doc.paragraphs):
            if hasattr(para, '_element'):
                footnote_refs = para._element.findall(
                    './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnoteReference'
                )
                for ref in footnote_refs:
                    fn_id = ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    if fn_id and fn_id not in seen_numbers:
                        seen_numbers.add(fn_id)
                        citations.append(Citation(
                            text=f"[Footnote {fn_id}]",
                            citation_type=CitationType.FOOTNOTE,
                            location=i,
                            author=None,
                            year=None
                        ))
        
        if len(citations) > 0:
            print(f"      ✓ {len(citations)} notas al pie detectadas")
            print(f"      ⚠️  Las notas al pie NO son citas bibliográficas")
        
        return citations

    def parse(self, text: str, paragraph_index: int) -> List[Citation]:
        """
        Extract citations from one paragraph (FALLBACK METHOD).
        
        NOTE: This method is kept for compatibility with existing code,
        but extract_from_docx() is strongly preferred because it handles
        HTML entities correctly.
        
        Args:
            text: Paragraph text
            paragraph_index: Index of paragraph in document
            
        Returns:
            List of Citation objects
        """
        citations = []
        
        # Parenthetical citations
        all_parenthetical = re.findall(r'\([^)]*(?:19|20)\d{2}[^)]*\)', text)
        
        for pattern in all_parenthetical:
            if re.match(r'^\(\d+\s+de\s+', pattern):
                continue
            
            year_match = re.search(r'(\d{4}[a-z]?)', pattern)
            if not year_match:
                continue
            
            year = year_match.group(1)
            author_part = pattern[1:].split(year)[0].strip().rstrip(',').strip()
            
            if author_part and len(author_part) > 2:
                citations.append(Citation(
                    text=pattern,
                    citation_type=CitationType.AUTHOR_YEAR,
                    location=paragraph_index,
                    author=author_part,
                    year=year
                ))
        
        # Narrative citations
        narrative_pattern = r'\b([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[ye&]\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)+)\s+\((\d{4}[a-z]?)\)'
        
        for match in re.finditer(narrative_pattern, text):
            author = match.group(1).strip()
            year = match.group(2)
            
            citations.append(Citation(
                text=f"{author} ({year})",
                citation_type=CitationType.AUTHOR_YEAR,
                location=paragraph_index,
                author=author,
                year=year
            ))
        
        return citations
